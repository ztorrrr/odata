"""
BigQuery service for loading and querying data
"""
import logging
from typing import Dict, List, Optional, Any

from google.cloud import bigquery
from google.cloud.exceptions import NotFound

from app.utils.setting import get_config
from app.utils.gcp_auth import get_gcp_auth

logger = logging.getLogger(__name__)


class BigQueryService:
    """BigQuery 서비스 클래스"""

    def __init__(self):
        self.config = get_config()
        self.gcp_auth = get_gcp_auth()
        self.client = None
        self.dataset_id = self.config.BIGQUERY_DATASET_ID
        self.table_name = self.config.BIGQUERY_TABLE_NAME

    def initialize(self):
        """BigQuery 클라이언트 초기화"""
        if not self.client:
            self.client = self.gcp_auth.get_bigquery_client()
            logger.info(f"BigQuery client initialized for project: {self.gcp_auth.project_id}")

    def create_dataset_if_not_exists(self):
        """데이터셋이 없으면 생성"""
        self.initialize()

        dataset_id = f"{self.gcp_auth.project_id}.{self.dataset_id}"

        try:
            dataset = self.client.get_dataset(dataset_id)
            logger.info(f"Dataset {dataset_id} already exists")
            return dataset
        except NotFound:
            dataset = bigquery.Dataset(dataset_id)
            dataset.location = "asia-northeast3"  # 서울 리전

            dataset = self.client.create_dataset(dataset, timeout=30)
            logger.info(f"Created dataset {dataset_id}")
            return dataset

    def _sanitize_column_name(self, name: str) -> str:
        """
        BigQuery 컬럼 이름 규칙에 맞게 정리합니다.
        - BOM 제거
        - 문자, 숫자, 언더스코어만 허용
        - 숫자로 시작하면 앞에 "col_" 추가

        Args:
            name: 원본 컬럼 이름

        Returns:
            정리된 컬럼 이름
        """
        import re

        # BOM 제거
        name = name.replace('\ufeff', '').replace('\ufffe', '')

        # 공백을 언더스코어로
        name = name.strip().replace(' ', '_')

        # 특수 문자를 언더스코어로 (문자, 숫자, 언더스코어만 유지)
        name = re.sub(r'[^a-zA-Z0-9_]', '_', name)

        # 연속된 언더스코어를 하나로
        name = re.sub(r'_+', '_', name)

        # 앞뒤 언더스코어 제거
        name = name.strip('_')

        # 숫자로 시작하면 앞에 col_ 추가
        if name and name[0].isdigit():
            name = f'col_{name}'

        # 빈 문자열이면 기본값
        if not name:
            name = 'unnamed_column'

        # 최대 길이 제한 (300자)
        if len(name) > 300:
            name = name[:300]

        return name

    def _get_csv_headers(self, gcs_uri: str) -> List[str]:
        """
        GCS CSV 파일의 헤더를 읽어옵니다.

        Args:
            gcs_uri: GCS URI (예: gs://bucket-name/file.csv)

        Returns:
            헤더 컬럼 리스트
        """
        # GCS URI 파싱
        parts = gcs_uri.replace("gs://", "").split("/", 1)
        bucket_name = parts[0]
        blob_name = parts[1]

        # Storage 클라이언트로 파일 읽기
        storage_client = self.gcp_auth.get_storage_client()
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(blob_name)

        # 첫 줄만 읽기 (최대 10KB)
        header_bytes = blob.download_as_bytes(start=0, end=10240)
        header_line = header_bytes.decode('utf-8-sig').split('\n')[0]  # utf-8-sig로 BOM 자동 처리

        # CSV 헤더 파싱 (간단한 쉼표 구분)
        headers = [h.strip().strip('"') for h in header_line.split(',')]

        # 컬럼 이름 정리
        sanitized_headers = []
        for i, header in enumerate(headers):
            sanitized = self._sanitize_column_name(header)
            # 중복 방지
            if sanitized in sanitized_headers:
                sanitized = f"{sanitized}_{i}"
            sanitized_headers.append(sanitized)

        return sanitized_headers

    def load_csv_from_gcs(self, gcs_uri: str = None, auto_detect: bool = True, use_string_schema: bool = False) -> bigquery.LoadJob:
        """
        GCS에서 CSV 파일을 BigQuery 테이블로 로드합니다.

        Args:
            gcs_uri: GCS URI (예: gs://bucket-name/file.csv)
                    None이면 설정에서 가져옴
            auto_detect: 스키마 자동 감지 여부
            use_string_schema: True면 모든 컬럼을 STRING으로 로드 (타입 추론 에러 방지)

        Returns:
            LoadJob 객체
        """
        self.initialize()
        self.create_dataset_if_not_exists()

        if gcs_uri is None:
            gcs_uri = f"gs://{self.config.GCS_BUCKET_NAME}/{self.config.CSV_FILE_NAME}"

        table_id = f"{self.gcp_auth.project_id}.{self.dataset_id}.{self.table_name}"

        # STRING 스키마 사용 시 헤더 읽고 스키마 생성
        schema = None
        if use_string_schema:
            logger.info("Reading CSV headers to create STRING schema...")
            headers = self._get_csv_headers(gcs_uri)
            schema = [bigquery.SchemaField(name, "STRING", mode="NULLABLE") for name in headers]
            logger.info(f"Created STRING schema with {len(schema)} columns")

        job_config = bigquery.LoadJobConfig(
            source_format=bigquery.SourceFormat.CSV,
            skip_leading_rows=1,  # 헤더 행 스킵
            autodetect=auto_detect if not use_string_schema else False,  # STRING 스키마 사용 시 자동 감지 비활성화
            schema=schema,  # 명시적 스키마 (use_string_schema=True인 경우)
            write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,  # 테이블 덮어쓰기
            max_bad_records=0 if use_string_schema else 100,  # STRING 사용 시 에러 없어야 함
            ignore_unknown_values=True,  # 알 수 없는 값 무시
            allow_jagged_rows=True,  # 컬럼 수가 다른 행 허용
            allow_quoted_newlines=True,  # 따옴표 안의 줄바꿈 허용
        )

        logger.info(f"Loading data from {gcs_uri} to {table_id}")

        load_job = self.client.load_table_from_uri(
            gcs_uri, table_id, job_config=job_config
        )

        # 작업 완료 대기
        load_job.result()

        # 로드된 테이블 정보 가져오기
        table = self.client.get_table(table_id)
        logger.info(f"Loaded {table.num_rows} rows to {table_id}")

        return load_job

    def get_table_schema(self) -> List[bigquery.SchemaField]:
        """테이블 스키마 가져오기"""
        self.initialize()

        table_id = f"{self.gcp_auth.project_id}.{self.dataset_id}.{self.table_name}"

        try:
            table = self.client.get_table(table_id)
            return table.schema
        except NotFound:
            logger.error(f"Table {table_id} not found")
            return []

    def query_table(
        self,
        select: Optional[List[str]] = None,
        filter: Optional[str] = None,
        orderby: Optional[str] = None,
        top: Optional[int] = None,
        skip: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        """
        테이블 쿼리 실행 (OData 파라미터 지원)

        Args:
            select: 선택할 컬럼 리스트
            filter: WHERE 조건
            orderby: ORDER BY 조건
            top: LIMIT 값
            skip: OFFSET 값

        Returns:
            쿼리 결과 리스트
        """
        self.initialize()

        table_id = f"{self.gcp_auth.project_id}.{self.dataset_id}.{self.table_name}"

        # SELECT 절 구성
        if select:
            select_clause = ", ".join(f"`{col}`" for col in select)
        else:
            select_clause = "*"

        # 기본 쿼리
        query = f"SELECT {select_clause} FROM `{table_id}`"

        # WHERE 절
        if filter:
            query += f" WHERE {filter}"

        # ORDER BY 절
        if orderby:
            query += f" ORDER BY {orderby}"

        # LIMIT/OFFSET
        if top:
            query += f" LIMIT {top}"
        if skip:
            query += f" OFFSET {skip}"

        logger.debug(f"Executing query: {query}")

        # 쿼리 실행
        query_job = self.client.query(query)
        results = query_job.result()

        # 결과를 딕셔너리 리스트로 변환
        rows = []
        for row in results:
            rows.append(dict(row.items()))

        return rows

    def get_row_count(self, filter: Optional[str] = None) -> int:
        """
        테이블 행 수 가져오기

        Args:
            filter: WHERE 조건

        Returns:
            행 수
        """
        self.initialize()

        table_id = f"{self.gcp_auth.project_id}.{self.dataset_id}.{self.table_name}"

        query = f"SELECT COUNT(*) as count FROM `{table_id}`"

        if filter:
            query += f" WHERE {filter}"

        query_job = self.client.query(query)
        results = query_job.result()

        for row in results:
            return row.count

        return 0

    def get_table_info(self) -> Dict[str, Any]:
        """테이블 정보 가져오기"""
        self.initialize()

        table_id = f"{self.gcp_auth.project_id}.{self.dataset_id}.{self.table_name}"

        try:
            table = self.client.get_table(table_id)
            return {
                "table_id": table_id,
                "num_rows": table.num_rows,
                "num_bytes": table.num_bytes,
                "created": table.created,
                "modified": table.modified,
                "schema": [
                    {
                        "name": field.name,
                        "type": field.field_type,
                        "mode": field.mode,
                    }
                    for field in table.schema
                ],
            }
        except NotFound:
            return None


# 싱글톤 인스턴스
_bigquery_service: Optional[BigQueryService] = None


def get_bigquery_service() -> BigQueryService:
    """BigQuery 서비스 싱글톤 인스턴스 반환"""
    global _bigquery_service
    if _bigquery_service is None:
        _bigquery_service = BigQueryService()
    return _bigquery_service