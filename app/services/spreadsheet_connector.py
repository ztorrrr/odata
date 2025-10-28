"""
Google Spreadsheet BigQuery connector service
"""
import logging
from typing import Dict, Optional, Any, List
from google.cloud import bigquery
from google.cloud.exceptions import NotFound
from app.utils.gcp_auth import get_gcp_auth
from app.utils.setting import get_config
from app.services.bigquery_service import get_bigquery_service

logger = logging.getLogger(__name__)


class SpreadsheetConnector:
    """Google Spreadsheet와 BigQuery 연동 서비스"""

    def __init__(self):
        self.config = get_config()
        self.gcp_auth = get_gcp_auth()
        self.bq_service = get_bigquery_service()
        self.client = None

    def initialize(self):
        """BigQuery 클라이언트 초기화"""
        if not self.client:
            self.client = self.gcp_auth.get_bigquery_client()
            self.bq_service.initialize()
            logger.info("SpreadsheetConnector initialized")

    def create_sample_view(
        self,
        source_table: str = None,
        view_name: str = None,
        sample_size: int = 100,
        force_recreate: bool = False
    ) -> str:
        """
        BigQuery 테이블에서 샘플 데이터 View 생성

        Args:
            source_table: 원본 테이블 이름 (None이면 설정에서 가져옴)
            view_name: View 이름 (None이면 자동 생성)
            sample_size: 샘플 행 수 (기본 100)
            force_recreate: 기존 View가 있어도 재생성 여부

        Returns:
            생성된 View의 전체 ID
        """
        self.initialize()

        # 원본 테이블 ID
        if source_table is None:
            source_table = self.config.BIGQUERY_TABLE_NAME

        source_table_id = (
            f"{self.gcp_auth.project_id}."
            f"{self.config.BIGQUERY_DATASET_ID}."
            f"{source_table}"
        )

        # View 이름 자동 생성
        if view_name is None:
            view_name = f"{source_table}_sample_{sample_size}"

        view_id = (
            f"{self.gcp_auth.project_id}."
            f"{self.config.BIGQUERY_DATASET_ID}."
            f"{view_name}"
        )

        # 기존 View 확인
        try:
            view = self.client.get_table(view_id)
            if not force_recreate:
                logger.info(f"View {view_id} already exists")
                return view_id
            else:
                # 기존 View 삭제
                self.client.delete_table(view_id)
                logger.info(f"Deleted existing view: {view_id}")
        except NotFound:
            pass

        # View 생성 SQL
        view_query = f"""
        SELECT *
        FROM `{source_table_id}`
        LIMIT {sample_size}
        """

        # View 생성
        view = bigquery.Table(view_id)
        view.view_query = view_query

        view = self.client.create_table(view)
        logger.info(f"Created view: {view_id} with {sample_size} sample rows")

        return view_id

    def get_connected_sheets_config(
        self,
        view_id: str = None,
        spreadsheet_id: str = None
    ) -> Dict[str, Any]:
        """
        Connected Sheets 연결에 필요한 설정 정보 반환

        Args:
            view_id: BigQuery View ID
            spreadsheet_id: Google Spreadsheet ID

        Returns:
            Connected Sheets 설정 정보
        """
        self.initialize()

        if view_id is None:
            # 기본 샘플 view 이름 사용
            view_name = f"{self.config.BIGQUERY_TABLE_NAME}_sample_100"
            view_id = (
                f"{self.gcp_auth.project_id}."
                f"{self.config.BIGQUERY_DATASET_ID}."
                f"{view_name}"
            )

        # View ID 파싱
        parts = view_id.split('.')
        project_id = parts[0] if len(parts) > 0 else self.gcp_auth.project_id
        dataset_id = parts[1] if len(parts) > 1 else self.config.BIGQUERY_DATASET_ID
        table_name = parts[2] if len(parts) > 2 else view_id

        config = {
            "bigquery": {
                "projectId": project_id,
                "datasetId": dataset_id,
                "tableId": table_name,
                "viewId": view_id
            },
            "connection_info": {
                "type": "BigQuery Connected Sheet",
                "refresh": "On-demand or scheduled",
                "authentication": "Service Account (automated) or User OAuth (manual)"
            }
        }

        if spreadsheet_id:
            config["spreadsheet"] = {
                "id": spreadsheet_id,
                "url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
            }

        # Connected Sheets를 위한 SQL 쿼리 (옵션)
        config["query"] = f"SELECT * FROM `{view_id}`"

        return config

    def create_data_source_for_sheet(
        self,
        spreadsheet_id: str,
        view_id: str = None
    ) -> Dict[str, Any]:
        """
        스프레드시트에 BigQuery 데이터 소스 설정 가이드 제공

        Note: Google Sheets API로는 Connected Sheets를 직접 생성할 수 없으므로
        사용자가 UI에서 설정할 수 있도록 가이드 제공

        Args:
            spreadsheet_id: Google Spreadsheet ID
            view_id: BigQuery View ID

        Returns:
            설정 가이드 및 정보
        """
        config = self.get_connected_sheets_config(view_id, spreadsheet_id)

        guide = {
            "manual_setup_steps": [
                "1. Open the spreadsheet in Google Sheets",
                "2. Go to Data > Data connectors > Connect to BigQuery",
                "3. Select your Google Cloud project",
                f"4. Browse to dataset: {config['bigquery']['datasetId']}",
                f"5. Select table/view: {config['bigquery']['tableId']}",
                "6. Click 'Connect' to establish the connection",
                "7. Choose columns to import (or select all)",
                "8. Set refresh schedule if needed"
            ],
            "connection_config": config,
            "spreadsheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}",
            "alternative_method": {
                "description": "Use BigQuery SQL query in Sheets",
                "steps": [
                    "1. In Google Sheets, go to Data > Data connectors > Connect to BigQuery",
                    "2. Choose 'Custom query' option",
                    f"3. Paste this query: {config['query']}",
                    "4. Click 'Connect'"
                ]
            }
        }

        return guide

    def get_sample_data(
        self,
        view_id: str = None,
        limit: int = 10
    ) -> List[Dict[str, Any]]:
        """
        샘플 View에서 데이터 미리보기

        Args:
            view_id: View ID
            limit: 미리보기 행 수

        Returns:
            샘플 데이터
        """
        self.initialize()

        if view_id is None:
            view_name = f"{self.config.BIGQUERY_TABLE_NAME}_sample_100"
            view_id = (
                f"{self.gcp_auth.project_id}."
                f"{self.config.BIGQUERY_DATASET_ID}."
                f"{view_name}"
            )

        query = f"SELECT * FROM `{view_id}` LIMIT {limit}"

        query_job = self.client.query(query)
        results = query_job.result()

        rows = []
        for row in results:
            rows.append(dict(row.items()))

        return rows

    def modify_view_with_test_suffix(
        self,
        view_name: str = None,
        column_to_modify: str = "Type",
        suffix: str = "_테스트"
    ) -> str:
        """
        View를 수정하여 특정 컬럼에 suffix 추가
        실시간 동기화 테스트용

        Args:
            view_name: View 이름 (None이면 기본 샘플 view)
            column_to_modify: 수정할 컬럼명
            suffix: 추가할 문자열

        Returns:
            수정된 View ID
        """
        self.initialize()

        if view_name is None:
            view_name = f"{self.config.BIGQUERY_TABLE_NAME}_sample_100"

        view_id = (
            f"{self.gcp_auth.project_id}."
            f"{self.config.BIGQUERY_DATASET_ID}."
            f"{view_name}"
        )

        # 원본 테이블 ID
        source_table_id = (
            f"{self.gcp_auth.project_id}."
            f"{self.config.BIGQUERY_DATASET_ID}."
            f"{self.config.BIGQUERY_TABLE_NAME}"
        )

        # 수정된 View SQL - Type 컬럼에 suffix 추가
        # 다른 컬럼은 그대로, Type 컬럼만 CONCAT으로 수정
        view_query = f"""
        SELECT
            * EXCEPT({column_to_modify}),
            CONCAT({column_to_modify}, '{suffix}') AS {column_to_modify}
        FROM `{source_table_id}`
        LIMIT 100
        """

        # View 업데이트 (CREATE OR REPLACE VIEW)
        update_query = f"""
        CREATE OR REPLACE VIEW `{view_id}` AS
        {view_query}
        """

        try:
            # View 업데이트 실행
            query_job = self.client.query(update_query)
            query_job.result()  # 작업 완료 대기

            logger.info(f"View {view_id} modified successfully with suffix '{suffix}' on column '{column_to_modify}'")
            return view_id

        except Exception as e:
            logger.error(f"Failed to modify view: {e}")
            raise

    def restore_original_view(
        self,
        view_name: str = None,
        sample_size: int = 100
    ) -> str:
        """
        View를 원래 상태로 복원

        Args:
            view_name: View 이름 (None이면 기본 샘플 view)
            sample_size: 샘플 행 수

        Returns:
            복원된 View ID
        """
        self.initialize()

        if view_name is None:
            view_name = f"{self.config.BIGQUERY_TABLE_NAME}_sample_100"

        view_id = (
            f"{self.gcp_auth.project_id}."
            f"{self.config.BIGQUERY_DATASET_ID}."
            f"{view_name}"
        )

        # 원본 테이블 ID
        source_table_id = (
            f"{self.gcp_auth.project_id}."
            f"{self.config.BIGQUERY_DATASET_ID}."
            f"{self.config.BIGQUERY_TABLE_NAME}"
        )

        # 원본 View SQL (수정 없이)
        view_query = f"""
        SELECT *
        FROM `{source_table_id}`
        LIMIT {sample_size}
        """

        # View 업데이트 (CREATE OR REPLACE VIEW)
        update_query = f"""
        CREATE OR REPLACE VIEW `{view_id}` AS
        {view_query}
        """

        try:
            # View 업데이트 실행
            query_job = self.client.query(update_query)
            query_job.result()  # 작업 완료 대기

            logger.info(f"View {view_id} restored to original state")
            return view_id

        except Exception as e:
            logger.error(f"Failed to restore view: {e}")
            raise


# 싱글톤 인스턴스
_connector: Optional[SpreadsheetConnector] = None


def get_spreadsheet_connector() -> SpreadsheetConnector:
    """SpreadsheetConnector 싱글톤 인스턴스 반환"""
    global _connector
    if _connector is None:
        _connector = SpreadsheetConnector()
    return _connector