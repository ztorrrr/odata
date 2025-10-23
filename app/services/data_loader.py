#!/usr/bin/env python
"""
GCS에서 BigQuery로 CSV 데이터를 로드하는 스크립트
"""
import logging
import sys
from pathlib import Path

# 프로젝트 루트를 Python 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from app.utils.setting import get_config, setup_gcp_auth
from app.services.bigquery_service import get_bigquery_service

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def load_data_from_gcs():
    """GCS에서 BigQuery로 데이터 로드"""
    try:
        # 설정 로드
        config = get_config()
        logger.info(f"Environment: {config.ENVIRONMENT}")
        logger.info(f"GCS Bucket: {config.GCS_BUCKET_NAME}")
        logger.info(f"CSV File: {config.CSV_FILE_NAME}")
        logger.info(f"BigQuery Dataset: {config.BIGQUERY_DATASET_ID}")
        logger.info(f"BigQuery Table: {config.BIGQUERY_TABLE_NAME}")

        # GCP 인증 설정
        logger.info("Setting up GCP authentication...")
        gcp_auth = setup_gcp_auth()
        logger.info(f"Authenticated with project: {gcp_auth.project_id}")

        # BigQuery 서비스 초기화
        logger.info("Initializing BigQuery service...")
        bq_service = get_bigquery_service()

        # CSV 데이터 로드 (모든 컬럼을 STRING으로 로드하여 타입 에러 방지)
        logger.info("Loading CSV data to BigQuery (all columns as STRING)...")
        gcs_uri = f"gs://{config.GCS_BUCKET_NAME}/{config.CSV_FILE_NAME}"
        load_job = bq_service.load_csv_from_gcs(gcs_uri, use_string_schema=True)

        logger.info(f"Load job completed successfully. Job ID: {load_job.job_id}")

        # 테이블 정보 출력
        table_info = bq_service.get_table_info()
        if table_info:
            logger.info("Table information:")
            logger.info(f"  - Table ID: {table_info['table_id']}")
            logger.info(f"  - Rows: {table_info['num_rows']:,}")
            logger.info(f"  - Size: {table_info['num_bytes']:,} bytes")
            logger.info(f"  - Schema:")
            for field in table_info['schema']:
                logger.info(f"    - {field['name']}: {field['type']}")

        # 샘플 데이터 쿼리
        logger.info("\nQuerying sample data...")
        sample_rows = bq_service.query_table(top=5)
        logger.info(f"Sample rows (first 5):")
        for i, row in enumerate(sample_rows, 1):
            logger.info(f"  Row {i}: {row}")

        logger.info("\nData loading completed successfully!")
        return True

    except Exception as e:
        logger.error(f"Error during data loading: {str(e)}", exc_info=True)
        return False


def main():
    """메인 함수"""
    success = load_data_from_gcs()
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
