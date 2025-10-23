import os
from dataclasses import dataclass
from functools import cache

from app.utils import aws_secret_manager as secret_manager

SERVICE = "gen-ai"  # 서비스 이름

@dataclass
class BaseConfig:
    ENVIRONMENT: str = os.getenv("ENVIRONMENT", "DEV")

    # OData Configuration
    HOST: str = os.getenv("HOST", "0.0.0.0")
    PORT: int = int(os.getenv("PORT", "8888"))

    # GCS Configuration
    GCS_BUCKET_NAME: str = os.getenv("GCS_BUCKET_NAME", "test-core-odata")
    CSV_FILE_NAME: str = os.getenv("CSV_FILE_NAME", "musinsa_sa_test_250909.csv")

    # BigQuery Configuration
    BIGQUERY_DATASET_ID: str = os.getenv("BIGQUERY_DATASET_ID", "odata_dataset")
    BIGQUERY_TABLE_NAME: str = os.getenv("BIGQUERY_TABLE_NAME", "musinsa_data")

    # AWS Secret Manager Keys
    # GCP 서비스 계정 키가 저장된 Secret Manager 키 dev/gen-ai/google/auth

    GCP_SERVICE_ACCOUNT_KEY: str = f"{ENVIRONMENT.lower()}/{SERVICE}/google/auth"

    # OData Configuration
    ODATA_SERVICE_NAME: str = "ODataService"
    ODATA_SERVICE_VERSION: str = "1.0"
    ODATA_MAX_PAGE_SIZE: int = 1000

    # Windows Excel Service (Optional - for COM-based Excel generation)
    WINDOWS_EXCEL_SERVICE_URL: str = os.getenv("WINDOWS_EXCEL_SERVICE_URL", "")
    WINDOWS_EXCEL_SERVICE_TIMEOUT: int = int(os.getenv("WINDOWS_EXCEL_SERVICE_TIMEOUT", "60"))


@dataclass
class ProductionConfig(BaseConfig):
    DOCS_URL: str = None
    REDOC_URL: str = None
    LOG_LEVEL: str = "INFO"


@dataclass
class DevelopmentConfig(BaseConfig):
    DOCS_URL: str = "/docs"
    REDOC_URL: str = "/redoc"
    LOG_LEVEL: str = "DEBUG"


@dataclass
class TestConfig(BaseConfig):
    DOCS_URL: str = "/docs"
    REDOC_URL: str = "/redoc"
    LOG_LEVEL: str = "DEBUG"


@cache
def get_config():
    env = BaseConfig().ENVIRONMENT.upper()
    if env == "PROD":
        return ProductionConfig()
    elif env == "DEV":
        return DevelopmentConfig()
    else:
        return TestConfig()


@cache
def get_gcp_service_account_key():
    """
    AWS Secret Manager에서 GCP 서비스 계정 키를 가져옵니다.
    """
    return secret_manager.get_secret(get_config().GCP_SERVICE_ACCOUNT_KEY)


def setup_gcp_auth():
    """
    GCP 인증을 설정합니다.
    AWS Secret Manager에서 서비스 계정 키를 가져와 설정합니다.
    """
    from app.utils.gcp_auth import get_gcp_auth

    gcp_auth = get_gcp_auth()
    gcp_auth.authenticate_from_secret(get_config().GCP_SERVICE_ACCOUNT_KEY)

    return gcp_auth
