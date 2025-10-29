"""
GCP Authentication utility module
AWS Secret Manager에서 GCP 서비스 계정 키를 가져와서 인증합니다.
"""
import json
import os
import tempfile
from typing import Optional

from google.oauth2 import service_account
from google.cloud import bigquery, storage
from google.auth import default
from google.auth.exceptions import DefaultCredentialsError

from app.utils import aws_secret_manager as secret_manager


class GCPAuth:
    """GCP 인증 관리 클래스"""

    def __init__(self):
        self.credentials = None
        self.project_id = None
        self.service_account_info = None
        self._initialized = False

    def authenticate_from_secret(self, secret_key: str) -> None:
        """
        AWS Secret Manager에서 GCP 서비스 계정 키를 가져와 인증합니다.

        Args:
            secret_key: AWS Secret Manager의 시크릿 키 이름
        """
        # AWS Secret Manager에서 GCP 서비스 계정 키 가져오기
        gcp_key_data = secret_manager.get_secret(secret_key)

        # 서비스 계정 키가 JSON 문자열로 저장된 경우 처리
        if isinstance(gcp_key_data, dict):
            if "service_account_key" in gcp_key_data:
                # 키가 service_account_key 필드에 문자열로 저장된 경우
                service_account_info = json.loads(gcp_key_data["service_account_key"])
            else:
                # 키 자체가 딕셔너리로 저장된 경우
                service_account_info = gcp_key_data
        else:
            # 문자열로 저장된 경우
            service_account_info = json.loads(gcp_key_data)

        self.service_account_info = service_account_info
        self.project_id = service_account_info.get("project_id")

        # 서비스 계정 크레덴셜 생성
        self.credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=[
                "https://www.googleapis.com/auth/cloud-platform",
                "https://www.googleapis.com/auth/bigquery",
                "https://www.googleapis.com/auth/devstorage.read_only",
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
                "https://www.googleapis.com/auth/script.projects",
            ]
        )

        # 환경변수에 프로젝트 ID 설정
        if self.project_id:
            os.environ["GOOGLE_CLOUD_PROJECT"] = self.project_id

    def authenticate_with_temp_file(self, secret_key: str) -> str:
        """
        임시 파일을 생성하여 GOOGLE_APPLICATION_CREDENTIALS 환경변수를 설정합니다.
        (일부 라이브러리가 파일 기반 인증만 지원하는 경우 사용)

        Args:
            secret_key: AWS Secret Manager의 시크릿 키 이름

        Returns:
            임시 파일 경로
        """
        # AWS Secret Manager에서 GCP 서비스 계정 키 가져오기
        gcp_key_data = secret_manager.get_secret(secret_key)

        if isinstance(gcp_key_data, dict):
            if "service_account_key" in gcp_key_data:
                service_account_key = gcp_key_data["service_account_key"]
            else:
                service_account_key = json.dumps(gcp_key_data)
        else:
            service_account_key = gcp_key_data

        # 임시 파일 생성
        temp_file = tempfile.NamedTemporaryFile(
            mode='w',
            suffix='.json',
            delete=False
        )
        temp_file.write(service_account_key)
        temp_file.close()

        # 환경변수 설정
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = temp_file.name

        # 프로젝트 ID 설정
        key_dict = json.loads(service_account_key)
        if "project_id" in key_dict:
            os.environ["GOOGLE_CLOUD_PROJECT"] = key_dict["project_id"]
            self.project_id = key_dict["project_id"]

        return temp_file.name

    def get_bigquery_client(self) -> bigquery.Client:
        """
        BigQuery 클라이언트를 반환합니다.
        """
        if self.credentials:
            return bigquery.Client(
                credentials=self.credentials,
                project=self.project_id
            )
        else:
            # 환경변수 기반 인증 사용
            return bigquery.Client()

    def get_storage_client(self) -> storage.Client:
        """
        Cloud Storage 클라이언트를 반환합니다.
        """
        if self.credentials:
            return storage.Client(
                credentials=self.credentials,
                project=self.project_id
            )
        else:
            # 환경변수 기반 인증 사용
            return storage.Client()

    def authenticate_with_adc(self) -> bool:
        """
        Application Default Credentials (ADC)로 인증을 시도합니다.

        Returns:
            인증 성공 여부
        """
        import logging
        logger = logging.getLogger(__name__)

        try:
            credentials, project = default(scopes=[
                "https://www.googleapis.com/auth/cloud-platform",
                "https://www.googleapis.com/auth/bigquery",
                "https://www.googleapis.com/auth/devstorage.read_only",
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
                "https://www.googleapis.com/auth/script.projects",
            ])

            self.credentials = credentials
            self.project_id = project or os.environ.get("GOOGLE_CLOUD_PROJECT")
            self._initialized = True

            logger.info(f"ADC authentication successful. Project: {self.project_id}")
            return True
        except DefaultCredentialsError as e:
            logger.warning(f"ADC not available: {e}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error during ADC authentication: {e}", exc_info=True)
            return False


# 싱글톤 인스턴스
_gcp_auth: Optional[GCPAuth] = None


def get_gcp_auth() -> GCPAuth:
    """
    GCP 인증 싱글톤 인스턴스를 반환합니다.
    """
    global _gcp_auth
    if _gcp_auth is None:
        _gcp_auth = GCPAuth()
    return _gcp_auth