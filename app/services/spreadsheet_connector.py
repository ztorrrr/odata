"""
Google Spreadsheet BigQuery connector service
"""
import logging
from typing import Dict, Optional, Any, List
from google.cloud import bigquery
from google.cloud.exceptions import NotFound
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
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
        self.sheets_service = None
        self.drive_service = None
        self.script_service = None

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

    def _get_sheets_service(self):
        """Google Sheets API 서비스 클라이언트 반환"""
        if not self.sheets_service:
            self.initialize()
            self.sheets_service = build('sheets', 'v4', credentials=self.gcp_auth.credentials)
        return self.sheets_service

    def _get_drive_service(self):
        """Google Drive API 서비스 클라이언트 반환"""
        if not self.drive_service:
            self.initialize()
            self.drive_service = build('drive', 'v3', credentials=self.gcp_auth.credentials)
        return self.drive_service

    def _get_script_service(self):
        """Google Apps Script API 서비스 클라이언트 반환"""
        if not self.script_service:
            self.initialize()
            self.script_service = build('script', 'v1', credentials=self.gcp_auth.credentials)
        return self.script_service

    def _find_folder_by_name(self, folder_name: str) -> Optional[str]:
        """
        폴더 이름으로 Google Drive 폴더 ID 검색

        Args:
            folder_name: 검색할 폴더 이름

        Returns:
            폴더 ID (찾지 못하면 None)
        """
        drive_service = self._get_drive_service()

        # 루트 폴더에서 검색
        query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"

        try:
            results = drive_service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)',
                pageSize=10
            ).execute()

            files = results.get('files', [])

            if files:
                folder_id = files[0]['id']
                logger.info(f"Found folder '{folder_name}': {folder_id}")
                return folder_id
            else:
                logger.warning(f"Folder '{folder_name}' not found")
                return None

        except HttpError as e:
            logger.error(f"Error searching for folder: {e}")
            return None

    def _move_to_folder(self, spreadsheet_id: str, folder_id: str):
        """스프레드시트를 특정 폴더로 이동"""
        drive_service = self._get_drive_service()

        # 현재 부모 폴더 가져오기
        file = drive_service.files().get(
            fileId=spreadsheet_id,
            fields='parents'
        ).execute()

        previous_parents = ",".join(file.get('parents', []))

        # 새 폴더로 이동
        drive_service.files().update(
            fileId=spreadsheet_id,
            addParents=folder_id,
            removeParents=previous_parents,
            fields='id, parents'
        ).execute()

    def create_spreadsheet_with_connected_bigquery(
        self,
        spreadsheet_title: str = None,
        view_id: str = None,
        folder_id: str = None,
        folder_name: str = "odata_test"
    ) -> Dict[str, Any]:
        """
        BigQuery Connected Sheets를 생성 (네이티브 연결)

        이 방식은 Apps Script 없이도 BigQuery 데이터를 직접 조회하고
        새로고침할 수 있는 네이티브 연결을 생성합니다.

        Args:
            spreadsheet_title: 생성할 스프레드시트 이름 (None이면 자동 생성: bigquery_connector_YYMMDD_HHMMSS)
            view_id: BigQuery View ID (None이면 기본 샘플 view 사용)
            folder_id: Google Drive 폴더 ID (None이면 folder_name으로 검색)
            folder_name: 폴더 이름 (folder_id가 None일 때 검색, 기본값: "odata_test")

        Returns:
            생성된 스프레드시트 정보
        """
        try:
            self.initialize()

            # 스프레드시트 제목 자동 생성
            if spreadsheet_title is None:
                from datetime import datetime
                timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
                spreadsheet_title = f"bigquery_connector_{timestamp}"
                logger.info(f"Auto-generated spreadsheet title: {spreadsheet_title}")

            # 폴더 ID 자동 검색
            if folder_id is None and folder_name:
                folder_id = self._find_folder_by_name(folder_name)
                if folder_id:
                    logger.info(f"Using folder '{folder_name}' (ID: {folder_id})")
                else:
                    logger.warning(f"Folder '{folder_name}' not found, creating in root")

            # View ID 설정
            if view_id is None:
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
            table_id = parts[2] if len(parts) > 2 else view_id

            logger.info(f"Creating Connected Sheets '{spreadsheet_title}' with BigQuery view: {view_id}")

            # 1. 빈 스프레드시트 생성
            sheets_service = self._get_sheets_service()
            spreadsheet_body = {
                'properties': {
                    'title': spreadsheet_title
                }
            }

            spreadsheet = sheets_service.spreadsheets().create(
                body=spreadsheet_body
            ).execute()

            spreadsheet_id = spreadsheet['spreadsheetId']
            spreadsheet_url = spreadsheet['spreadsheetUrl']

            logger.info(f"Created spreadsheet: {spreadsheet_url}")

            # 2. 폴더로 이동 (folder_id가 제공된 경우)
            if folder_id:
                self._move_to_folder(spreadsheet_id, folder_id)
                logger.info(f"Moved spreadsheet to folder: {folder_id}")

            # 3. BigQuery Data Source 추가
            data_source_response = self._create_bigquery_data_source(
                spreadsheet_id,
                project_id,
                dataset_id,
                table_id
            )

            data_source_id = data_source_response['dataSource']['dataSourceId']
            logger.info(f"Created BigQuery data source: {data_source_id}")

            # 4. 기본 빈 시트 제거 및 데이터 소스 시트를 첫 번째로 이동
            # (데이터 소스 실행 완료를 기다리지 않고 바로 cleanup)
            self._cleanup_default_sheet(spreadsheet_id)
            logger.info(f"Cleaned up default sheet")

            # 5. 데이터 소스 실행 완료 대기 (백그라운드)
            # 참고: 데이터 소스는 백그라운드에서 실행되며,
            # 스프레드시트를 열면 자동으로 데이터가 로드됩니다.
            execution_status = "RUNNING"
            logger.info(f"Data source execution started in background")

            return {
                "success": True,
                "status": "CREATED",
                "spreadsheet": {
                    "id": spreadsheet_id,
                    "url": spreadsheet_url,
                    "title": spreadsheet_title
                },
                "bigquery": {
                    "view_id": view_id,
                    "data_source_id": data_source_id,
                    "project_id": project_id,
                    "dataset_id": dataset_id,
                    "table_id": table_id
                },
                "folder": {
                    "id": folder_id,
                    "name": folder_name if folder_id else None
                },
                "connected_sheets": True,
                "message": f"✓ SUCCESS: BigQuery Connected Sheets created successfully!",
                "instructions": [
                    f"1. Open spreadsheet: {spreadsheet_url}",
                    "2. Data is connected to BigQuery and ready to use",
                    "3. Refresh data: Data > Data sources > Refresh",
                    "4. The spreadsheet will automatically load the latest BigQuery data"
                ]
            }

        except HttpError as e:
            logger.error(f"Google API error: {e}")
            raise Exception(f"Failed to create Connected Sheets: {e}")
        except Exception as e:
            logger.error(f"Error creating Connected Sheets: {e}")
            raise

    def _create_bigquery_data_source(
        self,
        spreadsheet_id: str,
        project_id: str,
        dataset_id: str,
        table_id: str
    ) -> Dict[str, Any]:
        """
        BigQuery 데이터 소스를 스프레드시트에 추가

        Args:
            spreadsheet_id: 스프레드시트 ID
            project_id: BigQuery 프로젝트 ID
            dataset_id: BigQuery 데이터셋 ID
            table_id: BigQuery 테이블/뷰 ID

        Returns:
            AddDataSourceResponse
        """
        sheets_service = self._get_sheets_service()

        requests = [{
            'addDataSource': {
                'dataSource': {
                    'spec': {
                        'bigQuery': {
                            'projectId': project_id,
                            'tableSpec': {
                                'tableProjectId': project_id,
                                'datasetId': dataset_id,
                                'tableId': table_id
                            }
                        }
                    }
                }
            }
        }]

        body = {
            'requests': requests
        }

        response = sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

        # AddDataSourceResponse 반환
        return response['replies'][0]['addDataSource']

    def _wait_for_data_source(
        self,
        spreadsheet_id: str,
        data_source_id: str,
        max_wait_seconds: int = 60
    ) -> str:
        """
        데이터 소스 실행 완료 대기 (polling)

        Args:
            spreadsheet_id: 스프레드시트 ID
            data_source_id: 데이터 소스 ID
            max_wait_seconds: 최대 대기 시간 (초)

        Returns:
            최종 실행 상태 ('SUCCEEDED' or 'FAILED')
        """
        import time

        sheets_service = self._get_sheets_service()
        start_time = time.time()

        while True:
            # 스프레드시트 정보 가져오기
            spreadsheet = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                includeGridData=False
            ).execute()

            # 데이터 소스 찾기
            data_sources = spreadsheet.get('dataSources', [])
            for ds in data_sources:
                if ds.get('dataSourceId') == data_source_id:
                    status = ds.get('dataExecution Status', {}).get('state', 'UNKNOWN')

                    if status in ['SUCCEEDED', 'FAILED']:
                        return status

            # 타임아웃 체크
            if time.time() - start_time > max_wait_seconds:
                logger.warning(f"Data source execution timeout after {max_wait_seconds}s")
                return 'TIMEOUT'

            # 잠시 대기 후 재시도
            time.sleep(2)

    def _cleanup_default_sheet(self, spreadsheet_id: str):
        """
        기본 빈 시트를 제거하고 데이터 소스 시트만 남김

        Args:
            spreadsheet_id: 스프레드시트 ID
        """
        sheets_service = self._get_sheets_service()

        # 스프레드시트 정보 가져오기
        spreadsheet = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
        ).execute()

        sheets = spreadsheet.get('sheets', [])

        # 데이터 소스가 연결되지 않은 빈 시트 찾기
        default_sheets = []
        data_source_sheets = []

        for sheet in sheets:
            sheet_id = sheet['properties']['sheetId']
            sheet_title = sheet['properties']['title']

            # 데이터 소스가 연결된 시트인지 확인
            has_data_source = 'dataSourceSheetProperties' in sheet.get('properties', {})

            if has_data_source:
                data_source_sheets.append((sheet_id, sheet_title))
            else:
                default_sheets.append((sheet_id, sheet_title))

        # 데이터 소스 시트가 있고, 기본 시트가 있으면 기본 시트 삭제
        if data_source_sheets and default_sheets:
            requests = []

            # 기본 시트 삭제 요청
            for sheet_id, sheet_title in default_sheets:
                requests.append({
                    'deleteSheet': {
                        'sheetId': sheet_id
                    }
                })
                logger.info(f"Deleting default sheet: {sheet_title} (ID: {sheet_id})")

            # 데이터 소스 시트를 첫 번째로 이동
            if data_source_sheets:
                first_data_source_id = data_source_sheets[0][0]
                requests.append({
                    'updateSheetProperties': {
                        'properties': {
                            'sheetId': first_data_source_id,
                            'index': 0
                        },
                        'fields': 'index'
                    }
                })
                logger.info(f"Moving data source sheet to first position")

            # 요청 실행
            if requests:
                body = {'requests': requests}
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=body
                ).execute()


# 싱글톤 인스턴스
_connector: Optional[SpreadsheetConnector] = None


def get_spreadsheet_connector() -> SpreadsheetConnector:
    """SpreadsheetConnector 싱글톤 인스턴스 반환"""
    global _connector
    if _connector is None:
        _connector = SpreadsheetConnector()
    return _connector