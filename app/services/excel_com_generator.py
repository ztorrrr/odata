"""
Excel COM Generator Service
Windows COM을 사용하여 Power Query OData 연결이 포함된 Excel 파일 생성
"""
import logging
import os
import tempfile
import time
from typing import Optional
from pathlib import Path
import pythoncom
import win32com.client

logger = logging.getLogger(__name__)


class ExcelCOMGenerator:
    """
    Windows COM을 사용하여 Excel 파일을 직접 생성
    Power Query를 통한 OData 연결 설정
    """

    def __init__(self):
        """초기화"""
        self.excel = None
        self.workbook = None

    def __enter__(self):
        """Context manager 진입"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager 종료 - 리소스 정리"""
        self.cleanup()

    def cleanup(self):
        """Excel COM 객체 정리"""
        try:
            if self.workbook:
                try:
                    self.workbook.Close(False)  # 저장하지 않고 닫기
                except Exception as e:
                    logger.warning(f"Error closing workbook: {e}")
                finally:
                    self.workbook = None

            if self.excel:
                try:
                    # Excel 속성 복원
                    self.excel.ScreenUpdating = True
                    self.excel.EnableEvents = True
                    self.excel.DisplayAlerts = True
                except Exception as e:
                    logger.warning(f"Error restoring Excel properties: {e}")

                try:
                    self.excel.Quit()
                except Exception as e:
                    logger.warning(f"Error quitting Excel: {e}")
                finally:
                    self.excel = None

        except Exception as e:
            logger.error(f"Error cleaning up Excel COM objects: {e}")

    def create_odata_connection(
        self,
        odata_url: str,
        table_name: str = "Data",
        output_path: Optional[str] = None
    ) -> str:
        """
        OData 연결이 포함된 Excel 파일 생성

        Args:
            odata_url: OData 엔드포인트 URL (쿼리 파라미터 포함 가능)
            table_name: Excel 테이블 이름
            output_path: 출력 파일 경로 (None이면 임시 파일 생성)

        Returns:
            생성된 Excel 파일 경로
        """
        # COM 스레드 초기화 (멀티스레드 지원)
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

        try:
            # 출력 경로 설정
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".xlsx",
                    dir=tempfile.gettempdir()
                )
                output_path = output_file.name
                output_file.close()
            else:
                output_path = str(Path(output_path).absolute())

            # Excel 애플리케이션 시작 - 새로운 인스턴스 강제 생성
            logger.info("Starting Excel COM application...")
            max_retries = 3
            retry_count = 0

            while retry_count < max_retries:
                try:
                    # DispatchEx를 사용하여 새로운 Excel 인스턴스 생성
                    self.excel = win32com.client.DispatchEx("Excel.Application")
                    break
                except Exception as e:
                    retry_count += 1
                    if retry_count >= max_retries:
                        # 마지막 시도: 기존 Dispatch 사용
                        logger.warning(f"Failed to create new Excel instance, trying existing: {e}")
                        self.excel = win32com.client.Dispatch("Excel.Application")
                    else:
                        logger.warning(f"Retry {retry_count}/{max_retries}: {e}")
                        time.sleep(1)

            # Excel이 완전히 초기화되도록 잠시 대기
            time.sleep(0.5)

            # Excel 속성 설정
            try:
                self.excel.Visible = False
                self.excel.DisplayAlerts = False
                self.excel.ScreenUpdating = False
                self.excel.EnableEvents = False
                logger.info("Excel properties configured successfully")
            except Exception as e:
                logger.warning(f"Some Excel properties could not be set: {e}")

            # 새 워크북 생성
            logger.info("Creating new workbook...")
            retry_count = 0
            while retry_count < max_retries:
                try:
                    self.workbook = self.excel.Workbooks.Add()
                    logger.info("Workbook created successfully")
                    break
                except Exception as e:
                    retry_count += 1
                    if retry_count >= max_retries:
                        raise
                    logger.warning(f"Retry creating workbook {retry_count}/{max_retries}: {e}")
                    time.sleep(1)

            # 첫 번째 워크시트 가져오기
            worksheet = self.workbook.Worksheets(1)
            worksheet.Name = table_name

            # Power Query를 통한 OData 연결 추가
            # M 코드 (Power Query 언어) 생성
            m_code = f'''
let
    Source = OData.Feed("{odata_url}", null, [Implementation="2.0"])
in
    Source
'''

            # 쿼리 추가
            query_name = f"Query_{table_name}"

            # Queries 컬렉션에 새 쿼리 추가
            try:
                # WorkbookQuery 객체 생성
                query = self.workbook.Queries.Add(
                    Name=query_name,
                    Formula=m_code
                )


                # 쿼리를 테이블로 로드
                # ListObject (테이블) 생성
                list_object = worksheet.ListObjects.Add(
                    SourceType=0,  # xlSrcExternal
                    Source=f"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={query_name};Extended Properties=\"\"",
                    Destination=worksheet.Range("A1")
                )

                # 쿼리 테이블 설정
                query_table = list_object.QueryTable
                query_table.CommandType = 6  # xlCmdSql
                query_table.CommandText = f"SELECT * FROM [{query_name}]"
                query_table.RowNumbers = False
                query_table.FillAdjacentFormulas = False
                query_table.PreserveFormatting = True
                query_table.RefreshOnFileOpen = False
                query_table.RefreshStyle = 1  # xlInsertDeleteCells
                query_table.SavePassword = False
                query_table.SaveData = True
                query_table.AdjustColumnWidth = True
                query_table.RefreshPeriod = 0
                query_table.PreserveColumnInfo = True
                query_table.SourceConnectionFile = ""

                # 백그라운드 쿼리 설정 (비동기 로드)
                query_table.BackgroundQuery = True

                # 쿼리 실행 (데이터 로드)
                try:
                    query_table.Refresh(BackgroundQuery=False)  # 동기적으로 실행
                except Exception as e:
                    # 대용량 데이터의 경우 타임아웃이 발생할 수 있으므로 무시
                    pass

            except Exception as e:
                logger.error(f"Error adding query: {e}")
                # 대안: Connection 객체를 통한 연결 추가
                self._add_connection_alternative(worksheet, odata_url, table_name)

            # 파일 저장
            self.workbook.SaveAs(output_path, FileFormat=51)  # xlOpenXMLWorkbook (*.xlsx)

            # 워크북 닫기
            self.workbook.Close(True)
            self.workbook = None

            # Excel 종료
            self.excel.Quit()
            self.excel = None

            return output_path

        except Exception as e:
            logger.error(f"Error creating Excel with OData connection: {e}", exc_info=True)
            self.cleanup()
            raise
        finally:
            # COM 스레드 정리
            pythoncom.CoUninitialize()

    def _add_connection_alternative(self, worksheet, odata_url: str, table_name: str):
        """
        대안: WorkbookConnection을 통한 OData 연결 추가
        """
        try:
            connection_name = f"OData_{table_name}"

            # OData 연결 문자열
            connection_string = (
                f"OLEDB;Provider=Microsoft.Mashup.OleDb.1;"
                f"Data Source=$Workbook$;"
                f"Location=\"{odata_url}\";"
                f"Extended Properties=\"\""
            )

            # 연결 추가
            connection = self.workbook.Connections.Add2(
                Name=connection_name,
                Description=f"OData connection to {table_name}",
                ConnectionString=connection_string,
                CommandText="",
                lCmdtype=2  # xlCmdTable
            )

            # OData 연결 속성 설정
            connection.ODataConnection.SourceDataFile = odata_url
            connection.ODataConnection.SavePassword = False
            connection.RefreshWithRefreshAll = True

            # 데이터를 워크시트에 로드
            list_object = worksheet.ListObjects.Add(
                SourceType=4,  # xlSrcQuery
                Source=connection,
                Destination=worksheet.Range("A1")
            )

            list_object.Name = table_name
            list_object.QueryTable.Refresh(BackgroundQuery=False)

        except Exception as e:
            logger.error(f"Error in alternative connection method: {e}")
            # 최소한의 연결 정보만 추가
            self._add_minimal_connection(worksheet, odata_url, table_name)

    def create_odata_connection_with_webapi_auth(
        self,
        odata_url: str,
        table_name: str = "Data",
        output_path: Optional[str] = None,
        api_key_name: str = "Authorization",
        skip_data_load: bool = True
    ) -> str:
        """
        Web API 인증을 사용하는 OData 연결이 포함된 Excel 파일 생성

        Excel의 "웹 API" 인증 탭을 활용하여 API 키 인증을 설정합니다.
        사용자는 Excel에서 데이터 새로고침 시 웹 API 탭에서 키를 입력할 수 있습니다.

        Args:
            odata_url: OData 엔드포인트 URL (쿼리 파라미터 포함 가능)
            table_name: Excel 테이블 이름
            output_path: 출력 파일 경로 (None이면 임시 파일 생성)
            api_key_name: API 키의 이름 (기본값: "Authorization")
            skip_data_load: True면 연결 정보만 저장하고 데이터 로드는 건너뜀 (인증 필요한 경우)

        Returns:
            생성된 Excel 파일 경로
        """
        # COM 스레드 초기화 (멀티스레드 지원)
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

        try:
            # 출력 경로 설정
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".xlsx",
                    dir=tempfile.gettempdir()
                )
                output_path = output_file.name
                output_file.close()
            else:
                output_path = str(Path(output_path).absolute())

            # Excel 애플리케이션 시작 - 새로운 인스턴스 강제 생성
            logger.info("Starting Excel COM application for Web API auth...")
            max_retries = 3
            retry_count = 0

            while retry_count < max_retries:
                try:
                    # DispatchEx를 사용하여 새로운 Excel 인스턴스 생성
                    self.excel = win32com.client.DispatchEx("Excel.Application")
                    break
                except Exception as e:
                    retry_count += 1
                    if retry_count >= max_retries:
                        # 마지막 시도: 기존 Dispatch 사용
                        logger.warning(f"Failed to create new Excel instance, trying existing: {e}")
                        self.excel = win32com.client.Dispatch("Excel.Application")
                    else:
                        logger.warning(f"Retry {retry_count}/{max_retries}: {e}")
                        time.sleep(1)

            # Excel이 완전히 초기화되도록 잠시 대기
            time.sleep(0.5)

            # Excel 속성 설정
            try:
                self.excel.Visible = False
                self.excel.DisplayAlerts = False
                self.excel.ScreenUpdating = False
                self.excel.EnableEvents = False
                logger.info("Excel properties configured successfully")
            except Exception as e:
                logger.warning(f"Some Excel properties could not be set: {e}")

            # 새 워크북 생성
            logger.info("Creating new workbook...")
            retry_count = 0
            while retry_count < max_retries:
                try:
                    self.workbook = self.excel.Workbooks.Add()
                    logger.info("Workbook created successfully")
                    break
                except Exception as e:
                    retry_count += 1
                    if retry_count >= max_retries:
                        raise
                    logger.warning(f"Retry creating workbook {retry_count}/{max_retries}: {e}")
                    time.sleep(1)

            # 첫 번째 워크시트 가져오기
            worksheet = self.workbook.Worksheets(1)
            worksheet.Name = table_name

            # Power Query M 코드 생성 - Web API 인증 설정 포함
            # ApiKeyName을 지정하여 Excel이 웹 API 탭을 표시하도록 함
            m_code = f'''
let
    Source = OData.Feed(
        "{odata_url}",
        null,
        [
            Implementation="2.0",
            ApiKeyName="{api_key_name}"
        ]
    )
in
    Source
'''

            # 쿼리 추가
            query_name = f"Query_{table_name}_WebAPI"

            # Queries 컬렉션에 새 쿼리 추가
            try:
                # WorkbookQuery 객체 생성
                query = self.workbook.Queries.Add(
                    Name=query_name,
                    Formula=m_code
                )

                # 쿼리를 테이블로 로드
                list_object = worksheet.ListObjects.Add(
                    SourceType=0,  # xlSrcExternal
                    Source=f"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={query_name};Extended Properties=\"\"",
                    Destination=worksheet.Range("A1")
                )

                # 쿼리 테이블 설정
                query_table = list_object.QueryTable
                query_table.CommandType = 6  # xlCmdSql
                query_table.CommandText = f"SELECT * FROM [{query_name}]"
                query_table.RowNumbers = False
                query_table.FillAdjacentFormulas = False
                query_table.PreserveFormatting = True
                query_table.RefreshOnFileOpen = False
                query_table.RefreshStyle = 1  # xlInsertDeleteCells
                query_table.SavePassword = False
                query_table.SaveData = True
                query_table.AdjustColumnWidth = True
                query_table.RefreshPeriod = 0
                query_table.PreserveColumnInfo = True
                query_table.SourceConnectionFile = ""
                query_table.BackgroundQuery = True

                # 데이터 로드는 기본적으로 건너뜀 (인증이 필요하므로)
                if not skip_data_load:
                    try:
                        query_table.Refresh(BackgroundQuery=False)
                    except Exception as e:
                        logger.warning(f"Failed to refresh data during file creation (expected with API auth): {e}")
                        pass
                else:
                    logger.info("Skipping data load - Web API authentication will be required on refresh")

            except Exception as e:
                logger.error(f"Error adding Web API query: {e}")
                # 대안: 연결 정보와 안내 추가
                self._add_webapi_connection_guide(worksheet, odata_url, table_name, api_key_name)

            # 파일 저장
            self.workbook.SaveAs(output_path, FileFormat=51)  # xlOpenXMLWorkbook (*.xlsx)

            # 워크북 닫기
            self.workbook.Close(True)
            self.workbook = None

            # Excel 종료
            self.excel.Quit()
            self.excel = None

            return output_path

        except Exception as e:
            logger.error(f"Error creating Excel with Web API OData connection: {e}", exc_info=True)
            self.cleanup()
            raise
        finally:
            # COM 스레드 정리
            pythoncom.CoUninitialize()

    def _add_webapi_connection_guide(self, worksheet, odata_url: str, table_name: str, api_key_name: str):
        """
        Web API 연결 정보 및 가이드 추가
        """
        try:
            # 안내 텍스트 추가
            worksheet.Range("A1").Value = "OData Web API 인증 템플릿"
            worksheet.Range("A2").Value = "URL:"
            worksheet.Range("B2").Value = odata_url
            worksheet.Range("A3").Value = "API Key Name:"
            worksheet.Range("B3").Value = api_key_name
            worksheet.Range("A5").Value = "사용 방법:"
            worksheet.Range("A6").Value = "1. 상단 '데이터' 탭 클릭"
            worksheet.Range("A7").Value = "2. '쿼리 및 연결' 클릭하여 쿼리 새로고침"
            worksheet.Range("A8").Value = "3. 인증 창이 나타나면 '웹 API' 탭 선택"
            worksheet.Range("A9").Value = f"4. 키 입력란에 'Bearer <your_token>' 형식으로 입력"
            worksheet.Range("A10").Value = "   예: Bearer tok_abc123xyz"
            worksheet.Range("A11").Value = "5. '연결' 클릭하여 데이터 로드"

            # 서식 설정
            worksheet.Range("A1").Font.Bold = True
            worksheet.Range("A1").Font.Size = 14
            worksheet.Range("A2:A11").Font.Bold = True
            worksheet.Range("B2:B3").Font.Color = -16776961  # 파란색

            # 열 너비 자동 조정
            worksheet.Columns("A:B").AutoFit()

        except Exception as e:
            logger.error(f"Error adding Web API connection guide: {e}")

    def _add_minimal_connection(self, worksheet, odata_url: str, table_name: str):
        """
        연결 정보 추가 및 가이드
        """
        try:
            # 안내 텍스트 추가
            worksheet.Range("A1").Value = "OData 데이터 템플릿"
            worksheet.Range("A2").Value = "URL:"
            worksheet.Range("B2").Value = odata_url
            worksheet.Range("A4").Value = "사용 방법:"
            worksheet.Range("A5").Value = "1. 상단 '데이터' 탭 클릭"
            worksheet.Range("A6").Value = "2. '쿼리 및 연결' 클릭"
            worksheet.Range("A7").Value = "3. 우측 사이드바에서 템플릿 이름의 쿼리를 우클릭하여 '다음으로 로드'"
            worksheet.Range("A8").Value = "4. '연결만 만들기' 및 '데이터 모델에 이 데이터 추가' 선택 후 '확인'"

            # 서식 설정
            worksheet.Range("A1").Font.Bold = True
            worksheet.Range("A1").Font.Size = 14
            worksheet.Range("A2:A8").Font.Bold = True
            worksheet.Range("B2").Font.Color = -16776961  # 파란색

            # 열 너비 자동 조정
            worksheet.Columns("A:B").AutoFit()

        except Exception as e:
            logger.error(f"Error adding minimal connection: {e}")


def create_excel_with_odata_com(
    odata_url: str,
    table_name: str = "Data",
    output_path: Optional[str] = None
) -> str:
    """
    편의 함수: OData 연결이 포함된 Excel 파일 생성

    Args:
        odata_url: OData 엔드포인트 URL
        table_name: Excel 테이블 이름
        output_path: 출력 파일 경로

    Returns:
        생성된 Excel 파일 경로
    """
    generator = ExcelCOMGenerator()
    try:
        return generator.create_odata_connection(odata_url, table_name, output_path)
    finally:
        generator.cleanup()


def create_excel_with_webapi_auth_com(
    odata_url: str,
    table_name: str = "Data",
    output_path: Optional[str] = None,
    api_key_name: str = "Authorization",
    skip_data_load: bool = True
) -> str:
    """
    편의 함수: Web API 인증을 사용하는 OData 연결이 포함된 Excel 파일 생성

    Excel의 "웹 API" 탭에서 API 키를 입력할 수 있도록 구성됩니다.
    사용자는 "Bearer <token>" 형식으로 키를 입력해야 합니다.

    Args:
        odata_url: OData 엔드포인트 URL
        table_name: Excel 테이블 이름
        output_path: 출력 파일 경로
        api_key_name: API 키의 이름 (기본값: "Authorization")
        skip_data_load: True면 연결 정보만 저장하고 데이터 로드는 건너뜀

    Returns:
        생성된 Excel 파일 경로
    """
    generator = ExcelCOMGenerator()
    try:
        return generator.create_odata_connection_with_webapi_auth(
            odata_url, table_name, output_path, api_key_name, skip_data_load
        )
    finally:
        generator.cleanup()