"""
Connection File Generator
Excel에서 수동으로 연결을 설정할 수 있는 다양한 파일 형식 생성
"""
import logging
import tempfile
from pathlib import Path
from typing import Optional
import json

logger = logging.getLogger(__name__)


class ConnectionFileGenerator:
    """
    Excel 연결 파일 생성기
    """

    def generate_dqy_file(
        self,
        odata_url: str,
        table_name: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        DQY (Database Query) 파일 생성
        Excel의 Legacy Web Query 파일 형식

        Args:
            odata_url: OData 서비스 URL
            table_name: 테이블 이름
            output_path: 출력 경로

        Returns:
            생성된 DQY 파일 경로
        """
        try:
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".dqy"
                )
                output_path = output_file.name
                output_file.close()

            # DQY 파일 내용 (더 간단한 형식)
            dqy_content = f"""WEB
1
{odata_url}

Selection={table_name}
Formatting=None
PreFormattedTextToColumns=True
ConsecutiveDelimitersAsOne=True
SingleBlockTextImport=False
DisableDateRecognition=False
DisableRedirections=False"""

            # 파일 저장
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(dqy_content)

            logger.info(f"Generated DQY file: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Error generating DQY file: {str(e)}", exc_info=True)
            raise

    def generate_iqy_file(
        self,
        odata_url: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        IQY (Internet Query) 파일 생성
        Excel의 Web Query 파일

        Args:
            odata_url: OData 서비스 URL
            output_path: 출력 경로

        Returns:
            생성된 IQY 파일 경로
        """
        try:
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".iqy"
                )
                output_path = output_file.name
                output_file.close()

            # IQY 파일 내용 (매우 간단)
            iqy_content = f"""WEB
1
{odata_url}?$format=json

Selection=
Formatting=None
PreFormattedTextToColumns=True
ConsecutiveDelimitersAsOne=True
SingleBlockTextImport=False
DisableDateRecognition=False
DisableRedirections=False
"""

            # 파일 저장
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(iqy_content)

            logger.info(f"Generated IQY file: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Error generating IQY file: {str(e)}", exc_info=True)
            raise

    def generate_connection_txt(
        self,
        odata_url: str,
        table_name: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        연결 정보 텍스트 파일 생성
        사용자가 복사/붙여넣기 할 수 있는 단순 텍스트 파일

        Args:
            odata_url: OData 서비스 URL
            table_name: 테이블 이름
            output_path: 출력 경로

        Returns:
            생성된 텍스트 파일 경로
        """
        try:
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".txt"
                )
                output_path = output_file.name
                output_file.close()

            # 텍스트 파일 내용
            txt_content = f"""===== OData 연결 정보 =====

테이블: {table_name}
URL: {odata_url}

===== Excel에서 연결하는 방법 =====

방법 1: Power Query 사용
------------------------
1. Excel 열기
2. 데이터 탭 → 데이터 가져오기 → 기타 원본에서 → OData 피드에서
3. URL 입력: {odata_url}
4. 연결 → 로드

방법 2: Power Query M 코드 직접 입력
------------------------------------
1. Excel 열기
2. 데이터 탭 → 데이터 가져오기 → 빈 쿼리 시작
3. 고급 편집기 열기
4. 아래 코드 붙여넣기:

let
    Source = OData.Feed("{odata_url}", null, [Implementation="2.0"])
in
    Source

5. 완료 → 닫기 및 로드

===== 필터 예제 =====

특정 조건으로 필터링:
{odata_url}?$filter=Media eq 'Naver'

특정 필드만 선택:
{odata_url}?$select=Date,Campaign,Clicks

정렬:
{odata_url}?$orderby=Date desc

상위 N개만:
{odata_url}?$top=100

===== 참고 사항 =====

- 이 연결은 수동으로 새로고침해야 합니다
- 데이터 탭 → 모두 새로고침 버튼을 사용하세요
- 자동 새로고침을 원하면 Excel의 연결 속성에서 설정 가능합니다
"""

            # 파일 저장
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(txt_content)

            logger.info(f"Generated connection text file: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Error generating text file: {str(e)}", exc_info=True)
            raise

    def generate_vba_script(
        self,
        odata_url: str,
        table_name: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        VBA 스크립트 생성
        Excel에서 매크로로 실행할 수 있는 VBA 코드

        Args:
            odata_url: OData 서비스 URL
            table_name: 테이블 이름
            output_path: 출력 경로

        Returns:
            생성된 VBA 파일 경로
        """
        try:
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".vba"
                )
                output_path = output_file.name
                output_file.close()

            # VBA 스크립트
            vba_content = f"""' OData 연결 VBA 매크로
' 이 코드를 Excel VBA 편집기에 붙여넣고 실행하세요

Sub CreateODataConnection()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim conn As WorkbookConnection

    Set wb = ActiveWorkbook
    Set ws = ActiveSheet

    ' Power Query 연결 생성
    wb.Queries.Add Name:="{table_name}", _
        Formula:="let" & Chr(13) & Chr(10) & _
                "    Source = OData.Feed(""{odata_url}"", null, [Implementation=""2.0""])" & Chr(13) & Chr(10) & _
                "in" & Chr(13) & Chr(10) & _
                "    Source"

    ' 연결을 워크시트에 로드
    With ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={table_name};Extended Properties=", _
        Destination:=Range("$A$1")).QueryTable
        .CommandText = Array("SELECT * FROM [{table_name}]")
        .ListObject.Name = "{table_name}"
        .Refresh BackgroundQuery:=False
    End With

    MsgBox "OData 연결이 생성되었습니다!", vbInformation
End Sub

Sub RefreshODataConnection()
    ' 데이터 새로고침
    ActiveWorkbook.RefreshAll
    MsgBox "데이터가 새로고침되었습니다!", vbInformation
End Sub

Sub DeleteODataConnection()
    ' 연결 삭제
    On Error Resume Next
    ActiveWorkbook.Queries("{table_name}").Delete
    ActiveSheet.ListObjects("{table_name}").Delete
    MsgBox "연결이 삭제되었습니다!", vbInformation
End Sub
"""

            # 파일 저장
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(vba_content)

            logger.info(f"Generated VBA script file: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Error generating VBA script: {str(e)}", exc_info=True)
            raise