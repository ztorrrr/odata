"""
Excel with Power Query Connection Generator
Power Query 연결이 이미 설정된 Excel 파일 생성 (원본 템플릿 기반)
"""
import logging
import tempfile
import zipfile
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
import base64
import uuid
import re

logger = logging.getLogger(__name__)


class ExcelWithPowerQueryGenerator:
    """
    Power Query 연결이 설정된 Excel 파일 생성

    - 간단한 DataMashup 사용 (중첩 ZIP 없음)
    - workbook.xml에 dataModel 참조 추가
    - "쿼리 및 연결"에 쿼리가 표시됨
    - 사용자가 새로고침으로 데이터 로드
    """

    def __init__(self, template_path: Optional[str] = None):
        """
        Args:
            template_path: 원본 템플릿 경로 (기본: app/template/odata_template.xlsx)
        """
        if template_path is None:
            # 기본 템플릿 경로
            self.template_path = Path(__file__).parent.parent / "template" / "odata_template.xlsx"
        else:
            self.template_path = Path(template_path)

        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {self.template_path}")

    def generate_excel_with_power_query(
        self,
        json_api_url: str,
        table_name: str,
        query_name: Optional[str] = None,
        output_path: Optional[str] = None
    ) -> str:
        """
        Power Query 연결이 설정된 Excel 파일 생성

        원본 템플릿을 기반으로 하되, DataMashup의 M 코드만 수정
        - model/item.data 포함 (원본 그대로 유지)
        - 간단한 DataMashup으로 M 코드 교체 (중첩 ZIP 없음)

        Args:
            json_api_url: JSON API URL
            table_name: 테이블 이름
            query_name: 쿼리 이름 (기본: table_name)
            output_path: 출력 경로

        Returns:
            생성된 Excel 파일 경로
        """
        try:
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".xlsx"
                )
                output_path = output_file.name
                output_file.close()

            if query_name is None:
                query_name = table_name

            # Power Query M 코드 생성 (JSON API용)
            m_code_text = f'''let
    Source = Json.Document(Web.Contents("{json_api_url}")),
    value = Source[value],
    ToTable = Table.FromRecords(value)
in
    ToTable'''

            # 1. 원본 템플릿 복사하여 시작
            # 템플릿에는 이미 model/item.data와 모든 필수 구조가 포함되어 있음
            temp_dir = tempfile.mkdtemp()

            with zipfile.ZipFile(self.template_path, 'r') as z:
                z.extractall(temp_dir)

            logger.debug(f"Extracted template to {temp_dir}")

            # 2. Sheet1의 내용만 수정
            ws_path = Path(temp_dir) / "xl" / "worksheets" / "sheet1.xml"
            if ws_path.exists():
                # 간단한 안내 메시지로 교체
                sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1">
<c r="A1" t="inlineStr"><is><t>{table_name} - Power Query 연결</t></is></c>
</row>
<row r="3">
<c r="A3" t="inlineStr"><is><t>✅ 이 파일에는 Power Query 연결이 설정되어 있습니다.</t></is></c>
</row>
<row r="5">
<c r="A5" t="inlineStr"><is><t>데이터 로드 방법:</t></is></c>
</row>
<row r="6">
<c r="A6" t="inlineStr"><is><t>1. 데이터 탭 → 쿼리 및 연결 클릭</t></is></c>
</row>
<row r="7">
<c r="A7" t="inlineStr"><is><t>2. '{query_name}' 쿼리를 마우스 오른쪽 클릭</t></is></c>
</row>
<row r="8">
<c r="A8" t="inlineStr"><is><t>3. '로드 대상...' 선택하여 데이터 로드</t></is></c>
</row>
<row r="10">
<c r="A10" t="inlineStr"><is><t>또는 '데이터' 탭 → '모두 새로고침' 클릭</t></is></c>
</row>
</sheetData>
</worksheet>'''
                ws_path.write_text(sheet_xml, encoding='utf-8')
                logger.debug("Updated sheet1.xml")

            # 3. customXml/item1.xml 수정 (간단한 DataMashup으로 교체)
            customxml_dir = Path(temp_dir) / "customXml"

            # 간단한 Mashup XML (중첩 ZIP 없음) - JSON API용 M 코드
            mashup_xml = f'''<Mashup xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/DataMashup">
<Client>EXCEL</Client>
<Version>2.116.622.0</Version>
<MinVersion>2.21.0.0</MinVersion>
<Culture>ko-KR</Culture>
<SafeCombine>false</SafeCombine>
<Items>
<Query Name="{query_name}">
<Formula><![CDATA[{m_code_text}]]></Formula>
<IsParameterQuery xsi:nil="true" />
<IsDirectQuery xsi:nil="true" />
</Query>
</Items>
</Mashup>'''

            # Base64 인코딩 (UTF-8로)
            mashup_base64 = base64.b64encode(mashup_xml.encode('utf-8')).decode('ascii')

            # item1.xml 교체 (UTF-8로 저장)
            item1_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">{mashup_base64}</DataMashup>'''

            item1_path = customxml_dir / "item1.xml"
            item1_path.write_text(item1_xml, encoding='utf-8')

            logger.debug(f"Replaced DataMashup with simple structure (no nested ZIP)")

            # 4. workbook.xml의 modelTable 연결 이름 업데이트
            workbook_path = Path(temp_dir) / "xl" / "workbook.xml"

            # workbook.xml을 텍스트로 읽어서 연결 이름 교체
            workbook_content = workbook_path.read_text(encoding='utf-8')

            # "쿼리 - 쿼리1" → "쿼리 - {query_name}"으로 교체
            workbook_content = workbook_content.replace('쿼리 - 쿼리1', f'쿼리 - {query_name}')
            # "쿼리1" → "{query_name}"으로 교체 (단, "쿼리 - " 뒤가 아닌 경우만)
            workbook_content = re.sub(r'name="쿼리1"', f'name="{query_name}"', workbook_content)

            workbook_path.write_text(workbook_content, encoding='utf-8')

            logger.debug(f"Updated workbook.xml: 쿼리1 → {query_name}")

            # 5. connections.xml 업데이트
            connections_path = Path(temp_dir) / "xl" / "connections.xml"
            connections_content = connections_path.read_text(encoding='utf-8')

            # "쿼리 - 쿼리1" → "쿼리 - {query_name}"으로 교체
            connections_content = connections_content.replace('쿼리 - 쿼리1', f'쿼리 - {query_name}')
            connections_content = re.sub(r'id="쿼리1"', f'id="{query_name}"', connections_content)

            connections_path.write_text(connections_content, encoding='utf-8')

            logger.debug(f"Updated connections.xml: 쿼리1 → {query_name}")

            # 6. 다시 ZIP으로 압축
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for file_path in Path(temp_dir).rglob('*'):
                    if file_path.is_file():
                        arcname = file_path.relative_to(temp_dir)
                        zf.write(file_path, arcname)

            # 7. 정리
            shutil.rmtree(temp_dir)

            logger.info(f"Generated Excel with Power Query connection: {output_path}")
            logger.info(f"Query: {query_name}, API: {json_api_url}")

            return output_path

        except Exception as e:
            logger.error(f"Error generating Excel with Power Query: {str(e)}", exc_info=True)
            # 정리
            if 'temp_xlsx' in locals():
                Path(temp_xlsx.name).unlink(missing_ok=True)
            if 'temp_dir' in locals():
                shutil.rmtree(temp_dir, ignore_errors=True)
            raise
