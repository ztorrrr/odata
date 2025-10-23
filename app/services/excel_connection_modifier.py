"""
Excel Power Query Connection Modifier Service
템플릿 Excel 파일의 OData 연결 정보를 동적으로 변경
"""
import logging
import shutil
import tempfile
from pathlib import Path
from typing import Optional
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

logger = logging.getLogger(__name__)


class ExcelConnectionModifier:
    """
    - Excel 파일(.xlsx)의 Power Query 연결 정보를 수정
    - Power Query 연결 정보는 xl/connections.xml 및 xl/queries/ 폴더에 저장되므로 이를 수정
    """

    def __init__(self, template_path: str):
        """
        Args:
            template_path: 템플릿 Excel 파일 경로
        """
        self.template_path = Path(template_path)

        if not self.template_path.exists():
            raise FileNotFoundError(f"Template file not found: {template_path}")

    def modify_odata_connection(
        self,
        new_odata_url: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        Excel 파일의 OData 연결 URL을 새로운 URL로 변경

        Args:
            new_odata_url: 새로운 OData endpoint URL
            output_path: 출력 파일 경로 (None이면 임시 파일 생성)

        Returns:
            수정된 Excel 파일 경로
        """
        try:
            # 임시 디렉토리 생성
            temp_dir = Path(tempfile.mkdtemp())

            # Excel 파일을 ZIP으로 압축 해제
            with ZipFile(self.template_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            logger.info(f"Extracted template to: {temp_dir}")

            # Power Query 연결 정보 수정
            modified = False

            # 1. xl/connections.xml 수정
            connections_path = temp_dir / "xl" / "connections.xml"
            if connections_path.exists():
                if self._modify_connections_xml(connections_path, new_odata_url):
                    modified = True
                    logger.info("Modified xl/connections.xml")

            # 2. xl/queries/ 폴더의 쿼리 파일들 수정
            queries_dir = temp_dir / "xl" / "queries"
            if queries_dir.exists():
                for query_file in queries_dir.glob("*.xml"):
                    if self._modify_query_xml(query_file, new_odata_url):
                        modified = True
                        logger.info(f"Modified {query_file.name}")

            # 3. customXml/ 폴더의 연결 정보 수정 (있는 경우)
            customxml_dir = temp_dir / "customXml"
            if customxml_dir.exists():
                for xml_file in customxml_dir.glob("*.xml"):
                    if self._modify_custom_xml(xml_file, new_odata_url):
                        modified = True
                        logger.info(f"Modified {xml_file.name}")

            if not modified:
                logger.warning("No OData connections found in template file")

            # 출력 파일 경로 결정
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".xlsx"
                )
                output_path = output_file.name
                output_file.close()

            # 수정된 파일들을 다시 ZIP으로 압축
            with ZipFile(output_path, 'w', ZIP_DEFLATED) as zip_out:
                for file_path in temp_dir.rglob('*'):
                    if file_path.is_file():
                        arcname = file_path.relative_to(temp_dir)
                        zip_out.write(file_path, arcname)

            logger.info(f"Created modified Excel file: {output_path}")

            # 임시 디렉토리 정리
            shutil.rmtree(temp_dir)

            return output_path

        except Exception as e:
            logger.error(f"Error modifying Excel connection: {str(e)}", exc_info=True)
            raise

    def _modify_connections_xml(self, file_path: Path, new_url: str) -> bool:
        """
        xl/connections.xml 파일의 OData URL 수정

        Returns:
            수정 여부
        """
        try:
            # XML 네임스페이스 등록
            namespaces = {
                '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
            }

            # 네임스페이스를 prefix로 등록
            for prefix, uri in namespaces.items():
                if prefix:
                    ET.register_namespace(prefix, uri)

            tree = ET.parse(file_path)
            root = tree.getroot()

            modified = False

            # connection 요소들 찾기
            for connection in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}connection'):
                # OData 연결 찾기
                odcFile = connection.get('odcFile')
                if odcFile and 'odata' in odcFile.lower():
                    # dbPr 요소의 connection string 수정
                    for dbPr in connection.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dbPr'):
                        conn_str = dbPr.get('connection')
                        if conn_str and 'Location=' in conn_str:
                            # Location= 뒤의 URL 교체
                            parts = conn_str.split('Location=')
                            if len(parts) == 2:
                                new_conn_str = f"{parts[0]}Location={new_url}"
                                dbPr.set('connection', new_conn_str)
                                modified = True
                                logger.debug(f"Updated connection string to: {new_url}")

            if modified:
                tree.write(file_path, encoding='UTF-8', xml_declaration=True)

            return modified

        except Exception as e:
            logger.error(f"Error modifying connections.xml: {str(e)}", exc_info=True)
            return False

    def _modify_query_xml(self, file_path: Path, new_url: str) -> bool:
        """
        xl/queries/*.xml 파일의 Power Query M 코드에서 OData URL 수정

        Returns:
            수정 여부
        """
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            modified = False

            # query 요소의 formula 찾기
            namespaces = {
                '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            }

            for prefix, uri in namespaces.items():
                if prefix:
                    ET.register_namespace(prefix, uri)

            # M 코드가 포함된 formula 요소 찾기
            for query in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}query'):
                formula = query.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}formula')

                if formula is not None and formula.text:
                    m_code = formula.text

                    # OData.Feed 함수 찾기 및 URL 교체
                    if 'OData.Feed' in m_code:
                        # 정규식 대신 간단한 문자열 처리
                        lines = m_code.split('\n')
                        for i, line in enumerate(lines):
                            if 'OData.Feed' in line and '"http' in line:
                                # URL 부분 찾기 (큰따옴표 사이의 URL)
                                start = line.find('"http')
                                if start != -1:
                                    end = line.find('"', start + 1)
                                    if end != -1:
                                        # URL 교체
                                        lines[i] = line[:start + 1] + new_url + line[end:]
                                        modified = True
                                        logger.debug(f"Updated M code URL to: {new_url}")

                        if modified:
                            formula.text = '\n'.join(lines)

            if modified:
                tree.write(file_path, encoding='UTF-8', xml_declaration=True)

            return modified

        except Exception as e:
            logger.error(f"Error modifying query XML: {str(e)}", exc_info=True)
            return False

    def _modify_custom_xml(self, file_path: Path, new_url: str) -> bool:
        """
        customXml/*.xml 파일의 OData URL 수정

        Returns:
            수정 여부
        """
        try:
            # 파일 내용 읽기
            content = file_path.read_text(encoding='utf-8')

            # OData URL 패턴 찾기 및 교체
            if 'http' in content and ('odata' in content.lower() or 'OData' in content):
                # 간단한 URL 패턴 찾기
                import re
                pattern = r'https?://[^\s<>"]+/odata[^\s<>"]*'

                if re.search(pattern, content):
                    new_content = re.sub(pattern, new_url, content)
                    file_path.write_text(new_content, encoding='utf-8')
                    logger.debug(f"Updated custom XML URL to: {new_url}")
                    return True

            return False

        except Exception as e:
            logger.error(f"Error modifying custom XML: {str(e)}", exc_info=True)
            return False
