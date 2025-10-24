"""
Excel Power Query Connection Modifier Service
템플릿 Excel 파일의 OData 연결 정보를 동적으로 변경
"""
import logging
import shutil
import tempfile
from pathlib import Path
from typing import Optional
from zipfile import ZipFile, ZIP_DEFLATED, ZipInfo
import zipfile
import xml.etree.ElementTree as ET
from app.services.datamashup_rebuilder import DataMashupRebuilder

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

        DataMashup 내부 ZIP을 완전히 재구축하여 안정성 보장

        Args:
            new_odata_url: 새로운 OData endpoint URL
            output_path: 출력 파일 경로 (None이면 임시 파일 생성)

        Returns:
            수정된 Excel 파일 경로
        """
        try:
            # 출력 파일 경로 결정
            if output_path is None:
                output_file = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".xlsx"
                )
                output_path = output_file.name
                output_file.close()

            # DataMashup 재구축기 초기화
            rebuilder = DataMashupRebuilder()

            # ZIP에서 customXml/item1.xml 읽기
            with ZipFile(self.template_path, 'r') as zip_in:
                try:
                    item1_content = zip_in.read('customXml/item1.xml')
                except KeyError:
                    logger.warning("customXml/item1.xml not found")
                    # 원본 복사
                    shutil.copy2(self.template_path, output_path)
                    return output_path

            # item1.xml 수정 (DataMashup ZIP 완전 재구축)
            new_item1_content = rebuilder.modify_item1_xml(item1_content, new_odata_url)

            # 새 Excel 파일 생성
            with ZipFile(self.template_path, 'r') as zip_in:
                with ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zip_out:
                    for item in zip_in.infolist():
                        if item.filename == 'customXml/item1.xml':
                            # 수정된 item1.xml 사용
                            zip_out.writestr(item, new_item1_content)
                        else:
                            # 원본 그대로 복사
                            zip_out.writestr(item, zip_in.read(item.filename))

            logger.info(f"Created modified Excel file: {output_path}")
            logger.info(f"Updated OData URL to: {new_odata_url}")

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
        DataMashup (Power Query M 코드) 내의 OData.Feed URL 수정

        Returns:
            수정 여부
        """
        try:
            import base64
            import re
            import zlib
            import struct
            import binascii
            import zipfile
            import io
            import tempfile

            # 파일을 바이너리로 읽기 (UTF-16LE 인코딩)
            with open(file_path, 'rb') as f:
                content_bytes = f.read()

            # UTF-16LE로 디코딩
            content = content_bytes.decode('utf-16-le', errors='ignore')

            # DataMashup 태그 찾기
            match = re.search(r'<DataMashup[^>]*>([^<]+)</DataMashup>', content)
            if not match:
                logger.debug(f"No DataMashup tag found in {file_path.name}")
                return False

            base64_data = match.group(1)
            # 공백 제거
            base64_data = base64_data.replace(' ', '').replace('\n', '').replace('\r', '')

            # Base64 디코딩
            decoded = base64.b64decode(base64_data)

            # DataMashup은 8바이트 헤더 + ZIP 파일 구조
            # 첫 8바이트는 헤더 (크기 정보 등)
            header = decoded[:8]

            # Section1.m 찾기 (low-level 방식)
            section_index = decoded.find(b'Section1.m')
            if section_index == -1:
                logger.debug("Section1.m not found in DataMashup")
                return False

            logger.debug(f"Found Section1.m at offset {section_index}")

            # Section1.m을 포함하는 PK 로컬 헤더 찾기
            pk_offset = -1
            for i in range(max(0, section_index - 100), section_index):
                if decoded[i:i+4] == b'PK\x03\x04':
                    pk_offset = i
                    break

            if pk_offset == -1:
                logger.error("Could not find PK signature for Section1.m")
                return False

            logger.debug(f"Found PK signature at offset {pk_offset}")

            # ZIP 로컬 파일 헤더 파싱
            compression = struct.unpack('<H', decoded[pk_offset+8:pk_offset+10])[0]
            compressed_size = struct.unpack('<I', decoded[pk_offset+18:pk_offset+22])[0]
            uncompressed_size = struct.unpack('<I', decoded[pk_offset+22:pk_offset+26])[0]
            filename_len = struct.unpack('<H', decoded[pk_offset+26:pk_offset+28])[0]
            extra_len = struct.unpack('<H', decoded[pk_offset+28:pk_offset+30])[0]

            # 압축된 데이터 위치
            data_start = pk_offset + 30 + filename_len + extra_len
            compressed_data = decoded[data_start:data_start+compressed_size]

            # 압축 해제
            if compression == 8:
                m_code = zlib.decompress(compressed_data, -15).decode('utf-8')
            else:
                m_code = compressed_data.decode('utf-8')

            logger.debug(f"Original M code:\n{m_code}")

            # OData URL 패턴 찾기 (정확한 URL이 없을 수도 있으므로)
            odata_pattern = r'OData\.Feed\("([^"]+)"'
            url_match = re.search(odata_pattern, m_code)

            if url_match:
                current_url = url_match.group(1)
                logger.debug(f"Found URL in M code: {current_url}")
                new_m_code = m_code.replace(current_url, new_url)
                logger.debug(f"Updated M code:\n{new_m_code}")

                # 다시 압축
                new_m_code_bytes = new_m_code.encode('utf-8')
                if compression == 8:
                    new_compressed = zlib.compress(new_m_code_bytes)[2:-4]  # Remove zlib header and trailer
                else:
                    new_compressed = new_m_code_bytes

                # 새로운 크기 계산
                new_compressed_size = len(new_compressed)
                new_uncompressed_size = len(new_m_code_bytes)

                # CRC32 계산 (비압축 데이터에 대해)
                new_crc32 = binascii.crc32(new_m_code_bytes) & 0xffffffff

                # 데이터 교체 및 Central Directory 업데이트
                size_diff = new_compressed_size - compressed_size

                # 새로운 바이너리 데이터 구성
                new_decoded = bytearray(decoded[:data_start]) + new_compressed + decoded[data_start + compressed_size:]

                # ZIP 로컬 파일 헤더 업데이트 (데이터 교체 후에 수행)
                # CRC32 업데이트
                new_decoded[pk_offset+14:pk_offset+18] = struct.pack('<I', new_crc32)
                # 압축된 크기 업데이트
                new_decoded[pk_offset+18:pk_offset+22] = struct.pack('<I', new_compressed_size)
                # 비압축 크기 업데이트
                new_decoded[pk_offset+22:pk_offset+26] = struct.pack('<I', new_uncompressed_size)

                # Central Directory 항목도 업데이트 필요
                # ZIP 파일 끝에서 Central Directory 찾기
                cd_signature = b'PK\x01\x02'
                cd_offset = new_decoded.find(cd_signature, data_start + new_compressed_size)

                if cd_offset != -1:
                    # Central Directory의 Section1.m 항목 찾기
                    # Central Directory 항목은 local file header와 유사한 구조
                    while cd_offset < len(new_decoded) - 46:
                        if new_decoded[cd_offset:cd_offset+4] == cd_signature:
                            # 파일명 길이 확인
                            cd_filename_len = struct.unpack('<H', new_decoded[cd_offset+28:cd_offset+30])[0]
                            cd_extra_len = struct.unpack('<H', new_decoded[cd_offset+30:cd_offset+32])[0]
                            cd_comment_len = struct.unpack('<H', new_decoded[cd_offset+32:cd_offset+34])[0]

                            # 파일명 추출
                            filename_start = cd_offset + 46
                            cd_filename = new_decoded[filename_start:filename_start+cd_filename_len]

                            if b'Section1.m' in cd_filename:
                                # Central Directory의 크기 정보 업데이트
                                new_decoded[cd_offset+20:cd_offset+24] = struct.pack('<I', new_compressed_size)
                                new_decoded[cd_offset+24:cd_offset+28] = struct.pack('<I', new_uncompressed_size)
                                new_decoded[cd_offset+16:cd_offset+20] = struct.pack('<I', new_crc32)
                                break

                            # 다음 Central Directory 항목으로
                            cd_offset += 46 + cd_filename_len + cd_extra_len + cd_comment_len
                        else:
                            break

                # EOCD (End of Central Directory) 업데이트
                # 첫 번째 EOCD 찾기 (내부 ZIP의 종료 레코드)
                eocd_signature = b'PK\x05\x06'
                eocd_offset = new_decoded.find(eocd_signature, data_start + new_compressed_size)

                if eocd_offset != -1 and eocd_offset < len(new_decoded) - 22:
                    # EOCD의 CD offset 업데이트
                    # CD offset은 ZIP 시작(offset 8)을 기준으로 하므로, 8을 빼야 함
                    old_cd_offset = struct.unpack('<I', new_decoded[eocd_offset+16:eocd_offset+20])[0]
                    new_cd_offset = old_cd_offset + size_diff

                    new_decoded[eocd_offset+16:eocd_offset+20] = struct.pack('<I', new_cd_offset)
                    logger.debug(f"Updated EOCD CD offset: {old_cd_offset} → {new_cd_offset}")

                # 새로운 Base64 인코딩
                new_base64 = base64.b64encode(bytes(new_decoded)).decode('ascii')

                # XML 문서 재구성
                new_content = content[:match.start(1)] + new_base64 + content[match.end(1):]

                # 파일 저장 (UTF-16LE)
                with open(file_path, 'wb') as f:
                    f.write(new_content.encode('utf-16-le'))

                logger.info(f"Successfully updated OData URL in {file_path.name} from {current_url} to {new_url}")
                return True
            else:
                logger.warning("No OData.Feed URL found in M code")
                dmz.close()
                return False

        except Exception as e:
            logger.error(f"Error modifying custom XML: {str(e)}", exc_info=True)
            return False
