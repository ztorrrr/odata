"""
DataMashup ZIP Rebuilder
Excel Power Query DataMashup 내부 ZIP을 완전히 재구축
"""
import logging
import base64
import zlib
import struct
import binascii
import io
import zipfile

logger = logging.getLogger(__name__)


class DataMashupRebuilder:
    """
    DataMashup 내부 ZIP 구조를 완전히 재구축하는 클래스
    """

    def rebuild_datamashup(self, original_base64: str, new_url: str) -> str:
        """
        DataMashup을 수동으로 파싱하여 재구축

        DataMashup은 표준 ZIP이 아닌 특수 구조이므로 수동 파싱 필요

        Args:
            original_base64: 원본 DataMashup Base64 데이터
            new_url: 새로운 OData URL

        Returns:
            새로운 DataMashup Base64 데이터
        """
        try:
            import re

            # Base64 디코딩
            decoded = base64.b64decode(original_base64)

            # Section1.m 찾기
            section_index = decoded.find(b'Section1.m')
            if section_index == -1:
                logger.warning("Section1.m not found in DataMashup")
                return original_base64

            # PK 헤더 찾기
            pk_offset = -1
            for i in range(max(0, section_index - 100), section_index):
                if decoded[i:i+4] == b'PK\x03\x04':
                    pk_offset = i
                    break

            if pk_offset == -1:
                logger.error("PK signature not found")
                return original_base64

            # ZIP 헤더 파싱
            compression = struct.unpack('<H', decoded[pk_offset+8:pk_offset+10])[0]
            compressed_size = struct.unpack('<I', decoded[pk_offset+18:pk_offset+22])[0]
            filename_len = struct.unpack('<H', decoded[pk_offset+26:pk_offset+28])[0]
            extra_len = struct.unpack('<H', decoded[pk_offset+28:pk_offset+30])[0]

            data_start = pk_offset + 30 + filename_len + extra_len
            compressed_data = decoded[data_start:data_start+compressed_size]

            # 압축 해제
            if compression == 8:
                m_code = zlib.decompress(compressed_data, -15).decode('utf-8')
            else:
                m_code = compressed_data.decode('utf-8')

            logger.debug(f"Original M code:\\n{m_code}")

            # URL 교체
            pattern = r'OData\.Feed\("([^"]+)"'
            url_match = re.search(pattern, m_code)

            if not url_match:
                logger.warning("No OData.Feed URL found")
                return original_base64

            old_url = url_match.group(1)
            new_m_code = m_code.replace(old_url, new_url)

            logger.debug(f"Updated M code:\\n{new_m_code}")
            logger.info(f"URL changed: {old_url} → {new_url}")

            # 재압축
            new_m_code_bytes = new_m_code.encode('utf-8')
            if compression == 8:
                new_compressed = zlib.compress(new_m_code_bytes)[2:-4]
            else:
                new_compressed = new_m_code_bytes

            # 크기 및 CRC32 계산
            new_compressed_size = len(new_compressed)
            new_uncompressed_size = len(new_m_code_bytes)
            new_crc32 = binascii.crc32(new_m_code_bytes) & 0xffffffff
            size_diff = new_compressed_size - compressed_size

            # 새 DataMashup 구성
            new_decoded = bytearray(decoded[:data_start]) + new_compressed + decoded[data_start + compressed_size:]

            # 로컬 파일 헤더 업데이트
            new_decoded[pk_offset+14:pk_offset+18] = struct.pack('<I', new_crc32)
            new_decoded[pk_offset+18:pk_offset+22] = struct.pack('<I', new_compressed_size)
            new_decoded[pk_offset+22:pk_offset+26] = struct.pack('<I', new_uncompressed_size)

            # Central Directory 업데이트
            cd_sig = b'PK\x01\x02'
            cd_offset = new_decoded.find(cd_sig, data_start + new_compressed_size)

            if cd_offset != -1:
                # Section1.m의 CD 항목 찾기
                current = cd_offset
                while current < len(new_decoded) - 46:
                    if new_decoded[current:current+4] == cd_sig:
                        cd_fname_len = struct.unpack('<H', new_decoded[current+28:current+30])[0]
                        cd_extra = struct.unpack('<H', new_decoded[current+30:current+32])[0]
                        cd_comment = struct.unpack('<H', new_decoded[current+32:current+34])[0]

                        fname = new_decoded[current+46:current+46+cd_fname_len]

                        if b'Section1.m' in fname:
                            # CD 항목 업데이트
                            new_decoded[current+16:current+20] = struct.pack('<I', new_crc32)
                            new_decoded[current+20:current+24] = struct.pack('<I', new_compressed_size)
                            new_decoded[current+24:current+28] = struct.pack('<I', new_uncompressed_size)
                            logger.debug(f"Updated Central Directory entry for Section1.m")
                            break

                        current += 46 + cd_fname_len + cd_extra + cd_comment
                    else:
                        break

            # EOCD는 원본과 동일하게 유지 (0으로 설정된 것이 Excel의 특수 처리)
            # Excel이 DataMashup 내부 ZIP을 특별하게 처리하므로 EOCD 업데이트 불필요

            # Base64 인코딩
            new_base64 = base64.b64encode(bytes(new_decoded)).decode('ascii')

            logger.info(f"Rebuilt DataMashup: {len(decoded)} → {len(new_decoded)} bytes")

            return new_base64

        except Exception as e:
            logger.error(f"Error rebuilding DataMashup: {str(e)}", exc_info=True)
            raise


    def modify_item1_xml(self, item1_content: bytes, new_url: str) -> bytes:
        """
        item1.xml 파일의 DataMashup 수정

        Args:
            item1_content: 원본 item1.xml 바이너리 데이터
            new_url: 새로운 OData URL

        Returns:
            수정된 item1.xml 바이너리 데이터
        """
        try:
            import re

            # UTF-16LE 디코딩
            text = item1_content.decode('utf-16-le', errors='ignore')

            # DataMashup 태그 찾기
            match = re.search(r'<DataMashup[^>]*>([^<]+)</DataMashup>', text)

            if not match:
                logger.warning("No DataMashup found in item1.xml")
                return item1_content

            # 원본 Base64 데이터
            original_base64 = match.group(1).replace(' ', '').replace('\n', '').replace('\r', '')

            # DataMashup 재구축
            new_base64 = self.rebuild_datamashup(original_base64, new_url)

            # XML 재구성
            new_text = text[:match.start(1)] + new_base64 + text[match.end(1):]

            # UTF-16LE로 인코딩
            new_content = new_text.encode('utf-16-le')

            return new_content

        except Exception as e:
            logger.error(f"Error modifying item1.xml: {str(e)}", exc_info=True)
            raise
