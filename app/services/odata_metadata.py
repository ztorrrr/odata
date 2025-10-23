"""
OData v4 Metadata generator
"""
import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import List, Dict, Any

from google.cloud import bigquery

from app.utils.setting import get_config
from app.services.bigquery_service import get_bigquery_service

# BigQuery 타입을 OData EDM 타입으로 매핑
TYPE_MAPPING = {
    "STRING": "Edm.String",
    "INT64": "Edm.Int64",
    "INTEGER": "Edm.Int64",
    "FLOAT64": "Edm.Double",
    "FLOAT": "Edm.Double",
    "NUMERIC": "Edm.Decimal",
    "DECIMAL": "Edm.Decimal",
    "BOOLEAN": "Edm.Boolean",
    "BOOL": "Edm.Boolean",
    "TIMESTAMP": "Edm.DateTimeOffset",
    "DATETIME": "Edm.DateTimeOffset",
    "DATE": "Edm.Date",
    "TIME": "Edm.TimeOfDay",
    "BYTES": "Edm.Binary",
}


class ODataMetadataGenerator:
    """OData v4 메타데이터 생성기"""

    def __init__(self):
        self.config = get_config()
        self.bq_service = get_bigquery_service()
        self.namespace = "OData.Service"
        self.container_name = "DefaultContainer"

    def generate_metadata(self) -> str:
        """
        OData v4 메타데이터 XML 생성

        Returns:
            메타데이터 XML 문자열
        """
        # 루트 엘리먼트 생성
        root = ET.Element("edmx:Edmx", {
            "xmlns:edmx": "http://docs.oasis-open.org/odata/ns/edmx",
            "Version": "4.0"
        })

        # DataServices
        data_services = ET.SubElement(root, "edmx:DataServices")

        # Schema
        schema = ET.SubElement(data_services, "Schema", {
            "xmlns": "http://docs.oasis-open.org/odata/ns/edm",
            "Namespace": self.namespace
        })

        # EntityType 생성
        entity_type = self._create_entity_type(schema)

        # EntityContainer 생성
        container = ET.SubElement(schema, "EntityContainer", {
            "Name": self.container_name
        })

        # EntitySet 생성
        ET.SubElement(container, "EntitySet", {
            "Name": self.config.BIGQUERY_TABLE_NAME,
            "EntityType": f"{self.namespace}.{self.config.BIGQUERY_TABLE_NAME}"
        })

        # XML을 문자열로 변환 (pretty print)
        rough_string = ET.tostring(root, encoding="unicode")
        reparsed = minidom.parseString(rough_string)

        # XML 선언과 함께 반환
        return reparsed.toprettyxml(indent="  ")

    def _create_entity_type(self, parent: ET.Element) -> ET.Element:
        """
        BigQuery 스키마를 기반으로 EntityType 생성

        Args:
            parent: 부모 엘리먼트 (Schema)

        Returns:
            EntityType 엘리먼트
        """
        # BigQuery 테이블 스키마 가져오기
        schema_fields = self.bq_service.get_table_schema()

        # EntityType 엘리먼트 생성
        entity_type = ET.SubElement(parent, "EntityType", {
            "Name": self.config.BIGQUERY_TABLE_NAME
        })

        # 첫 번째 필드를 Key로 사용 (또는 특정 필드 지정)
        key_field = None
        if schema_fields:
            # 'id' 필드가 있으면 사용, 없으면 첫 번째 필드 사용
            for field in schema_fields:
                if field.name.lower() in ['id', 'key', 'code']:
                    key_field = field.name
                    break
            if not key_field:
                key_field = schema_fields[0].name

            # Key 엘리먼트 생성
            key = ET.SubElement(entity_type, "Key")
            ET.SubElement(key, "PropertyRef", {"Name": key_field})

        # 각 필드를 Property로 추가
        for field in schema_fields:
            self._add_property(entity_type, field)

        return entity_type

    def _add_property(self, parent: ET.Element, field: bigquery.SchemaField):
        """
        BigQuery 필드를 OData Property로 추가

        Args:
            parent: 부모 엘리먼트 (EntityType)
            field: BigQuery 스키마 필드
        """
        # OData EDM 타입으로 매핑
        edm_type = TYPE_MAPPING.get(field.field_type, "Edm.String")

        # Property 속성 설정
        attributes = {
            "Name": field.name,
            "Type": edm_type
        }

        # Nullable 설정 (REQUIRED가 아닌 경우)
        if field.mode != "REQUIRED":
            attributes["Nullable"] = "true"
        else:
            attributes["Nullable"] = "false"

        # REPEATED 필드는 Collection으로 처리
        if field.mode == "REPEATED":
            attributes["Type"] = f"Collection({edm_type})"

        # Property 엘리먼트 생성
        ET.SubElement(parent, "Property", attributes)

    def get_service_document(self) -> Dict[str, Any]:
        """
        OData Service Document 생성

        Returns:
            Service Document (JSON)
        """
        base_url = f"http://{self.config.HOST}:{self.config.PORT}/odata"

        return {
            "@odata.context": f"{base_url}/$metadata",
            "value": [
                {
                    "name": self.config.BIGQUERY_TABLE_NAME,
                    "kind": "EntitySet",
                    "url": self.config.BIGQUERY_TABLE_NAME
                }
            ]
        }