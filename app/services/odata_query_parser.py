"""
OData v4 query parameter parser
"""
import re
from typing import Optional, List, Dict, Any
from urllib.parse import unquote


class ODataQueryParser:
    """OData v4 쿼리 파라미터 파서"""

    def __init__(self):
        # OData 연산자를 SQL로 매핑
        self.operator_mapping = {
            "eq": "=",
            "ne": "!=",
            "gt": ">",
            "ge": ">=",
            "lt": "<",
            "le": "<=",
            "and": "AND",
            "or": "OR",
            "not": "NOT"
        }

    def parse_filter(self, filter_str: str) -> Optional[str]:
        """
        OData $filter를 SQL WHERE 절로 변환

        예시:
        - "name eq 'John'" -> "name = 'John'"
        - "age gt 20" -> "age > 20"
        - "name eq 'John' and age gt 20" -> "name = 'John' AND age > 20"

        Args:
            filter_str: OData $filter 문자열

        Returns:
            SQL WHERE 절 (WHERE 키워드 제외)
        """
        if not filter_str:
            return None

        # URL 디코딩
        filter_str = unquote(filter_str)

        # 기본 변환
        sql_filter = filter_str

        # OData 연산자를 SQL로 변경
        for odata_op, sql_op in self.operator_mapping.items():
            # 단어 경계를 사용하여 전체 단어만 매치
            pattern = r'\b' + re.escape(odata_op) + r'\b'
            sql_filter = re.sub(pattern, sql_op, sql_filter, flags=re.IGNORECASE)

        # 함수 처리
        sql_filter = self._parse_functions(sql_filter)

        # 필드명에 백틱 추가 (안전한 쿼리를 위해)
        sql_filter = self._add_field_backticks(sql_filter)

        return sql_filter

    def _parse_functions(self, filter_str: str) -> str:
        """
        OData 함수를 SQL로 변환

        지원 함수:
        - contains(field, 'value') -> field LIKE '%value%'
        - startswith(field, 'value') -> field LIKE 'value%'
        - endswith(field, 'value') -> field LIKE '%value'

        Args:
            filter_str: 필터 문자열

        Returns:
            변환된 문자열
        """
        # contains 함수
        pattern = r"contains\(([^,]+),\s*'([^']+)'\)"
        filter_str = re.sub(
            pattern,
            lambda m: f"{m.group(1)} LIKE '%{m.group(2)}%'",
            filter_str,
            flags=re.IGNORECASE
        )

        # startswith 함수
        pattern = r"startswith\(([^,]+),\s*'([^']+)'\)"
        filter_str = re.sub(
            pattern,
            lambda m: f"{m.group(1)} LIKE '{m.group(2)}%'",
            filter_str,
            flags=re.IGNORECASE
        )

        # endswith 함수
        pattern = r"endswith\(([^,]+),\s*'([^']+)'\)"
        filter_str = re.sub(
            pattern,
            lambda m: f"{m.group(1)} LIKE '%{m.group(2)}'",
            filter_str,
            flags=re.IGNORECASE
        )

        return filter_str

    def _add_field_backticks(self, filter_str: str) -> str:
        """
        필드명에 백틱 추가 (예약어 충돌 방지)

        Args:
            filter_str: 필터 문자열

        Returns:
            백틱이 추가된 문자열
        """
        # 간단한 필드명 패턴 (연산자 앞에 오는 단어)
        operators = ['=', '!=', '>', '>=', '<', '<=', 'LIKE']
        for op in operators:
            pattern = r'(\b\w+)\s*' + re.escape(op)
            filter_str = re.sub(
                pattern,
                lambda m: f"`{m.group(1)}` {op}",
                filter_str
            )

        return filter_str

    def parse_select(self, select_str: str) -> Optional[List[str]]:
        """
        OData $select를 필드 리스트로 변환

        예시:
        - "name,age" -> ["name", "age"]
        - "name, age, email" -> ["name", "age", "email"]

        Args:
            select_str: OData $select 문자열

        Returns:
            필드 리스트
        """
        if not select_str:
            return None

        # URL 디코딩
        select_str = unquote(select_str)

        # 쉼표로 분리하고 공백 제거
        fields = [field.strip() for field in select_str.split(",")]

        # 빈 문자열 제거
        fields = [field for field in fields if field]

        return fields if fields else None

    def parse_orderby(self, orderby_str: str) -> Optional[str]:
        """
        OData $orderby를 SQL ORDER BY 절로 변환

        예시:
        - "name" -> "`name`"
        - "name desc" -> "`name` DESC"
        - "name,age desc" -> "`name`, `age` DESC"

        Args:
            orderby_str: OData $orderby 문자열

        Returns:
            SQL ORDER BY 절 (ORDER BY 키워드 제외)
        """
        if not orderby_str:
            return None

        # URL 디코딩
        orderby_str = unquote(orderby_str)

        # 쉼표로 분리
        parts = [part.strip() for part in orderby_str.split(",")]

        # 각 부분 처리
        sql_parts = []
        for part in parts:
            if not part:
                continue

            # asc/desc 확인
            tokens = part.split()
            if len(tokens) == 1:
                # 기본값은 ASC
                sql_parts.append(f"`{tokens[0]}`")
            elif len(tokens) == 2:
                field, direction = tokens
                if direction.upper() in ["ASC", "DESC"]:
                    sql_parts.append(f"`{field}` {direction.upper()}")
                else:
                    sql_parts.append(f"`{part}`")
            else:
                sql_parts.append(f"`{part}`")

        return ", ".join(sql_parts) if sql_parts else None

    def parse_top(self, top_str: str) -> Optional[int]:
        """
        OData $top을 정수로 변환

        Args:
            top_str: OData $top 문자열

        Returns:
            정수 값
        """
        if not top_str:
            return None

        try:
            return int(top_str)
        except ValueError:
            return None

    def parse_skip(self, skip_str: str) -> Optional[int]:
        """
        OData $skip을 정수로 변환

        Args:
            skip_str: OData $skip 문자열

        Returns:
            정수 값
        """
        if not skip_str:
            return None

        try:
            return int(skip_str)
        except ValueError:
            return None

    def parse_count(self, count_str: str) -> bool:
        """
        OData $count를 boolean으로 변환

        Args:
            count_str: OData $count 문자열

        Returns:
            count 여부
        """
        if not count_str:
            return False

        return count_str.lower() in ["true", "1", "yes"]

    def parse_all(self, query_params: Dict[str, str]) -> Dict[str, Any]:
        """
        모든 OData 쿼리 파라미터 파싱

        Args:
            query_params: 쿼리 파라미터 딕셔너리

        Returns:
            파싱된 파라미터 딕셔너리
        """
        return {
            "filter": self.parse_filter(query_params.get("$filter")),
            "select": self.parse_select(query_params.get("$select")),
            "orderby": self.parse_orderby(query_params.get("$orderby")),
            "top": self.parse_top(query_params.get("$top")),
            "skip": self.parse_skip(query_params.get("$skip")),
            "count": self.parse_count(query_params.get("$count"))
        }