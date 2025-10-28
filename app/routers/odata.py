"""
OData v4 Router
BigQuery 데이터를 OData v4 프로토콜로 제공하는 라우터
"""
import logging
import os
from datetime import datetime
from typing import Optional

import pandas as pd
from fastapi import APIRouter, BackgroundTasks, Depends, Query, Request
from fastapi.responses import FileResponse, JSONResponse, Response, StreamingResponse

from app.services.bigquery_service import get_bigquery_service
from app.services.odata_metadata import ODataMetadataGenerator
from app.services.odata_query_parser import ODataQueryParser
from app.services.excel_com_generator import create_excel_with_odata_com
from app.utils.setting import get_config
from app.utils.auth import get_current_user, get_current_user_with_header_token

# 설정 및 로거
config = get_config()
logger = logging.getLogger(__name__)

# 라우터 생성
router = APIRouter(
    prefix="/odata",
    tags=["OData"],
    responses={
        404: {"description": "Not found"},
        500: {"description": "Internal server error"},
    },
)


@router.get("/")
async def get_service_document(request: Request):
    """
    OData Service Document
    사용 가능한 Entity Sets 목록을 반환합니다.
    """
    base_url = str(request.base_url).rstrip('/')

    return {
        "@odata.context": f"{base_url}/odata/$metadata",
        "value": [
            {
                "name": config.BIGQUERY_TABLE_NAME,
                "kind": "EntitySet",
                "url": config.BIGQUERY_TABLE_NAME
            }
        ]
    }


@router.get("/$metadata")
async def get_metadata():
    """
    OData Metadata Document
    서비스의 데이터 모델을 XML 형식으로 반환합니다.
    """
    try:
        # OData 메타데이터 생성
        metadata_gen = ODataMetadataGenerator()
        metadata_xml = metadata_gen.generate_metadata()

        return Response(
            content=metadata_xml,
            media_type="application/xml",
            headers={
                "OData-Version": "4.0",
            }
        )

    except Exception as e:
        logger.error(f"Error generating metadata: {str(e)}", exc_info=True)
        return Response(
            content=f"<error>{str(e)}</error>",
            media_type="application/xml",
            status_code=500
        )


@router.get(f"/{config.BIGQUERY_TABLE_NAME}")
async def get_entity_set(
    request: Request,
    username: str = Depends(get_current_user_with_header_token),
    filter: Optional[str] = Query(None, alias="$filter", description="OData filter expression"),
    select: Optional[str] = Query(None, alias="$select", description="Comma-separated list of properties"),
    orderby: Optional[str] = Query(None, alias="$orderby", description="Order by expression"),
    top: Optional[int] = Query(config.ODATA_MAX_PAGE_SIZE, alias="$top", description="Maximum number of records"),
    skip: Optional[int] = Query(0, alias="$skip", description="Number of records to skip"),
    count: Optional[bool] = Query(False, alias="$count", description="Include total count"),
):
    """
    Entity Set 조회
    BigQuery 테이블의 데이터를 OData 형식으로 반환합니다.

    지원하는 쿼리 옵션:
    - $filter: 필터 조건 (예: Name eq 'John')
    - $select: 선택할 필드 (예: Name,Age)
    - $orderby: 정렬 조건 (예: Name asc, Age desc)
    - $top: 최대 레코드 수
    - $skip: 건너뛸 레코드 수
    - $count: 전체 개수 포함 여부
    """
    try:
        bq_service = get_bigquery_service()
        parser = ODataQueryParser()

        # 쿼리 파라미터 파싱
        query_params = {
            'filter': filter,
            'select': select,
            'orderby': orderby,
            'top': min(top, config.ODATA_MAX_PAGE_SIZE) if top else config.ODATA_MAX_PAGE_SIZE,
            'skip': skip,
            'count': count
        }

        # BigQuery 쿼리 실행
        result = bq_service.query_table(
            parser=parser,
            **query_params
        )

        # 응답 데이터 구성
        base_url = str(request.base_url).rstrip('/')
        response_data = {
            "@odata.context": f"{base_url}/odata/$metadata#{config.BIGQUERY_TABLE_NAME}",
            "value": result['rows']
        }

        # count 요청 시 전체 개수 포함
        if count and 'total_count' in result:
            response_data["@odata.count"] = result['total_count']

        # 페이징 처리 - top과 정확히 일치하는 경우만 nextLink 추가
        if result['row_count'] == query_params['top']:
            next_skip = skip + query_params['top']
            query_parts = []
            if filter:
                query_parts.append(f"$filter={filter}")
            if select:
                query_parts.append(f"$select={select}")
            if orderby:
                query_parts.append(f"$orderby={orderby}")
            query_parts.append(f"$top={query_params['top']}")
            query_parts.append(f"$skip={next_skip}")
            if count:
                query_parts.append("$count=true")

            next_link = f"{base_url}/odata/{config.BIGQUERY_TABLE_NAME}?" + "&".join(query_parts)
            response_data["@odata.nextLink"] = next_link

        return JSONResponse(
            content=response_data,
            headers={
                "OData-Version": "4.0",
            }
        )

    except Exception as e:
        logger.error(f"Error querying entity set: {str(e)}", exc_info=True)
        return JSONResponse(
            status_code=500,
            content={
                "error": {
                    "code": "InternalServerError",
                    "message": str(e)
                }
            }
        )


@router.get(f"/{config.BIGQUERY_TABLE_NAME}/$count")
async def get_count(
    request: Request,
    username: str = Depends(get_current_user_with_header_token),
    filter: Optional[str] = Query(None, alias="$filter", description="OData filter expression"),
):
    """
    레코드 개수 조회
    필터 조건에 맞는 전체 레코드 수를 반환합니다.
    """
    try:
        bq_service = get_bigquery_service()
        parser = ODataQueryParser()

        # count만 조회
        result = bq_service.query_table(
            parser=parser,
            filter=filter,
            count_only=True
        )

        # 텍스트로 개수만 반환 (OData 표준)
        return Response(
            content=str(result.get('total_count', 0)),
            media_type="text/plain",
            headers={
                "OData-Version": "4.0",
            }
        )

    except Exception as e:
        logger.error(f"Error getting count: {str(e)}", exc_info=True)
        return Response(
            content=str(0),
            media_type="text/plain",
            status_code=500
        )


@router.get("/health")
async def health_check():
    """
    헬스 체크 엔드포인트
    서비스 상태를 확인합니다.
    """
    try:
        bq_service = get_bigquery_service()
        table_info = bq_service.get_table_info()

        return {
            "status": "healthy",
            "service": "OData v4 Service",
            "table": table_info.get("table_id") if table_info else "Not found",
            "rows": table_info.get("num_rows") if table_info else 0
        }
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}")
        return JSONResponse(
            status_code=503,
            content={
                "status": "unhealthy",
                "error": str(e)
            }
        )


@router.get("/debug/routes")
async def debug_routes():
    """디버그: 등록된 모든 라우트 확인"""
    import sys
    routes_info = []

    # 현재 모듈 파일 경로
    module_file = sys.modules[__name__].__file__

    # 이 라우터의 모든 라우트
    for route in router.routes:
        if hasattr(route, 'path'):
            routes_info.append({
                "path": route.path,
                "methods": list(route.methods) if hasattr(route, 'methods') else []
            })

    return {
        "module_file": module_file,
        "total_routes": len(routes_info),
        "routes": routes_info
    }


@router.get(f"/{config.BIGQUERY_TABLE_NAME}/export")
async def export_to_csv(
    request: Request,
    username: str = Depends(get_current_user_with_header_token),
    filter: Optional[str] = Query(None, alias="$filter", description="OData filter expression"),
    select: Optional[str] = Query(None, alias="$select", description="Comma-separated list of properties"),
    orderby: Optional[str] = Query(None, alias="$orderby", description="Order by expression"),
    top: Optional[int] = Query(None, alias="$top", description="Maximum number of records"),
    skip: Optional[int] = Query(0, alias="$skip", description="Number of records to skip"),
):
    """
    CSV 파일로 내보내기
    쿼리 결과를 CSV 파일로 다운로드합니다.
    """
    try:
        bq_service = get_bigquery_service()
        parser = ODataQueryParser()

        # 최대 행 수 제한 (CSV 다운로드용)
        max_rows = min(top, config.CSV_MAX_ROWS) if top else config.CSV_MAX_ROWS

        # BigQuery 쿼리 실행
        result = bq_service.query_table(
            parser=parser,
            filter=filter,
            select=select,
            orderby=orderby,
            top=max_rows,
            skip=skip,
            count=False
        )

        # pandas DataFrame으로 변환
        df = pd.DataFrame(result['rows'])

        # CSV 문자열로 변환 (UTF-8 BOM 추가)
        from io import StringIO
        csv_buffer = StringIO()
        csv_buffer.write('\ufeff')  # UTF-8 BOM
        df.to_csv(csv_buffer, index=False, encoding='utf-8')
        csv_content = csv_buffer.getvalue()

        # 파일명 생성
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{config.BIGQUERY_TABLE_NAME}_{timestamp}.csv"

        return Response(
            content=csv_content,
            media_type="text/csv; charset=utf-8",
            headers={
                "Content-Disposition": f"attachment; filename={filename}",
            }
        )

    except Exception as e:
        logger.error(f"Error exporting to CSV: {str(e)}", exc_info=True)
        return JSONResponse(
            status_code=500,
            content={
                "error": {
                    "code": "InternalServerError",
                    "message": str(e)
                }
            }
        )


@router.get(f"/{config.BIGQUERY_TABLE_NAME}/excel-com-webapi-key")
async def get_excel_with_webapi_auth(
    request: Request,
    background_tasks: BackgroundTasks,
    filter: Optional[str] = Query(None, alias="$filter", description="OData filter expression"),
    select: Optional[str] = Query(None, alias="$select", description="Comma-separated list of properties"),
    orderby: Optional[str] = Query(None, alias="$orderby", description="Order by expression"),
    table_name: Optional[str] = Query("Data", description="Excel table name"),
    api_key_name: Optional[str] = Query("Authorization", description="API Key name for Web API auth")
):
    """
    Windows COM을 사용한 Excel 파일 생성 (Web API 인증 템플릿)

    Web API 인증 방식을 사용하는 Excel 템플릿 파일을 생성합니다.
    이 엔드포인트는 인증 없이 접근 가능합니다 (템플릿 다운로드용).

    실제 데이터 보안:
    - 다운로드된 Excel 파일에는 데이터가 없음 (템플릿만 제공)
    - Excel에서 데이터 새로고침 시 Web API 키 입력 필요
    - 실제 데이터는 인증 후에만 접근 가능

    Excel에서 인증 방법:
    - 데이터 새로고침 시 "웹 API" 탭 선택
    - 키 입력란에 "Bearer <your_api_token>" 형식으로 입력
    - 예: Bearer tok_abc123xyz

    특징:
    - Windows 환경에서만 작동
    - Excel이 설치되어 있어야 함
    - Web API 인증 설정이 포함된 Excel 템플릿 생성

    지원하는 파라미터:
    - $filter: 필터 조건 (예: Media eq 'Naver')
    - $select: 선택할 필드 (예: Date,Campaign,Clicks)
    - $orderby: 정렬 조건 (예: Date desc)
    - table_name: Excel 테이블 이름 (기본: Data)
    - api_key_name: API 키 이름 (기본: Authorization)
    """
    try:
        from app.services.excel_com_generator import create_excel_with_webapi_auth_com

        # OData URL 구성
        base_url = str(request.base_url).rstrip('/')
        odata_url = f"{base_url}/odata/{config.BIGQUERY_TABLE_NAME}"

        # 쿼리 파라미터 추가
        query_params = []
        if filter:
            query_params.append(f"$filter={filter}")
        if select:
            query_params.append(f"$select={select}")
        if orderby:
            query_params.append(f"$orderby={orderby}")

        if query_params:
            odata_url += "?" + "&".join(query_params)

        # Web API 인증을 사용하여 Excel 파일 생성
        output_path = create_excel_with_webapi_auth_com(
            odata_url=odata_url,
            table_name=table_name,
            api_key_name=api_key_name,
            skip_data_load=True  # 인증 필요하므로 데이터 로드 건너뜀
        )

        # 파일명 생성
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{config.BIGQUERY_TABLE_NAME}_webapi_{timestamp}.xlsx"

        # 임시 파일 삭제 작업을 background task로 추가
        background_tasks.add_task(os.unlink, output_path)

        # FileResponse로 반환
        return FileResponse(
            path=output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except ImportError as e:
        logger.error(f"pywin32 not installed or not available on this platform: {e}")
        return JSONResponse(
            status_code=501,
            content={
                "error": {
                    "code": "NotImplemented",
                    "message": "Windows COM support is not available. This endpoint requires Windows with Excel installed.",
                    "details": str(e)
                }
            }
        )

    except Exception as e:
        logger.error(f"Error generating Excel with Web API auth: {str(e)}", exc_info=True)
        return JSONResponse(
            status_code=500,
            content={
                "error": {
                    "code": "InternalServerError",
                    "message": str(e)
                }
            }
        )


@router.get(f"/{config.BIGQUERY_TABLE_NAME}/excel-com")
async def get_excel_with_com(
    request: Request,
    background_tasks: BackgroundTasks,
    filter: Optional[str] = Query(None, alias="$filter", description="OData filter expression"),
    select: Optional[str] = Query(None, alias="$select", description="Comma-separated list of properties"),
    orderby: Optional[str] = Query(None, alias="$orderby", description="Order by expression"),
    table_name: Optional[str] = Query("Data", description="Excel table name")
):
    """
    Windows COM을 사용한 Excel 파일 생성 (Basic Auth 템플릿)

    HTTP Basic Authentication (ID/PW) 방식을 사용하는 Excel 템플릿 파일을 생성합니다.
    이 엔드포인트는 인증 없이 접근 가능합니다 (템플릿 다운로드용).

    실제 데이터 보안:
    - 다운로드된 Excel 파일에는 데이터가 없음 (템플릿만 제공)
    - Excel에서 데이터 새로고침 시 "기본" 탭에서 ID/PW 입력 필요
    - 실제 데이터는 인증 후에만 접근 가능

    Excel에서 인증 방법:
    - 데이터 새로고침 시 "기본" 탭 선택
    - 사용자명과 비밀번호 입력
    - Basic Auth로 서버에 인증

    특징:
    - Windows 환경에서만 작동
    - Excel이 설치되어 있어야 함
    - Basic Auth 방식의 OData 연결 템플릿 생성

    지원하는 파라미터:
    - $filter: 필터 조건 (예: Media eq 'Naver')
    - $select: 선택할 필드 (예: Date,Campaign,Clicks)
    - $orderby: 정렬 조건 (예: Date desc)
    - table_name: Excel 테이블 이름 (기본: Data)
    """
    try:
        # OData URL 구성
        base_url = str(request.base_url).rstrip('/')
        odata_url = f"{base_url}/odata/{config.BIGQUERY_TABLE_NAME}"

        # 쿼리 파라미터 추가
        query_params = []
        if filter:
            query_params.append(f"$filter={filter}")
        if select:
            query_params.append(f"$select={select}")
        if orderby:
            query_params.append(f"$orderby={orderby}")

        if query_params:
            odata_url += "?" + "&".join(query_params)

        # COM을 사용하여 Excel 파일 생성
        output_path = create_excel_with_odata_com(
            odata_url=odata_url,
            table_name=table_name
        )

        # 파일명 생성
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{config.BIGQUERY_TABLE_NAME}_odata_{timestamp}.xlsx"

        # 임시 파일 삭제 작업을 background task로 추가
        background_tasks.add_task(os.unlink, output_path)

        # FileResponse로 반환
        return FileResponse(
            path=output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except ImportError as e:
        logger.error(f"pywin32 not installed or not available on this platform: {e}")
        return JSONResponse(
            status_code=501,
            content={
                "error": {
                    "code": "NotImplemented",
                    "message": "Windows COM support is not available. This endpoint requires Windows with Excel installed.",
                    "details": str(e)
                }
            }
        )

    except Exception as e:
        logger.error(f"Error generating Excel with COM: {str(e)}", exc_info=True)
        return JSONResponse(
            status_code=500,
            content={
                "error": {
                    "code": "InternalServerError",
                    "message": str(e)
                }
            }
        )