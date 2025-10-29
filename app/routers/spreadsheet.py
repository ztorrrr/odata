"""
Google Spreadsheet integration endpoints
"""
import logging
from typing import Optional
from fastapi import APIRouter, HTTPException, Query, Depends
from fastapi.responses import JSONResponse

from app.services.spreadsheet_connector import get_spreadsheet_connector
from app.utils.auth import get_current_user_with_header_token

logger = logging.getLogger(__name__)

router = APIRouter(
    prefix="/spreadsheet",
    tags=["spreadsheet"],
    responses={404: {"description": "Not found"}},
)


@router.post("/create-sample-view")
async def create_sample_view(
    sample_size: int = Query(100, description="Number of sample rows"),
    source_table: Optional[str] = Query(None, description="Source table name"),
    view_name: Optional[str] = Query(None, description="Custom view name"),
    force_recreate: bool = Query(False, description="Recreate if exists"),
    current_user: str = Depends(get_current_user_with_header_token)
):
    """
    BigQuery 테이블에서 샘플 데이터 View 생성
    """
    try:
        connector = get_spreadsheet_connector()
        view_id = connector.create_sample_view(
            source_table=source_table,
            view_name=view_name,
            sample_size=sample_size,
            force_recreate=force_recreate
        )

        return JSONResponse(
            content={
                "success": True,
                "view_id": view_id,
                "sample_size": sample_size,
                "message": f"Sample view created successfully with {sample_size} rows"
            }
        )
    except Exception as e:
        logger.error(f"Error creating sample view: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/sample-data")
async def get_sample_data(
    view_id: Optional[str] = Query(None, description="BigQuery View ID"),
    limit: int = Query(10, description="Number of rows to preview"),
    current_user: str = Depends(get_current_user_with_header_token)
):
    """
    샘플 View 데이터 미리보기
    """
    try:
        connector = get_spreadsheet_connector()
        data = connector.get_sample_data(
            view_id=view_id,
            limit=limit
        )

        return JSONResponse(
            content={
                "rows": data,
                "count": len(data),
                "view_id": view_id
            }
        )
    except Exception as e:
        logger.error(f"Error getting sample data: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/modify-view-test")
async def modify_view_for_test(
    column_name: str = Query("Type", description="Column to modify"),
    suffix: str = Query("_테스트", description="Suffix to add"),
    view_name: Optional[str] = Query(None, description="View name"),
    current_user: str = Depends(get_current_user_with_header_token)
):
    """
    View를 수정하여 실시간 동기화 테스트
    Type 컬럼에 '_테스트' suffix 추가
    """
    try:
        connector = get_spreadsheet_connector()
        view_id = connector.modify_view_with_test_suffix(
            view_name=view_name,
            column_to_modify=column_name,
            suffix=suffix
        )

        # 수정 후 샘플 데이터 확인
        sample_data = connector.get_sample_data(view_id=view_id, limit=3)

        return JSONResponse(
            content={
                "success": True,
                "view_id": view_id,
                "modified_column": column_name,
                "suffix_added": suffix,
                "sample_data": sample_data,
                "message": f"View modified successfully. Column '{column_name}' now has suffix '{suffix}'. Please refresh your Google Sheet to see changes."
            }
        )
    except Exception as e:
        logger.error(f"Error modifying view: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/restore-view")
async def restore_original_view(
    view_name: Optional[str] = Query(None, description="View name"),
    sample_size: int = Query(100, description="Sample size"),
    current_user: str = Depends(get_current_user_with_header_token)
):
    """
    View를 원본 상태로 복원
    """
    try:
        connector = get_spreadsheet_connector()
        view_id = connector.restore_original_view(
            view_name=view_name,
            sample_size=sample_size
        )

        # 복원 후 샘플 데이터 확인
        sample_data = connector.get_sample_data(view_id=view_id, limit=3)

        return JSONResponse(
            content={
                "success": True,
                "view_id": view_id,
                "sample_size": sample_size,
                "sample_data": sample_data,
                "message": "View restored to original state. Please refresh your Google Sheet to see changes."
            }
        )
    except Exception as e:
        logger.error(f"Error restoring view: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.api_route("/create-connected-bigquery", methods=["GET", "POST"])
async def create_connected_bigquery_spreadsheet(
    title: Optional[str] = Query(None, description="Spreadsheet title (auto-generated if not provided: bigquery_connector_YYMMDD_HHMMSS)"),
    view_id: Optional[str] = Query(None, description="BigQuery View ID"),
    folder_id: Optional[str] = Query(None, description="Google Drive folder ID"),
    folder_name: str = Query("odata_test", description="Folder name to search (used if folder_id is not provided)"),
    current_user: str = Depends(get_current_user_with_header_token)
):
    """
    BigQuery Connected Sheets를 생성 (네이티브 연결)

    이 방식은 Apps Script 없이도 BigQuery 데이터를 직접 조회하고
    새로고침할 수 있는 네이티브 연결을 생성합니다.

    특징:
    - Google Sheets의 네이티브 Connected Sheets 기능 사용
    - 스프레드시트 UI에서 직접 데이터 새로고침 가능
    - 수동 설정 없이 즉시 사용 가능
    - Apps Script 설정 불필요
    - 파일명 자동 생성 (timestamp 포함)
    - "odata_test" 폴더에 자동 저장
    """
    try:
        connector = get_spreadsheet_connector()

        result = connector.create_spreadsheet_with_connected_bigquery(
            spreadsheet_title=title,
            view_id=view_id,
            folder_id=folder_id,
            folder_name=folder_name
        )

        return JSONResponse(content=result)

    except Exception as e:
        logger.error(f"Error creating Connected Sheets with BigQuery: {e}")
        raise HTTPException(status_code=500, detail=str(e))