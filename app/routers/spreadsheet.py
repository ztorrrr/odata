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


@router.get("/connection-config")
async def get_connection_config(
    spreadsheet_id: str = Query(..., description="Google Spreadsheet ID"),
    view_id: Optional[str] = Query(None, description="BigQuery View ID"),
    current_user: str = Depends(get_current_user_with_header_token)
):
    """
    스프레드시트와 BigQuery 연동 설정 정보 반환
    """
    try:
        connector = get_spreadsheet_connector()
        config = connector.get_connected_sheets_config(
            view_id=view_id,
            spreadsheet_id=spreadsheet_id
        )

        return JSONResponse(content=config)
    except Exception as e:
        logger.error(f"Error getting connection config: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/setup-guide")
async def get_setup_guide(
    spreadsheet_id: str = Query(..., description="Google Spreadsheet ID"),
    view_id: Optional[str] = Query(None, description="BigQuery View ID"),
    current_user: str = Depends(get_current_user_with_header_token)
):
    """
    Connected Sheets 설정 가이드 제공
    """
    try:
        connector = get_spreadsheet_connector()
        guide = connector.create_data_source_for_sheet(
            spreadsheet_id=spreadsheet_id,
            view_id=view_id
        )

        return JSONResponse(content=guide)
    except Exception as e:
        logger.error(f"Error getting setup guide: {e}")
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


@router.post("/setup-spreadsheet/{spreadsheet_id}")
async def setup_spreadsheet_connection(
    spreadsheet_id: str,
    sample_size: int = Query(100, description="Number of sample rows"),
    source_table: Optional[str] = Query(None, description="Source table name"),
    current_user: str = Depends(get_current_user_with_header_token)
):
    """
    스프레드시트 연동을 위한 전체 설정 프로세스
    1. 샘플 View 생성
    2. 연동 가이드 제공
    """
    try:
        connector = get_spreadsheet_connector()

        # 1. 샘플 View 생성
        view_id = connector.create_sample_view(
            source_table=source_table,
            sample_size=sample_size,
            force_recreate=False
        )

        # 2. 연동 가이드 생성
        guide = connector.create_data_source_for_sheet(
            spreadsheet_id=spreadsheet_id,
            view_id=view_id
        )

        # 3. 샘플 데이터 미리보기 (5개 행)
        sample_data = connector.get_sample_data(view_id=view_id, limit=5)

        return JSONResponse(
            content={
                "success": True,
                "view_id": view_id,
                "sample_size": sample_size,
                "setup_guide": guide,
                "data_preview": sample_data,
                "message": "Sample view created. Follow the setup guide to connect to your spreadsheet."
            }
        )
    except Exception as e:
        logger.error(f"Error setting up spreadsheet connection: {e}")
        raise HTTPException(status_code=500, detail=str(e))