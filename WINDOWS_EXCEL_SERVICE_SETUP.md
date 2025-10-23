# Windows Excel Service 배포 가이드

## 개요

Windows COM을 사용하여 OData 연결이 포함된 실제 Excel 파일을 생성하는 독립 서비스입니다.

### 아키텍처

```
┌─────────────┐      ┌──────────────────┐      ┌─────────────────┐
│   Client    │─────→│  macOS Server    │─────→│ Windows Server  │
│  (Excel)    │←─────│  (Main OData)    │←─────│ (Excel COM)     │
└─────────────┘      └──────────────────┘      └─────────────────┘
                       Port 8888                 Port 8889

                       • OData endpoints         • /excel/odata
                       • Proxy endpoint          • Excel COM 자동화
                       • /excel-live
```

## 필수 요구사항

### Windows Server
- **OS**: Windows Server 2016 이상
- **Excel**: Microsoft Excel 2016 이상 (반드시 설치 필요)
- **Python**: 3.11 이상

### ⚠️ 중요 사항
- Microsoft는 서버 환경에서 Office 자동화를 공식적으로 권장하지 않습니다
- 프로덕션 환경에서는 동시 요청 수를 제한하는 것을 권장합니다
- Excel 프로세스가 종료되지 않는 경우를 대비한 모니터링 필요

## 설치 방법

### 1. Windows Server에 Python 및 uv 설치

Python 3.11 이상 설치 권장
ㅇ
uv 설치 (Windows PowerShell):
```powershell
# uv 설치
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# 또는 pip로 설치
pip install uv
```

### 2. 프로젝트 디렉토리 생성

```powershell
mkdir windows-excel-service
cd windows-excel-service
```

### 3. pyproject.toml 생성

프로젝트 루트에 `pyproject.toml` 파일 생성:

```toml
[project]
name = "windows-excel-service"
version = "1.0.0"
description = "Windows Excel Service for OData connections"
requires-python = ">=3.11"
dependencies = [
    "fastapi>=0.119.0",
    "uvicorn>=0.38.0",
    "pywin32>=308",
]
```

### 4. 의존성 설치

```powershell
uv sync
```

### 5. 서비스 파일 생성

아래 샘플 코드를 `windows_excel_service.py`로 저장:

```python
"""
Windows Excel Service

Windows Server에서만 실행되는 독립 서비스
Excel COM을 사용하여 OData 연결이 포함된 Excel 파일 생성

Requirements:
- Windows Server
- Microsoft Excel 2016+
- pip install fastapi uvicorn pywin32

Usage:
    python windows_excel_service.py
"""

import logging
import os
import sys
import platform
import time
import tempfile
from typing import Optional

from fastapi import FastAPI, Query, HTTPException
from fastapi.responses import FileResponse
import uvicorn

# Windows 환경 체크
if platform.system() != "Windows":
    print("ERROR: This service must run on Windows Server")
    print("Current platform:", platform.system())
    sys.exit(1)

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    print("ERROR: pywin32 is not installed")
    print("Please run: uv add pywin32")
    sys.exit(1)

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# FastAPI 앱
app = FastAPI(
    title="Windows Excel Service",
    description="COM-based Excel generation service for OData connections",
    version="1.0.0"
)


@app.get("/")
async def root():
    """서비스 상태 확인"""
    return {
        "service": "Windows Excel Service",
        "status": "running",
        "platform": platform.system(),
        "python_version": sys.version,
        "endpoints": [
            "/excel/odata",
            "/health"
        ]
    }


@app.get("/health")
async def health_check():
    """헬스 체크 - Excel COM 객체 생성 가능 여부 확인"""
    try:
        pythoncom.CoInitialize()
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        version = xl.Version
        xl.Quit()
        pythoncom.CoUninitialize()

        return {
            "status": "healthy",
            "excel_available": True,
            "excel_version": version
        }
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}", exc_info=True)
        return {
            "status": "unhealthy",
            "excel_available": False,
            "error": str(e)
        }


@app.get("/excel/odata")
async def make_odata_excel(
    odata_url: str = Query(..., description="OData endpoint URL"),
    sheet_name: str = Query("Data", description="Excel 시트 이름"),
    query_name: str = Query("ODataQuery", description="Power Query 이름"),
):
    """
    OData 연결이 포함된 Excel 파일 생성

    Parameters:
    - odata_url: OData 엔드포인트 URL (필수)
    - sheet_name: Excel 시트 이름 (기본: "Data")
    - query_name: Power Query 이름 (기본: "ODataQuery")

    Returns:
    - Excel 파일 (.xlsx)
    """
    logger.info(f"Excel generation requested: URL={odata_url}")

    xl = None
    wb = None
    temp_file_path = None

    try:
        # Excel COM 초기화 (STA 모드 필요)
        pythoncom.CoInitialize()

        # Excel 애플리케이션 시작
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        logger.info(f"Excel started (Version: {xl.Version})")

        # 새 워크북 생성
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = sheet_name

        # Power Query M 스크립트
        m_code = f'''let
    Source = OData.Feed("{odata_url}", null, [MoreColumns=true])
in
    Source'''.strip()

        # 쿼리 추가
        wb.Queries.Add(query_name, m_code)
        logger.info(f"Query added: {query_name}")

        # 쿼리를 시트 테이블로 로드
        conn = f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={query_name};Extended Properties=""'

        lo = ws.ListObjects.Add(
            SourceType=0,  # xlSrcExternal
            Source=conn,
            Destination=ws.Range("A1")
        )
        lo.Name = query_name
        qt = lo.QueryTable
        qt.RefreshStyle = 1
        qt.Refresh(False)

        logger.info("Query loaded and refreshed")

        # 임시 파일로 저장
        tmpdir = tempfile.mkdtemp()
        temp_file_path = os.path.join(tmpdir, f"{query_name}.xlsx")
        wb.SaveAs(temp_file_path)

        # 워크북 닫기
        wb.Close(SaveChanges=False)
        wb = None
        xl.Quit()
        xl = None
        pythoncom.CoUninitialize()

        time.sleep(0.2)  # 파일 핸들 정리

        logger.info(f"Excel file created: {temp_file_path}")

        return FileResponse(
            path=temp_file_path,
            filename=f"{query_name}.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            background=lambda: _cleanup_temp_file(temp_file_path, tmpdir)
        )

    except Exception as e:
        logger.error(f"Error: {str(e)}", exc_info=True)

        # 정리
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if xl:
            try:
                xl.Quit()
            except:
                pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
            except:
                pass

        raise HTTPException(
            status_code=500,
            detail={
                "error": "Excel generation failed",
                "message": str(e)
            }
        )


def _cleanup_temp_file(file_path: str, dir_path: str):
    """임시 파일 및 디렉토리 정리"""
    try:
        if os.path.exists(file_path):
            os.unlink(file_path)
        if os.path.exists(dir_path):
            os.rmdir(dir_path)
    except Exception as e:
        logger.warning(f"Cleanup failed: {str(e)}")


if __name__ == "__main__":
    HOST = os.getenv("HOST", "0.0.0.0")
    PORT = int(os.getenv("PORT", "8889"))

    print("\n" + "="*60)
    print("Windows Excel Service")
    print("="*60)
    print(f"Platform: {platform.system()}")
    print(f"Server: http://{HOST}:{PORT}")
    print("="*60 + "\n")

    uvicorn.run(app, host=HOST, port=PORT, log_level="info")
```

### 6. Excel 권한 설정

Excel COM 객체가 서비스 계정으로 실행되는 경우, 다음 폴더가 존재해야 합니다:

- 32-bit Excel: `C:\Windows\System32\config\systemprofile\Desktop`
- 64-bit Excel: `C:\Windows\SysWOW64\config\systemprofile\Desktop`

폴더 생성:
```powershell
New-Item -Path "C:\Windows\System32\config\systemprofile\Desktop" -ItemType Directory -Force
```

## 서비스 실행

### 방법 1: 직접 실행 (개발/테스트)

```powershell
uv run python windows_excel_service.py
```

기본 설정:
- Host: `0.0.0.0`
- Port: `8889`

### 방법 2: 환경 변수로 설정 변경

```powershell
$env:HOST = "127.0.0.1"
$env:PORT = "9000"
uv run python windows_excel_service.py
```

### 방법 3: Windows 서비스로 등록 (프로덕션)

NSSM (Non-Sucking Service Manager) 사용:

1. NSSM 다운로드: https://nssm.cc/download

2. uv 경로 확인:
```powershell
where.exe uv
# 예: C:\Users\YourUser\.cargo\bin\uv.exe
```

3. 서비스 설치:
```powershell
nssm install WindowsExcelService "C:\Users\YourUser\.cargo\bin\uv.exe" "run python windows_excel_service.py"
nssm set WindowsExcelService AppDirectory "C:\path\to\windows-excel-service"
nssm start WindowsExcelService
```

## 서비스 테스트

### 1. 헬스 체크
쳐기
```powershell
curl http://localhost:8889/health
```

예상 응답:
```json
{
  "status": "healthy",
  "excel_available": true,
  "excel_version": "16.0"
}
```

### 2. Excel 생성 테스트

```powershell
curl "http://localhost:8889/excel/odata?odata_url=https://your-odata-server/odata/TableName&sheet_name=Data&query_name=ODataQuery" -o test.xlsx
```

## macOS 서버 연동

### 1. macOS 서버의 `.env` 파일 설정

```bash
# Windows Excel Service URL
WINDOWS_EXCEL_SERVICE_URL=http://your-windows-server-ip:8889
WINDOWS_EXCEL_SERVICE_TIMEOUT=60
```

예시:
```bash
WINDOWS_EXCEL_SERVICE_URL=http://192.168.1.100:8889
WINDOWS_EXCEL_SERVICE_TIMEOUT=60
```

### 2. macOS 서버 의존성 설치 및 재시작

```bash
# httpx 의존성 설치 (이미 pyproject.toml에 포함됨)
uv sync

# 서버 실행
uv run python run_server.py
```

### 3. 연동 테스트

```bash
# macOS 서버를 통해 Excel 생성
curl "http://localhost:8888/odata/musinsa_data/excel-live?sheet_name=Data" -o test.xlsx
```

## API 엔드포인트

### Windows Excel Service (Port 8889)

#### `GET /`
서비스 정보 확인

#### `GET /health`
헬스 체크 - Excel COM 가용성 확인

#### `GET /excel/odata`
Excel 파일 생성

**Parameters:**
- `odata_url` (required): OData 엔드포인트 URL
- `sheet_name` (optional): 시트 이름 (기본: "Data")
- `query_name` (optional): 쿼리 이름 (기본: "ODataQuery")

**Response:**
- Excel 파일 (.xlsx)

### macOS Server (Port 8888)

#### `GET /odata/{table_name}/excel-live`
Windows Excel Service를 통한 Excel 생성 (Proxy)

**Parameters:**
- `$filter` (optional): OData 필터
- `$select` (optional): 필드 선택
- `$orderby` (optional): 정렬
- `sheet_name` (optional): 시트 이름
- `query_name` (optional): 쿼리 이름

**Response:**
- Excel 파일 (.xlsx)
- Power Query 연결이 내장되어 있어 Excel에서 "새로고침" 가능

## 트러블슈팅

### Excel 프로세스가 종료되지 않음

Task Manager에서 수동 종료:
```powershell
Get-Process EXCEL | Stop-Process -Force
```

자동화 스크립트:
```powershell
# cleanup_excel.ps1
$excelProcesses = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excelProcesses) {
    $excelProcesses | Stop-Process -Force
    Write-Host "Excel processes terminated: $($excelProcesses.Count)"
}
```

### COM 초기화 실패

- Excel이 제대로 설치되어 있는지 확인
- Excel을 한 번 실행하여 초기 설정 완료
- Desktop 폴더 권한 확인

### Windows 서버 연결 실패 (macOS 서버에서)

1. 방화벽 확인:
```powershell
# Windows Server에서
New-NetFirewallRule -DisplayName "Excel Service" -Direction Inbound -LocalPort 8889 -Protocol TCP -Action Allow
```

2. 네트워크 연결 테스트:
```bash
# macOS에서
curl http://windows-server-ip:8889/health
```

### 타임아웃 오류

`.env`에서 타임아웃 증가:
```bash
WINDOWS_EXCEL_SERVICE_TIMEOUT=120  # 120초
```

## 보안 고려사항

### 1. 네트워크 격리
Windows Excel Service는 내부 네트워크에서만 접근 가능하도록 설정

### 2. 인증 추가 (선택사항)
API Key 또는 Basic Auth 추가:

```python
from fastapi import Header, HTTPException

async def verify_token(x_api_key: str = Header(...)):
    if x_api_key != "your-secret-key":
        raise HTTPException(status_code=401, detail="Invalid API Key")

@app.get("/excel/odata", dependencies=[Depends(verify_token)])
async def make_odata_excel(...):
    ...
```

### 3. Rate Limiting
동시 요청 수 제한:

```python
from slowapi import Limiter
from slowapi.util import get_remote_address

limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter

@app.get("/excel/odata")
@limiter.limit("10/minute")
async def make_odata_excel(...):
    ...
```

## 모니터링

### 로그 확인

서비스 로그는 stdout으로 출력됩니다.

파일로 저장:
```powershell
python windows_excel_service.py > excel_service.log 2>&1
```

### 프로세스 모니터링

```powershell
# Excel 프로세스 수 확인
(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
```

## 대안 (Windows 서버 없이)

Windows Excel Service를 사용할 수 없는 경우 다음 대안 사용:

### 1. ODC 파일 (추천)
```bash
GET /odata/musinsa_data/connection
```
- Excel에서 더블클릭하여 OData 연결 생성
- 파일 크기: ~1KB
- 크로스 플랫폼

### 2. Excel 템플릿
```bash
GET /odata/musinsa_data/template?sample=10
```
- Power Query M 코드 포함
- 사용 방법 안내 포함
- 샘플 데이터 미리보기

## 참고 자료

- [Microsoft Office Server-Side Automation Considerations](https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2)
- [pywin32 Documentation](https://github.com/mhammond/pywin32)
- [FastAPI Documentation](https://fastapi.tiangolo.com/)
