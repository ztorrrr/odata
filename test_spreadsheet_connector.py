"""
Google Spreadsheet Connector 테스트 스크립트
"""
import sys
import json
import requests
from typing import Optional


def test_spreadsheet_connector(
    spreadsheet_id: str,
    sample_size: int = 100,
    token: str = "test_token",
    base_url: str = "http://localhost:8888"
):
    """
    Spreadsheet Connector API 테스트

    Args:
        spreadsheet_id: Google Spreadsheet ID
        sample_size: 샘플 데이터 크기
        token: Bearer 토큰
        base_url: API 베이스 URL
    """

    headers = {"Authorization": f"Bearer {token}"}

    print("\n" + "=" * 60)
    print("Google Spreadsheet Connector Test")
    print("=" * 60)
    print(f"Spreadsheet ID: {spreadsheet_id}")
    print(f"Sample Size: {sample_size}")
    print(f"Base URL: {base_url}")
    print("=" * 60 + "\n")

    # 1. 전체 설정 프로세스 테스트
    print("1. Testing full setup process...")
    print("-" * 40)

    setup_url = f"{base_url}/spreadsheet/setup-spreadsheet/{spreadsheet_id}"
    params = {"sample_size": sample_size}

    try:
        response = requests.post(setup_url, headers=headers, params=params)
        response.raise_for_status()

        result = response.json()
        print(f"[OK] Setup successful!")
        print(f"  - View ID: {result['view_id']}")
        print(f"  - Sample Size: {result['sample_size']}")
        print(f"  - Data Preview: {len(result.get('data_preview', []))} rows")

        view_id = result['view_id']

    except requests.exceptions.RequestException as e:
        print(f"[FAIL] Setup failed: {e}")
        if hasattr(e.response, 'text'):
            print(f"  Error details: {e.response.text}")
        return

    print("\n" + "=" * 60)
    print("Setup Guide for Google Sheets:")
    print("-" * 40)

    if 'setup_guide' in result:
        guide = result['setup_guide']

        print("\nManual Setup Steps:")
        for step in guide.get('manual_setup_steps', []):
            print(f"  {step}")

        print("\nConnection Configuration:")
        config = guide.get('connection_config', {})
        bq_config = config.get('bigquery', {})
        print(f"  - Project ID: {bq_config.get('projectId')}")
        print(f"  - Dataset ID: {bq_config.get('datasetId')}")
        print(f"  - Table ID: {bq_config.get('tableId')}")

        print(f"\nSpreadsheet URL:")
        print(f"  {guide.get('spreadsheet_url')}")

        print("\nSQL Query for Custom Connection:")
        print(f"  {config.get('query')}")

    # 2. 샘플 데이터 미리보기 테스트
    print("\n" + "=" * 60)
    print("2. Testing sample data preview...")
    print("-" * 40)

    preview_url = f"{base_url}/spreadsheet/sample-data"
    params = {"view_id": view_id, "limit": 3}

    try:
        response = requests.get(preview_url, headers=headers, params=params)
        response.raise_for_status()

        preview = response.json()
        print(f"[OK] Preview successful!")
        print(f"  - Retrieved {preview['count']} rows")

        if preview.get('rows'):
            print("\n  Sample Data (first row):")
            first_row = preview['rows'][0]
            for key, value in list(first_row.items())[:5]:  # 처음 5개 컬럼만
                print(f"    - {key}: {value}")
            print(f"    ... and {len(first_row) - 5} more columns")

    except requests.exceptions.RequestException as e:
        print(f"[FAIL] Preview failed: {e}")

    # 3. 연결 설정 정보 테스트
    print("\n" + "=" * 60)
    print("3. Testing connection config...")
    print("-" * 40)

    config_url = f"{base_url}/spreadsheet/connection-config"
    params = {"spreadsheet_id": spreadsheet_id, "view_id": view_id}

    try:
        response = requests.get(config_url, headers=headers, params=params)
        response.raise_for_status()

        config = response.json()
        print(f"[OK] Config retrieved successfully!")
        print(f"  - BigQuery View: {config['bigquery']['viewId']}")
        print(f"  - Connection Type: {config['connection_info']['type']}")

    except requests.exceptions.RequestException as e:
        print(f"[FAIL] Config retrieval failed: {e}")

    print("\n" + "=" * 60)
    print("Test Complete!")
    print("=" * 60)


if __name__ == "__main__":
    # 기본 설정
    DEFAULT_SPREADSHEET_ID = "14v8oM27b8WN5gQFWLf-VvdCDyN2FJY3yqf7J8Zyn_vY"

    # 명령줄 인자 처리
    spreadsheet_id = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SPREADSHEET_ID
    sample_size = int(sys.argv[2]) if len(sys.argv) > 2 else 100

    # 테스트 실행
    test_spreadsheet_connector(
        spreadsheet_id=spreadsheet_id,
        sample_size=sample_size
    )