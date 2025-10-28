"""
BigQuery View 수정 및 Google Sheets 실시간 동기화 테스트
"""
import sys
import time
import json
import requests
from typing import Optional


def test_view_modification(
    token: str = "test_token",
    base_url: str = "http://localhost:8888"
):
    """
    View 수정과 복원 테스트
    """
    headers = {"Authorization": f"Bearer {token}"}

    print("\n" + "=" * 60)
    print("BigQuery View Modification Test")
    print("=" * 60)
    print(f"Base URL: {base_url}")
    print("=" * 60 + "\n")

    # 1. View 수정 - Type 컬럼에 '_테스트' 추가
    print("1. Modifying View - Adding '_테스트' suffix to Type column...")
    print("-" * 40)

    modify_url = f"{base_url}/spreadsheet/modify-view-test"

    try:
        response = requests.post(modify_url, headers=headers)
        response.raise_for_status()

        result = response.json()
        print(f"[OK] View modified successfully!")
        print(f"  - View ID: {result['view_id']}")
        print(f"  - Modified Column: {result['modified_column']}")
        print(f"  - Suffix Added: {result['suffix_added']}")

        # 샘플 데이터 확인
        if result.get('sample_data'):
            sample = result['sample_data'][0] if result['sample_data'] else {}
            type_value = sample.get('Type', 'N/A')
            print(f"  - Sample Type Value: {type_value}")

        print("\n[ACTION REQUIRED] Please check your Google Sheet:")
        print("  1. Open your connected Google Sheet")
        print("  2. Right-click on data area → 'Data connectors' → 'Refresh'")
        print("  3. Check if Type column now shows '_테스트' suffix")
        print("\nPress Enter after checking Sheet...")
        input()

    except requests.exceptions.RequestException as e:
        print(f"[FAIL] Modification failed: {e}")
        return

    # 2. View 복원 - 원본 상태로 되돌리기
    print("\n2. Restoring View to original state...")
    print("-" * 40)

    restore_url = f"{base_url}/spreadsheet/restore-view"

    try:
        response = requests.post(restore_url, headers=headers)
        response.raise_for_status()

        result = response.json()
        print(f"[OK] View restored successfully!")
        print(f"  - View ID: {result['view_id']}")

        # 샘플 데이터 확인
        if result.get('sample_data'):
            sample = result['sample_data'][0] if result['sample_data'] else {}
            type_value = sample.get('Type', 'N/A')
            print(f"  - Sample Type Value: {type_value}")

        print("\n[ACTION REQUIRED] Please check your Google Sheet again:")
        print("  1. Refresh the Sheet data again")
        print("  2. Check if Type column is back to original (no suffix)")

    except requests.exceptions.RequestException as e:
        print(f"[FAIL] Restore failed: {e}")
        return

    print("\n" + "=" * 60)
    print("Test Complete!")
    print("=" * 60)
    print("\n[SUMMARY]")
    print("If the Google Sheet data changed after refresh:")
    print("  [OK] Real-time sync is working!")
    print("If the data didn't change:")
    print("  - Check BigQuery permissions")
    print("  - Try 'Data > Data connectors > Reconnect'")
    print("  - Check if View ID matches in Sheet connection")


def test_custom_modification(
    column_name: str,
    suffix: str,
    token: str = "test_token",
    base_url: str = "http://localhost:8888"
):
    """
    사용자 정의 컬럼과 suffix로 View 수정 테스트
    """
    headers = {"Authorization": f"Bearer {token}"}

    print(f"\nModifying column '{column_name}' with suffix '{suffix}'...")

    modify_url = f"{base_url}/spreadsheet/modify-view-test"
    params = {
        "column_name": column_name,
        "suffix": suffix
    }

    try:
        response = requests.post(modify_url, headers=headers, params=params)
        response.raise_for_status()

        result = response.json()
        print(f"[OK] View modified!")
        print(f"  Modified column: {result['modified_column']}")
        print(f"  Suffix added: {result['suffix_added']}")

    except requests.exceptions.RequestException as e:
        print(f"[FAIL] Modification failed: {e}")
        if hasattr(e.response, 'text'):
            print(f"  Error: {e.response.text}")


if __name__ == "__main__":
    # 기본 테스트 실행
    if len(sys.argv) == 1:
        test_view_modification()
    # 사용자 정의 테스트
    elif len(sys.argv) == 3:
        column = sys.argv[1]
        suffix = sys.argv[2]
        test_custom_modification(column, suffix)
    else:
        print("Usage:")
        print("  python test_view_modification.py")
        print("  python test_view_modification.py <column_name> <suffix>")
        print("\nExamples:")
        print("  python test_view_modification.py")
        print("  python test_view_modification.py Media _modified")
        print("  python test_view_modification.py Device _TEST")