import pandas as pd
import config

# ============================================================
# 함수 0-1. Excel input
# ============================================================

def get_input_file():
    """
    현재 개발 환경에서는
    input 폴더의 차량정보.xlsx를 사용합니다.

    추후 Windows GUI 버전에서는
    Excel 파일 선택창을 사용하도록 변경합니다.
    """

    return "input/차량정보.xlsx"
# ============================================================
# 함수 1. Excel 파일 읽기
# ============================================================

def load_excel(file_path):
    """
    Excel 파일을 읽어서 pandas DataFrame으로 반환합니다.

    Parameters
    ----------
    file_path : str
        읽어올 Excel 파일의 경로

    Returns
    -------
    pandas.DataFrame
        Excel에서 읽어온 차량 데이터
    """

    # Excel 파일 읽기
    df = pd.read_excel(
        file_path,
        sheet_name=config.SHEET_NAME
    )

    print(f"총 {len(df):,}개의 차량 데이터를 읽었습니다.")

    return df


# ============================================================
# 함수 2. 필수 열 존재 여부 확인
# ============================================================

def validate_columns(df):
    """
    Excel 파일에 프로그램 실행에 필요한 열이
    모두 존재하는지 검사합니다.

    필요한 열이 없으면 오류를 발생시킵니다.
    """

    # 프로그램에서 반드시 필요한 Excel 열
    required_columns = [
        config.COLUMN_VEHICLE_ID,
        config.COLUMN_VEHICLE_TYPE,
        config.COLUMN_OIL_CHANGE_DATE,
        config.COLUMN_OIL_CHANGE_MILEAGE,
        config.COLUMN_CURRENT_MILEAGE
    ]

    # Excel에 존재하지 않는 열 찾기
    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    # 누락된 열이 있는 경우
    if missing_columns:

        print("\n[오류] 다음 열이 Excel 파일에 없습니다.")

        for column in missing_columns:
            print(f" - {column}")

        raise ValueError(
            "Excel 파일의 열 이름을 확인해주세요."
        )



# ============================================================
# 함수 3. 숫자 데이터 정리
# ============================================================

def clean_numeric_data(df):
    """
    주행거리 데이터를 숫자 형태로 변환합니다.

    Excel에서 숫자가 문자열로 저장되어 있거나
    잘못된 값이 들어있는 경우를 처리합니다.
    """

    oil_mileage = config.COLUMN_OIL_CHANGE_MILEAGE
    current_mileage = config.COLUMN_CURRENT_MILEAGE

    # 최근 엔진오일 교체 당시 주행거리
    df[oil_mileage] = pd.to_numeric(
        df[oil_mileage],
        errors="coerce"
    )

    # 현재 총 주행거리
    df[current_mileage] = pd.to_numeric(
        df[current_mileage],
        errors="coerce"
    )

    return df
