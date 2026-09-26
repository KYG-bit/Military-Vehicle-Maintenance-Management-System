# ============================================================
# main.py
# 군용 차량 관리 프로그램
#
# 현재 버전의 기능:
#
# 1. Excel 파일 읽기
# 2. Excel 데이터 구조 검사
# 3. 엔진오일 교체 후 주행거리 계산
# 4. 차량 종류별 엔진오일 교체 기준 적용
# 5. 차량별 엔진오일 상태 분류
# 6. 차량 종류별 차량 대수 및 비율 계산
# 7. 계산 결과를 콘솔창에 표 형태로 출력
#
# ※ 현재는 결과를 Excel로 저장하지 않습니다.
# ※ 추후 GUI / Dashboard로 확장할 것을 고려하여
#    각각의 기능을 함수로 분리해 두었습니다.
# ============================================================


import pandas as pd
import tkinter as tk
from tkinter import filedialog

# 차량 종류와 교체 기준 등이 저장된 설정 파일
import config


# ============================================================
# 함수 0. Excel 파일 선택
# ============================================================

def select_excel_file():
    """
    사용자가 분석할 Excel 파일을 직접 선택하도록 하는 함수입니다.

    Returns
    -------
    str
        사용자가 선택한 Excel 파일의 경로
    """

    # tkinter 기본 창 생성
    root = tk.Tk()

    # 빈 tkinter 창은 화면에 표시하지 않음
    root.withdraw()

    # Windows 파일 선택창 열기
    file_path = filedialog.askopenfilename(
        title="차량정보 Excel 파일을 선택하세요",
        filetypes=[
            ("Excel 파일", "*.xlsx"),
            ("Excel 파일", "*.xls"),
            ("모든 파일", "*.*")
        ]
    )

    # tkinter 종료
    root.destroy()

    return file_path

# ============================================================
# 함수 0-1. Excel input
# ============================================================
def get_input_file():
    return "input/차량정보.xlsx"
    """
    빌드할 때 이 부분으로
    root = tk.Tk()
        root.withdraw()

        file_path = filedialog.askopenfilename(
            title="차량정보 Excel 파일을 선택하세요",
            filetypes=[
                ("Excel 파일", "*.xlsx"),
                ("Excel 파일", "*.xls")
            ]
        )

        root.destroy()

        return file_path
 """

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


# ============================================================
# 함수 4. 엔진오일 교체 후 주행거리 계산
# ============================================================

def calculate_oil_mileage(df):
    """
    최근 엔진오일 교체 이후 차량이
    몇 km를 주행했는지 계산합니다.

    계산식:

    현재 총 주행거리
    -
    최근 엔진오일 교체 당시 주행거리
    =
    엔진오일 교체 후 주행거리
    """

    current_mileage = config.COLUMN_CURRENT_MILEAGE
    oil_change_mileage = config.COLUMN_OIL_CHANGE_MILEAGE

    # 엔진오일 교체 이후 주행거리 계산
    df["엔진오일교체후주행거리"] = (
        df[current_mileage]
        - df[oil_change_mileage]
    )

    return df


# ============================================================
# 함수 5. 차량 종류별 엔진오일 교체 주기 가져오기
# ============================================================

def get_oil_change_interval(vehicle_type):
    """
    차량 종류를 입력받아 해당 차량의
    엔진오일 교체 주기를 반환합니다.

    예:

    A → 5000km
    B → 7000km
    C → 10000km

    config.py에 등록되지 않은 차량은 None을 반환합니다.
    """

    return config.OIL_CHANGE_INTERVAL.get(
        vehicle_type,
        None
    )


# ============================================================
# 함수 6. 차량 한 대의 엔진오일 상태 판단
# ============================================================

def classify_oil_status(row):
    """
    차량 1대의 엔진오일 교체 상태를 판단하는 함수
    """

    vehicle_type = row[config.COLUMN_VEHICLE_TYPE]
    current_mileage = row["엔진오일교체후주행거리"]

    # 차량 종류별 교체 기준 확인
    interval = get_oil_change_interval(vehicle_type)

    # 교체 기준이 등록되지 않은 차량
    if interval is None:
        return "기준 없음"

    # 주행거리 데이터가 없는 경우
    if pd.isna(current_mileage):
        return "데이터 확인 필요"

    # 교체 기준 이상이면 무조건 교체 시급
    if current_mileage >= interval:
        return "교체 시급"

    # 교체 기준까지 남은 거리
    remaining_km = interval - current_mileage

    # 기준까지 500km 이하로 남았으면 교체주기 임박
    if remaining_km <= config.REMAINING_WARNING_KM:
        return "교체주기 임박"

    # 그 외에는 정상
    return "정상"


# ============================================================
# 함수 7. 모든 차량의 엔진오일 상태 계산
# ============================================================

def classify_all_vehicles(df):
    """
    Excel에 있는 모든 차량의 엔진오일 상태를
    계산합니다.
    """

    df["엔진오일상태"] = df.apply(
        classify_oil_status,
        axis=1
    )

    return df


# ============================================================
# 함수 8. 차량 종류별 통계 계산
# ============================================================

def calculate_vehicle_type_statistics(df):
    """
    차량 종류별로 엔진오일 상태를 집계합니다.

    예:

    A | 정상           | 30000 | 60.00
    A | 교체주기 임박   | 10000 | 20.00
    A | 교체 필요       |  7000 | 14.00
    A | 교체 시급       |  3000 |  6.00

    B | 정상           | ...
    """

    result = []

    # --------------------------------------------------------
    # 차량 종류별로 데이터를 분리
    # --------------------------------------------------------

    for vehicle_type, group in df.groupby(
        config.COLUMN_VEHICLE_TYPE
    ):

        # 해당 차량 종류의 전체 차량 수
        total = len(group)

        # 상태별 차량 수 계산
        status_counts = (
            group["엔진오일상태"]
            .value_counts()
        )

        # ----------------------------------------------------
        # 미리 정해놓은 순서대로 출력하기 위한 상태 목록
        # ----------------------------------------------------

        status_order = [
            "정상",
            "교체주기 임박",
            "교체 필요",
            "교체 시급",
            "기준 없음",
            "데이터 확인 필요"
        ]

        # ----------------------------------------------------
        # 상태 순서에 따라 결과 생성
        # ----------------------------------------------------

        for status in status_order:

            # 해당 상태의 차량이 없으면
            # 결과에 추가하지 않습니다.
            if status not in status_counts:
                continue

            count = status_counts[status]

            # 해당 차량 종류 내부에서의 비율
            percentage = (
                count / total * 100
            )

            result.append({
                "차량종류": vehicle_type,
                "엔진오일상태": status,
                "차량대수": count,
                "비율(%)": round(
                    percentage,
                    2
                )
            })

    return pd.DataFrame(result)


# ============================================================
# 함수 9. 콘솔에 결과 출력
# ============================================================

def print_statistics(statistics_df):
    """
    차량 종류별 통계 결과를
    콘솔창에 표 형태로 출력합니다.
    """

    print("\n")
    print("=" * 70)
    print("차량 종류별 관리 현황")
    print("=" * 70)

    # pandas DataFrame을 콘솔에 출력
    #
    # index=False
    # → 왼쪽에 0, 1, 2 등의 번호가 나오지 않게 함
    #
    # formatters
    # → 비율을 소수점 둘째 자리까지 표시
    # --------------------------------------------------------

    print(
        statistics_df.to_string(
            index=False,
            formatters={
                "비율(%)": lambda x: f"{x:.2f}"
            }
        )
    )

    print("=" * 70)


# ============================================================
# 함수 10. 프로그램 전체 실행
# ============================================================

def main():

    print("=" * 70)
    print("군용 차량 관리 프로그램")
    print("=" * 70)

    # --------------------------------------------------------
    # 1. 사용자가 Excel 파일 선택
    # --------------------------------------------------------

    input_file = get_input_file()
    
    # --------------------------------------------------------
    # 2. Excel 파일 읽기
    # --------------------------------------------------------

    df = load_excel(input_file)

    # --------------------------------------------------------
    # 3. Excel의 열 구조 확인
    # --------------------------------------------------------

    validate_columns(df)

    # --------------------------------------------------------
    # 4. 숫자 데이터 정리
    # --------------------------------------------------------

    df = clean_numeric_data(df)

    # --------------------------------------------------------
    # 5. 엔진오일 교체 후 주행거리 계산
    # --------------------------------------------------------

    df = calculate_oil_mileage(df)

    # --------------------------------------------------------
    # 6. 모든 차량의 엔진오일 상태 계산
    # --------------------------------------------------------

    df = classify_all_vehicles(df)

    # --------------------------------------------------------
    # 7. 차량 종류별 통계 계산
    # --------------------------------------------------------

    statistics_df = (
        calculate_vehicle_type_statistics(df)
    )

    # --------------------------------------------------------
    # 8. 콘솔창에 결과 출력
    # --------------------------------------------------------

    print_statistics(
        statistics_df
    )

    print("\n프로그램 실행이 완료되었습니다.")


# ============================================================
# 프로그램 시작점
# ============================================================

if __name__ == "__main__":
    main()
