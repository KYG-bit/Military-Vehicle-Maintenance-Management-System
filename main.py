# ============================================================
# main.py
# 차량 관리 프로그램
#
# 기능:
# 1. Excel 파일 읽기
# 2. 데이터 형식 확인
# 3. 차량별 엔진오일 사용 거리 계산
# 4. 차량 종류별 엔진오일 교체 기준 적용
# 5. 차량을 상태별로 분류
# 6. 전체 차량의 상태별 비율 계산
# 7. 결과를 Excel 파일로 저장
# ============================================================


import os
import pandas as pd

# 우리가 만든 설정 파일을 불러옵니다.
import config


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

    print("Excel 파일을 읽는 중입니다...")

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

    필요한 열이 없다면 오류를 발생시킵니다.
    """

    required_columns = [
        config.COLUMN_VEHICLE_ID,
        config.COLUMN_VEHICLE_TYPE,
        config.COLUMN_OIL_CHANGE_DATE,
        config.COLUMN_OIL_CHANGE_MILEAGE,
        config.COLUMN_CURRENT_MILEAGE
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:

        print("\n[오류] 다음 열이 Excel 파일에 없습니다.")

        for column in missing_columns:
            print(f" - {column}")

        raise ValueError(
            "Excel 파일의 열 이름을 확인해주세요."
        )

    print("Excel 열 구조 확인 완료")


# ============================================================
# 함수 3. 숫자 데이터 정리
# ============================================================

def clean_numeric_data(df):
    """
    주행거리 데이터를 숫자 형태로 변환합니다.

    Excel에서 숫자가 문자열로 저장되어 있거나
    빈 값이 존재하는 경우를 처리합니다.
    """

    oil_mileage = config.COLUMN_OIL_CHANGE_MILEAGE
    current_mileage = config.COLUMN_CURRENT_MILEAGE

    # 숫자로 변환할 수 없는 값은 NaN으로 처리
    df[oil_mileage] = pd.to_numeric(
        df[oil_mileage],
        errors="coerce"
    )

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

    df["엔진오일교체후주행거리"] = (
        df[current_mileage]
        - df[oil_change_mileage]
    )

    return df


# ============================================================
# 함수 5. 차량 종류별 엔진오일 교체 기준 가져오기
# ============================================================

def get_oil_change_interval(vehicle_type):
    """
    차량 종류를 입력받아 해당 차량의
    엔진오일 교체 주기를 반환합니다.

    예:

    A → 5000
    B → 7000
    C → 10000
    """

    return config.OIL_CHANGE_INTERVAL.get(
        vehicle_type,
        None
    )


# ============================================================
# 함수 6. 차량별 엔진오일 상태 분류
# ============================================================

def classify_oil_status(row):
    """
    차량 한 대의 엔진오일 상태를 판단합니다.

    상태는 다음 세 가지로 분류합니다.

    1. 교체주기 임박
    2. 교체 필요
    3. 교체 시급

    추가적으로 데이터를 계산할 수 없는 경우
    "데이터 확인 필요"로 분류합니다.
    """

    vehicle_type = row[config.COLUMN_VEHICLE_TYPE]

    current_mileage = row["엔진오일교체후주행거리"]

    # 차량 종류에 해당하는 교체 주기 가져오기
    interval = get_oil_change_interval(vehicle_type)

    # 차량 종류에 대한 기준이 없는 경우
    if interval is None:
        return "기준 없음"

    # 주행거리 데이터가 없는 경우
    if pd.isna(current_mileage):
        return "데이터 확인 필요"

    # --------------------------------------------------------
    # 교체주기까지 남은 거리
    # --------------------------------------------------------

    remaining_km = interval - current_mileage

    # --------------------------------------------------------
    # 1단계
    # 아직 교체주기에 도달하지 않았고
    # 일정 거리 이내로 가까워진 경우
    # --------------------------------------------------------

    if remaining_km > 0:

        if remaining_km <= config.REMAINING_WARNING_KM:
            return "교체주기 임박"

        else:
            return "정상"

    # --------------------------------------------------------
    # 2단계
    # 교체주기에 도달했거나 초과한 경우
    # --------------------------------------------------------

    overdue_km = abs(remaining_km)

    # 교체주기를 초과했지만
    # 일정 수준 이내인 경우
    if overdue_km < config.OVERDUE_URGENT_KM:
        return "교체 필요"

    # 교체주기를 크게 초과한 경우
    else:
        return "교체 시급"


# ============================================================
# 함수 7. 전체 차량에 엔진오일 상태 적용
# ============================================================

def classify_all_vehicles(df):
    """
    모든 차량에 대해 엔진오일 상태를 계산합니다.
    """

    print("엔진오일 상태를 분류하는 중입니다...")

    df["엔진오일상태"] = df.apply(
        classify_oil_status,
        axis=1
    )

    return df


# ============================================================
# 함수 8. 상태별 통계 계산
# ============================================================

def calculate_statistics(df):
    """
    전체 차량 중 각 상태에 해당하는 차량의
    대수와 비율을 계산합니다.
    """

    total_count = len(df)

    # 상태별 차량 수 계산
    status_counts = (
        df["엔진오일상태"]
        .value_counts()
    )

    # 결과를 저장할 리스트
    statistics = []

    for status, count in status_counts.items():

        percentage = (
            count / total_count * 100
        )

        statistics.append({
            "엔진오일상태": status,
            "차량대수": count,
            "비율(%)": round(percentage, 2)
        })

    # DataFrame으로 변환
    statistics_df = pd.DataFrame(
        statistics
    )

    # 전체 차량 수 추가
    total_row = pd.DataFrame([{
        "엔진오일상태": "전체",
        "차량대수": total_count,
        "비율(%)": 100.0
    }])

    statistics_df = pd.concat(
        [statistics_df, total_row],
        ignore_index=True
    )

    return statistics_df


# ============================================================
# 함수 9. 차량 종류별 통계 계산
# ============================================================

def calculate_vehicle_type_statistics(df):
    """
    차량 종류별로 엔진오일 상태를 계산합니다.

    예:

    A 차량
    - 정상 50%
    - 교체주기 임박 20%
    - 교체 필요 20%
    - 교체 시급 10%

    B 차량
    ...
    """

    result = []

    for vehicle_type, group in df.groupby(
        config.COLUMN_VEHICLE_TYPE
    ):

        total = len(group)

        status_counts = (
            group["엔진오일상태"]
            .value_counts()
        )

        for status, count in status_counts.items():

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
# 함수 10. 결과 Excel 저장
# ============================================================

def save_result(
    df,
    statistics_df,
    vehicle_type_statistics_df,
    output_path
):
    """
    계산 결과를 하나의 Excel 파일에
    여러 Sheet로 저장합니다.

    Sheet 구성:

    1. 차량별결과
    2. 전체통계
    3. 차량종류별통계
    """

    print("결과 Excel 파일을 생성하는 중입니다...")

    # 출력 폴더가 없다면 생성
    output_folder = os.path.dirname(output_path)

    if output_folder:
        os.makedirs(
            output_folder,
            exist_ok=True
        )

    # ExcelWriter를 사용하여
    # 여러 개의 Sheet를 하나의 Excel 파일에 저장
    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # Sheet 1
        # 차량별 상세 결과
        # ----------------------------------------------------

        df.to_excel(
            writer,
            sheet_name="차량별결과",
            index=False
        )

        # ----------------------------------------------------
        # Sheet 2
        # 전체 통계
        # ----------------------------------------------------

        statistics_df.to_excel(
            writer,
            sheet_name="전체통계",
            index=False
        )

        # ----------------------------------------------------
        # Sheet 3
        # 차량 종류별 통계
        # ----------------------------------------------------

        vehicle_type_statistics_df.to_excel(
            writer,
            sheet_name="차량종류별통계",
            index=False
        )

    print("\n결과 파일 생성 완료!")
    print(f"저장 위치: {output_path}")


# ============================================================
# 함수 11. 프로그램 실행 함수
# ============================================================

def main():

    print("=" * 60)
    print("군용 차량 관리 프로그램")
    print("=" * 60)

    # --------------------------------------------------------
    # 1. 입력 Excel 파일 경로
    # --------------------------------------------------------
    #
    # 처음에는 input 폴더에
    # 차량정보.xlsx
    # 파일을 넣어두고 사용합니다.
    #
    # 나중에 GUI 파일 선택 기능을 추가할 수도 있습니다.
    # --------------------------------------------------------

    input_file = "input/차량정보.xlsx"

    output_file = (
        "output/"
        + config.OUTPUT_FILE
    )

    # --------------------------------------------------------
    # 2. Excel 읽기
    # --------------------------------------------------------

    df = load_excel(input_file)

    # --------------------------------------------------------
    # 3. Excel 구조 검사
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
    # 6. 엔진오일 상태 분류
    # --------------------------------------------------------

    df = classify_all_vehicles(df)

    # --------------------------------------------------------
    # 7. 전체 통계
    # --------------------------------------------------------

    statistics_df = calculate_statistics(df)

    # --------------------------------------------------------
    # 8. 차량 종류별 통계
    # --------------------------------------------------------

    vehicle_type_statistics_df = (
        calculate_vehicle_type_statistics(df)
    )

    # --------------------------------------------------------
    # 9. 결과 Excel 저장
    # --------------------------------------------------------

    save_result(
        df,
        statistics_df,
        vehicle_type_statistics_df,
        output_file
    )

    print("\n프로그램 실행이 완료되었습니다.")


# ============================================================
# 프로그램 시작점
# ============================================================

if __name__ == "__main__":
    main()
