import pandas as pd
import config

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

