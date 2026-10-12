"""
마케팅 핵심 성과 지표(KPI) 및 채널별/월별/연도별 집계를 담당하는 모듈입니다.
모든 함수는 명확한 타입 힌트와 상세 주석을 포함합니다.
"""

from typing import Dict, List, Optional, Any
import pandas as pd
import numpy as np


def filter_dataframe(df: pd.DataFrame, target_year: Optional[int] = None, target_month: Optional[str] = None) -> pd.DataFrame:
    """
    지정된 연도 및 월 조건으로 DataFrame을 필터링합니다.

    Args:
        df: 원본 데이터프레임
        target_year: 필터링할 연도 (None이면 전체)
        target_month: 필터링할 월 (None이면 전체, 예: '2027-01')

    Returns:
        pd.DataFrame: 필터링된 데이터프레임
    """
    filtered = df.copy()
    if target_year is not None and "년도" in filtered.columns:
        filtered = filtered[filtered["년도"] == target_year]
    if target_month is not None and target_month != "전체" and "월" in filtered.columns:
        filtered = filtered[filtered["월"] == target_month]
    return filtered


def calculate_summary_metrics(data_dict: Dict[str, pd.DataFrame], target_year: Optional[int] = None, target_month: Optional[str] = None) -> Dict[str, float]:
    """
    통합 마케팅 대시보드 상단에 표시되는 6대 핵심 성과 지표(KPI)를 계산합니다.

    - 총 광고비: 네이버 + 쿠팡 + 메타 광고비의 합계
    - 총 매출: DB_매출의 전환매출(판매매출) 합계
    - 통합 ROAS: 총 매출 / 총 광고비 (배수)
    - 통합 CPA: 총 광고비 / 총 전환수 (원)
    - 통합 CTR: 총 클릭수 / 총 노출수 (%)
    - 통합 CVR: 총 전환수 / 총 클릭수 (%)
    - 총 노출수, 총 클릭수, 총 전환수, 총 판매수량, 총 배송건수

    Args:
        data_dict: 로드된 전체 DB 딕셔너리
        target_year: 대상 연도 (None일 시 전체)
        target_month: 대상 월 (None일 시 전체)

    Returns:
        Dict[str, float]: 계산된 성과 지표 딕셔너리
    """
    # 각 DB 필터링
    filtered_naver = filter_dataframe(data_dict.get("DB_네이버", pd.DataFrame()), target_year, target_month)
    filtered_coupang = filter_dataframe(data_dict.get("DB_쿠팡", pd.DataFrame()), target_year, target_month)
    filtered_meta = filter_dataframe(data_dict.get("DB_메타", pd.DataFrame()), target_year, target_month)
    filtered_sales = filter_dataframe(data_dict.get("DB_매출", pd.DataFrame()), target_year, target_month)

    # 마케팅 채널 광고비 합계
    ad_cost_naver: float = float(filtered_naver["광고비"].sum()) if "광고비" in filtered_naver.columns else 0.0
    ad_cost_coupang: float = float(filtered_coupang["광고비"].sum()) if "광고비" in filtered_coupang.columns else 0.0
    ad_cost_meta: float = float(filtered_meta["광고비"].sum()) if "광고비" in filtered_meta.columns else 0.0
    total_ad_cost: float = ad_cost_naver + ad_cost_coupang + ad_cost_meta

    # 마케팅 채널 노출, 클릭, 전환
    total_impressions: float = sum([
        float(filtered_naver["노출수"].sum()) if "노출수" in filtered_naver.columns else 0.0,
        float(filtered_coupang["노출수"].sum()) if "노출수" in filtered_coupang.columns else 0.0,
        float(filtered_meta["노출수"].sum()) if "노출수" in filtered_meta.columns else 0.0
    ])

    total_clicks: float = sum([
        float(filtered_naver["클릭수"].sum()) if "클릭수" in filtered_naver.columns else 0.0,
        float(filtered_coupang["클릭수"].sum()) if "클릭수" in filtered_coupang.columns else 0.0,
        float(filtered_meta["클릭수"].sum()) if "클릭수" in filtered_meta.columns else 0.0
    ])

    total_convs: float = sum([
        float(filtered_naver["전환수"].sum()) if "전환수" in filtered_naver.columns else 0.0,
        float(filtered_coupang["전환수"].sum()) if "전환수" in filtered_coupang.columns else 0.0,
        float(filtered_meta["전환수"].sum()) if "전환수" in filtered_meta.columns else 0.0
    ])

    # 총 매출 (DB_매출 기준, 만약 DB_매출이 비어있다면 광고 채널 전환매출의 합)
    if not filtered_sales.empty and "전환매출" in filtered_sales.columns and filtered_sales["전환매출"].sum() > 0:
        total_revenue: float = float(filtered_sales["전환매출"].sum())
    else:
        total_revenue = sum([
            float(filtered_naver["전환매출"].sum()) if "전환매출" in filtered_naver.columns else 0.0,
            float(filtered_coupang["전환매출"].sum()) if "전환매출" in filtered_coupang.columns else 0.0,
            float(filtered_meta["전환매출"].sum()) if "전환매출" in filtered_meta.columns else 0.0
        ])

    # 비율 지표 산출 (CPA, CPC, 객단가는 원화 정수 단위로 산출)
    roas: float = (total_revenue / total_ad_cost) if total_ad_cost > 0 else 0.0
    cpa: float = float(round(total_ad_cost / total_convs)) if total_convs > 0 else 0.0
    ctr: float = (total_clicks / total_impressions * 100.0) if total_impressions > 0 else 0.0
    cvr: float = (total_convs / total_clicks * 100.0) if total_clicks > 0 else 0.0
    cpc: float = float(round(total_ad_cost / total_clicks)) if total_clicks > 0 else 0.0

    total_quantity: float = float(filtered_sales["판매수량"].sum()) if "판매수량" in filtered_sales.columns else 0.0
    order_aov: float = float(round(total_revenue / total_quantity)) if total_quantity > 0 else 0.0

    return {
        "총 광고비": total_ad_cost,
        "총 매출": total_revenue,
        "통합 ROAS": roas,
        "통합 CPA": cpa,
        "통합 CTR": ctr,
        "통합 CVR": cvr,
        "통합 CPC": cpc,
        "평균 객단가": order_aov,
        "총 노출수": total_impressions,
        "총 클릭수": total_clicks,
        "총 전환수": total_convs,
        "총 판매수량": total_quantity
    }


def get_channel_performance_table(data_dict: Dict[str, pd.DataFrame], target_year: Optional[int] = None, target_month: Optional[str] = None) -> pd.DataFrame:
    """
    구글 스프레드시트의 '채널별 상세 성과 (기준 월 연동)' 표를 생성합니다.
    채널: 네이버, 쿠팡, 메타, 매출

    컬럼: 채널, 광고비, 노출수, 클릭수, 전환수, 전환매출, ROAS, CPA

    Args:
        data_dict: 전체 DB 딕셔너리
        target_year: 대상 연도
        target_month: 대상 월

    Returns:
        pd.DataFrame: 채널별 상세 성과 요약 테이블
    """
    channels = [
        ("네이버", data_dict.get("DB_네이버", pd.DataFrame())),
        ("쿠팡", data_dict.get("DB_쿠팡", pd.DataFrame())),
        ("메타", data_dict.get("DB_메타", pd.DataFrame())),
        ("매출(전체)", data_dict.get("DB_매출", pd.DataFrame()))
    ]

    rows: List[Dict[str, Any]] = []

    for name, df_channel in channels:
        fdf = filter_dataframe(df_channel, target_year, target_month)

        cost = float(fdf["광고비"].sum()) if "광고비" in fdf.columns else 0.0
        imp = float(fdf["노출수"].sum()) if "노출수" in fdf.columns else 0.0
        clicks = float(fdf["클릭수"].sum()) if "클릭수" in fdf.columns else 0.0
        convs = float(fdf["전환수"].sum()) if "전환수" in fdf.columns else 0.0
        rev = float(fdf["전환매출"].sum()) if "전환매출" in fdf.columns else 0.0

        roas = (rev / cost) if cost > 0 else 0.0
        cpa = float(round(cost / convs)) if convs > 0 else 0.0

        rows.append({
            "채널": name,
            "광고비": cost,
            "노출수": int(imp),
            "클릭수": int(clicks),
            "전환수": int(convs),
            "전환매출": rev,
            "ROAS": roas,
            "CPA": cpa
        })

    return pd.DataFrame(rows)


def get_sales_channel_summary_table(
    data_dict: Dict[str, pd.DataFrame],
    target_year: Optional[int] = None,
    target_month: Optional[str] = None
) -> pd.DataFrame:
    """
    '채널별 매출 성과' 표를 생성합니다.
    전체 매출(DB_매출) 데이터를 기준으로 판매 채널(매출분류)별 실적을 요약 집계합니다.
    컬럼: 판매 채널, 매출액, 판매수량, 배송건수, 배송비, 객단가

    Args:
        data_dict: 전체 DB 딕셔너리
        target_year: 대상 연도
        target_month: 대상 월

    Returns:
        pd.DataFrame: 채널별 매출 요약 테이블
    """
    df_sales = data_dict.get("DB_매출", pd.DataFrame())
    if df_sales.empty:
        return pd.DataFrame(columns=["판매 채널", "매출액", "판매수량", "배송건수", "배송비", "객단가"])

    fdf = filter_dataframe(df_sales, target_year, target_month)
    if fdf.empty:
        return pd.DataFrame(columns=["판매 채널", "매출액", "판매수량", "배송건수", "배송비", "객단가"])

    cat_col = "매출분류" if "매출분류" in fdf.columns else "채널"
    if cat_col not in fdf.columns:
        # 분류 컬럼이 없는 경우 단일 행
        rev = float(fdf["전환매출"].sum()) if "전환매출" in fdf.columns else 0.0
        qty = float(fdf["판매수량"].sum()) if "판매수량" in fdf.columns else 0.0
        deliv = float(fdf["배송건수"].sum()) if "배송건수" in fdf.columns else 0.0
        deliv_fee = float(fdf["배송비"].sum()) if "배송비" in fdf.columns else 0.0
        aov = float(round(rev / qty)) if qty > 0 else 0.0
        return pd.DataFrame([{
            "판매 채널": "전체 매출",
            "매출액": rev,
            "판매수량": int(qty),
            "배송건수": int(deliv),
            "배송비": deliv_fee,
            "객단가": aov
        }])

    grp = fdf.groupby(cat_col).agg({
        "전환매출": "sum",
        "판매수량": "sum",
        "배송건수": "sum",
        "배송비": "sum"
    }).reset_index()

    grp["객단가"] = grp.apply(
        lambda r: float(round(r["전환매출"] / r["판매수량"])) if r["판매수량"] > 0 else 0.0,
        axis=1
    )

    grp.rename(columns={
        cat_col: "판매 채널",
        "전환매출": "매출액"
    }, inplace=True)

    # 매출액 기준 내림차순 정렬
    grp.sort_values(by="매출액", ascending=False, inplace=True)

    return grp[["판매 채널", "매출액", "판매수량", "배송건수", "배송비", "객단가"]]


def get_monthly_trend_table(data_dict: Dict[str, pd.DataFrame], target_year: Optional[int] = None) -> pd.DataFrame:
    """
    구글 스프레드시트 하단의 '월별 추이 테이블'을 계산합니다.
    컬럼: 월, 총 광고비, 총 매출, ROAS, CPA, 객단가, CPC, CTR, CVR

    Args:
        data_dict: 전체 DB 딕셔너리
        target_year: 대상 연도

    Returns:
        pd.DataFrame: 월별 성과 요약 테이블
    """
    # 모든 DB에서 등장하는 월 목록 취합
    all_months = set()
    for df in data_dict.values():
        if "월" in df.columns:
            m_series = df["월"].dropna().astype(str).str.strip()
            if target_year is not None and "년도" in df.columns:
                m_series = df[df["년도"] == target_year]["월"].dropna().astype(str).str.strip()
            all_months.update(m_series.tolist())

    sorted_months = sorted([m for m in all_months if m and m != "nan"])
    if not sorted_months:
        sorted_months = [f"{target_year or 2027}-{i:02d}" for i in range(1, 13)]

    rows: List[Dict[str, Any]] = []

    for m in sorted_months:
        metrics = calculate_summary_metrics(data_dict, target_year, m)
        rows.append({
            "월": m,
            "총 광고비": metrics["총 광고비"],
            "총 매출": metrics["총 매출"],
            "ROAS": metrics["통합 ROAS"],
            "CPA": metrics["통합 CPA"],
            "객단가": metrics["평균 객단가"],
            "CPC": metrics["통합 CPC"],
            "CTR": metrics["통합 CTR"],
            "CVR": metrics["통합 CVR"],
            "노출수": metrics["총 노출수"],
            "클릭수": metrics["총 클릭수"],
            "전환수": metrics["총 전환수"]
        })

    return pd.DataFrame(rows)


def get_yearly_trend_table(data_dict: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    연도별 성과 요약 테이블을 계산합니다 (월별 추이 성과표와 동일한 9대 핵심 지표 적용).
    컬럼: 연도, 총 광고비, 총 매출, ROAS, CPA, 객단가, CPC, CTR, CVR

    Args:
        data_dict: 전체 DB 딕셔너리

    Returns:
        pd.DataFrame: 연도별 요약 테이블
    """
    all_years = set()
    for df in data_dict.values():
        if "년도" in df.columns:
            all_years.update(df["년도"].dropna().astype(int).tolist())

    sorted_years = sorted(list(all_years)) if all_years else [2027, 2028, 2029]
    rows: List[Dict[str, Any]] = []

    for y in sorted_years:
        metrics = calculate_summary_metrics(data_dict, target_year=y, target_month=None)
        rows.append({
            "연도": f"{y}년",
            "총 광고비": metrics["총 광고비"],
            "총 매출": metrics["총 매출"],
            "ROAS": metrics["통합 ROAS"],
            "CPA": metrics["통합 CPA"],
            "객단가": metrics["평균 객단가"],
            "CPC": metrics["통합 CPC"],
            "CTR": metrics["통합 CTR"],
            "CVR": metrics["통합 CVR"],
            "노출수": metrics["총 노출수"],
            "클릭수": metrics["총 클릭수"],
            "전환수": metrics["총 전환수"]
        })

    return pd.DataFrame(rows)
