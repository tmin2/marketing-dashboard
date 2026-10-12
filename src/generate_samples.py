"""
초기 마케팅 샘플 데이터 및 템플릿 생성 모듈입니다.
구글 스프레드시트의 2027년 구조를 기반으로 실감나는 샘플 데이터를 생성합니다.
모든 함수는 명확한 타입 힌트와 상세 주석을 포함합니다.
"""

from typing import Dict, List
import os
import pandas as pd
import numpy as np

PROJECT_ROOT: str = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE_DIR: str = os.path.join(PROJECT_ROOT, "data")


def generate_all_samples() -> None:
    """
    네이버, 쿠팡, 메타, 매출 4대 데이터베이스의 초기 2027년 샘플 데이터와
    엑셀 업로드 템플릿 파일들을 생성하여 data 디렉토리에 저장합니다.
    """
    os.makedirs(BASE_DIR, exist_ok=True)
    templates_dir = os.path.join(BASE_DIR, "templates")
    os.makedirs(templates_dir, exist_ok=True)

    # 2027년 1월 ~ 12월 리스트 생성
    months: List[str] = [f"2027-{i:02d}" for i in range(1, 13)]
    np.random.seed(42)

    # 1. DB_네이버 샘플 데이터 생성
    naver_categories: List[str] = ["쇼핑검색", "플러스", "파워링크", "애드부스트", "기타"]
    naver_rows: List[Dict] = []
    for m in months:
        for cat in naver_categories:
            cost = int(np.random.uniform(800_000, 3_500_000))
            imp = int(cost * np.random.uniform(8, 25))
            clicks = int(imp * np.random.uniform(0.015, 0.045))
            convs = int(clicks * np.random.uniform(0.02, 0.08))
            rev = int(cost * np.random.uniform(2.5, 6.0))
            naver_rows.append({
                "년도": 2027,
                "월": m,
                "광고비": cost,
                "노출수": imp,
                "클릭수": clicks,
                "전환수": convs,
                "전환매출": rev,
                "마케팅 분류": cat
            })
    df_naver = pd.DataFrame(naver_rows)
    df_naver.to_csv(os.path.join(BASE_DIR, "DB_네이버.csv"), index=False, encoding="utf-8-sig")

    # 2. DB_쿠팡 샘플 데이터 생성
    coupang_rows: List[Dict] = []
    for m in months:
        cost = int(np.random.uniform(4_000_000, 9_000_000))
        imp = int(cost * np.random.uniform(10, 30))
        clicks = int(imp * np.random.uniform(0.02, 0.05))
        convs = int(clicks * np.random.uniform(0.04, 0.10))
        rev = int(cost * np.random.uniform(3.0, 7.5))
        coupang_rows.append({
            "년도": 2027,
            "월": m,
            "광고비": cost,
            "노출수": imp,
            "클릭수": clicks,
            "전환수": convs,
            "전환매출": rev
        })
    df_coupang = pd.DataFrame(coupang_rows)
    df_coupang.to_csv(os.path.join(BASE_DIR, "DB_쿠팡.csv"), index=False, encoding="utf-8-sig")

    # 3. DB_메타 샘플 데이터 생성
    meta_rows: List[Dict] = []
    for m in months:
        cost = int(np.random.uniform(2_500_000, 6_000_000))
        imp = int(cost * np.random.uniform(15, 40))
        clicks = int(imp * np.random.uniform(0.01, 0.035))
        convs = int(clicks * np.random.uniform(0.015, 0.05))
        rev = int(cost * np.random.uniform(2.0, 4.8))
        meta_rows.append({
            "년도": 2027,
            "월": m,
            "광고비": cost,
            "노출수": imp,
            "클릭수": clicks,
            "전환수": convs,
            "전환매출": rev
        })
    df_meta = pd.DataFrame(meta_rows)
    df_meta.to_csv(os.path.join(BASE_DIR, "DB_메타.csv"), index=False, encoding="utf-8-sig")

    # 4. DB_매출 샘플 데이터 생성
    sales_categories: List[str] = ["네이버", "쿠팡", "자사몰", "SK베네피아", "GS샵", "전화주문", "기타"]
    sales_rows: List[Dict] = []
    for m in months:
        for cat in sales_categories:
            cost = int(np.random.uniform(0, 500_000)) if cat in ["네이버", "쿠팡", "자사몰"] else 0
            imp = int(cost * np.random.uniform(5, 15)) if cost > 0 else 0
            clicks = int(imp * 0.03) if imp > 0 else 0
            qty = int(np.random.uniform(100, 1500))
            deliv_count = int(qty * np.random.uniform(0.8, 0.95))
            deliv_fee = deliv_count * 3000
            conv_sales = int(qty * np.random.uniform(25_000, 45_000))
            convs = int(deliv_count)

            sales_rows.append({
                "년도": 2027,
                "월": m,
                "광고비": cost,
                "노출수": imp,
                "클릭수": clicks,
                "전환수": convs,
                "전환매출": conv_sales,
                "판매수량": qty,
                "배송건수": deliv_count,
                "배송비": deliv_fee,
                "매출분류": cat
            })
    df_sales = pd.DataFrame(sales_rows)
    df_sales.to_csv(os.path.join(BASE_DIR, "DB_매출.csv"), index=False, encoding="utf-8-sig")

    # 5. 빈 엑셀 템플릿(양식 파일) 생성
    template_path = os.path.join(templates_dir, "마케팅_통합데이터_업로드_양식.xlsx")
    with pd.ExcelWriter(template_path, engine="openpyxl") as writer:
        df_naver.head(2).to_excel(writer, sheet_name="DB_네이버", index=False)
        df_coupang.head(2).to_excel(writer, sheet_name="DB_쿠팡", index=False)
        df_meta.head(2).to_excel(writer, sheet_name="DB_메타", index=False)
        df_sales.head(2).to_excel(writer, sheet_name="DB_매출", index=False)

    # 개별 템플릿 생성
    df_naver.head(0).to_excel(os.path.join(templates_dir, "양식_DB_네이버.xlsx"), index=False)
    df_coupang.head(0).to_excel(os.path.join(templates_dir, "양식_DB_쿠팡.xlsx"), index=False)
    df_meta.head(0).to_excel(os.path.join(templates_dir, "양식_DB_메타.xlsx"), index=False)
    df_sales.head(0).to_excel(os.path.join(templates_dir, "양식_DB_매출.xlsx"), index=False)


if __name__ == "__main__":
    generate_all_samples()
    print("Sample datasets and templates generated successfully.")
