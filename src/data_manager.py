"""
데이터 로드, 검증, 저장, 정제 및 네이버/메타/쿠팡 광고 원본 보고서 연동 모듈입니다.
네이버, 메타, 쿠팡 광고센터에서 내려받은 원본 엑셀/CSV 보고서를 자동 인식 및 변환합니다.
모든 함수는 명확한 타입 힌트와 상세 주석을 포함합니다.
"""

from typing import Dict, List, Optional, Tuple, Any
import os
import io
import re
import pandas as pd
import numpy as np

# 프로젝트 루트 및 데이터 디렉토리 동적 경로 설정 (클라우드/Linux/Windows 호환)
PROJECT_ROOT: str = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR: str = os.path.join(PROJECT_ROOT, "data")
TEMPLATES_DIR: str = os.path.join(DATA_DIR, "templates")

# 관리 대상 데이터베이스 키 정의
DB_KEYS: List[str] = ["DB_네이버", "DB_쿠팡", "DB_메타", "DB_매출"]

# 각 DB별 기본 컬럼 정의 (캠페인 이름 등 세부정보도 함께 보관 가능하도록 확장)
DB_COLUMNS: Dict[str, List[str]] = {
    "DB_네이버": ["년도", "월", "광고비", "노출수", "클릭수", "전환수", "전환매출", "마케팅 분류", "캠페인 이름"],
    "DB_쿠팡": ["년도", "월", "광고비", "노출수", "클릭수", "전환수", "전환매출", "캠페인 이름"],
    "DB_메타": ["년도", "월", "광고비", "노출수", "클릭수", "전환수", "전환매출", "캠페인 이름"],
    "DB_매출": ["년도", "월", "광고비", "노출수", "클릭수", "전환수", "전환매출", "판매수량", "배송건수", "배송비", "매출분류"]
}

# 숫자형 변환 대상 컬럼 목록
NUMERIC_COLUMNS: List[str] = [
    "광고비", "노출수", "클릭수", "전환수", "전환매출", "판매수량", "배송건수", "배송비"
]


def get_db_file_path(db_name: str) -> str:
    """DB 이름에 해당하는 CSV 파일의 절대 경로를 반환합니다."""
    return os.path.join(DATA_DIR, f"{db_name}.csv")


def load_db(db_name: str) -> pd.DataFrame:
    """
    지정된 데이터베이스의 CSV 파일을 로드하여 정제된 DataFrame으로 반환합니다.
    파일이 없을 경우 빈 DataFrame을 반환합니다.
    """
    file_path: str = get_db_file_path(db_name)
    expected_cols: List[str] = DB_COLUMNS.get(db_name, [])

    if not os.path.exists(file_path):
        return pd.DataFrame(columns=expected_cols)

    try:
        df: pd.DataFrame = pd.read_csv(file_path, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(file_path, encoding="cp949")

    # 누락된 필수 컬럼 추가
    for col in expected_cols:
        if col not in df.columns:
            df[col] = 0 if col in NUMERIC_COLUMNS else ""

    # 숫자형 컬럼 결측치 치환 및 타입 변환
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("₩", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 년도 및 월 정리
    if "년도" in df.columns:
        df["년도"] = pd.to_numeric(df["년도"], errors="coerce").fillna(2026).astype(int)
    if "월" in df.columns:
        df["월"] = df["월"].astype(str).str.strip()

    return df


def load_all_dbs() -> Dict[str, pd.DataFrame]:
    """모든 데이터베이스를 딕셔너리 형태로 로드합니다."""
    return {db_key: load_db(db_key) for db_key in DB_KEYS}


def save_db(db_name: str, df: pd.DataFrame) -> None:
    """DataFrame을 해당 DB의 CSV 파일로 저장합니다."""
    os.makedirs(DATA_DIR, exist_ok=True)
    file_path: str = get_db_file_path(db_name)
    df.to_csv(file_path, index=False, encoding="utf-8-sig")


def clean_numeric_val(val: Any) -> float:
    """문자열이나 숫자형 데이터를 안전한 float으로 변환합니다."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).replace(",", "").replace("₩", "").replace("%", "").strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0


# ----------------------------------------------------------------------
# 네이버 / 메타 / 쿠팡 원본 광고 보고서 특화 파서
# ----------------------------------------------------------------------

def parse_naver_ad_report(df_raw: pd.DataFrame, default_year: int, default_month: str) -> pd.DataFrame:
    """
    네이버 검색광고/성과형디스플레이 보고서 원본을 정제하여 DB_네이버 포맷으로 변환합니다.

    원본 헤더 예시:
    ON/OFF, 광고 구분, 상태, 보조 상태, 캠페인 ID, 캠페인 이름, 캠페인 분류, 노출수, 클릭수, 클릭률(%),
    평균 CPC, 평균 CPM, 총비용, 총 전환수, 총 전환매출액, 총 전환율(%), 총 전환당비용, 구매완료 광고수익률(%)

    - 마지막 행의 요약 행(예: '캠페인 10개 결과', '전체')은 자동 제외
    """
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # 합계/요약 행 제거 (캠페인 이름 또는 캠페인 ID에 결과/합계/전체가 들어가거나 ON/OFF가 비어있는 경우)
    if "ON/OFF" in df.columns:
        df = df[df["ON/OFF"].notna() & (df["ON/OFF"].astype(str).str.strip() != "")]
    if "캠페인 이름" in df.columns:
        df = df[~df["캠페인 이름"].astype(str).str.contains("결과|합계|전체", regex=True, na=False)]

    records: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        # 총비용 매핑
        cost = clean_numeric_val(row.get("총비용", row.get("광고비", 0)))
        imp = clean_numeric_val(row.get("노출수", row.get("노출", 0)))
        clicks = clean_numeric_val(row.get("클릭수", row.get("클릭", 0)))
        convs = clean_numeric_val(row.get("총 전환수", row.get("전환수", 0)))
        rev = clean_numeric_val(row.get("총 전환매출액", row.get("전환매출", row.get("전환매출액", 0))))

        cat = str(row.get("캠페인 분류", row.get("마케팅 분류", "네이버광고"))).strip()
        camp_name = str(row.get("캠페인 이름", "")).strip()

        records.append({
            "년도": default_year,
            "월": default_month,
            "광고비": cost,
            "노출수": int(imp),
            "클릭수": int(clicks),
            "전환수": int(convs),
            "전환매출": rev,
            "마케팅 분류": cat if cat and cat != "nan" else "기타",
            "캠페인 이름": camp_name
        })

    return pd.DataFrame(records)


def parse_meta_ad_report(df_raw: pd.DataFrame, default_year: int, default_month: str) -> pd.DataFrame:
    """
    메타(페이스북/인스타그램) 광고 보고서 원본을 정제하여 DB_메타 포맷으로 변환합니다.

    원본 헤더 예시:
    보고 시작, 보고 종료, 캠페인 이름, 캠페인 게재, 지출 금액 (KRW), 노출, 조회, 구매 전환값, 구매 등
    """
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    records: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        camp_name = str(row.get("캠페인 이름", "")).strip()
        if not camp_name or camp_name == "nan" or "결과" in camp_name or "합계" in camp_name:
            continue

        # 날짜 파싱 시도 (보고 시작: '2026-10-11')
        row_year = default_year
        row_month = default_month
        start_date = str(row.get("보고 시작", "")).strip()
        if re.match(r"^\d{4}-\d{2}", start_date):
            parts = start_date.split("-")
            row_year = int(parts[0])
            row_month = f"{parts[0]}-{parts[1]}"

        cost = clean_numeric_val(row.get("지출 금액 (KRW)", row.get("지출 금액", row.get("광고비", 0))))
        imp = clean_numeric_val(row.get("노출", row.get("노출수", 0)))

        # 클릭수 (조회, 링크 클릭, 클릭수 중 매칭)
        clicks = clean_numeric_val(row.get("조회", row.get("링크 클릭", row.get("클릭수", row.get("클릭", 0)))))

        # 전환수: '구매', '결과', '총 전환수', '전환수' 순으로 탐색
        convs = clean_numeric_val(row.get("구매", 0))
        if convs == 0:
            convs = clean_numeric_val(row.get("결과", row.get("전환수", row.get("총 전환수", 0))))

        # 전환매출: '구매 전환값', '웹사이트 구매 전환값', 'Meta 구매 전환값', '전환매출' 순으로 탐색
        rev = clean_numeric_val(row.get("구매 전환값", 0))
        if rev == 0:
            rev = clean_numeric_val(row.get("웹사이트 구매 전환값", row.get("Meta 구매 전환값", row.get("전환매출", 0))))

        # 전환매출이 0인데 지출금액과 ROAS가 있는 경우 역산 (광고비 * ROAS)
        if rev == 0 and cost > 0:
            roas_cand = clean_numeric_val(row.get("구매 ROAS(광고 지출 대비 수익률)", row.get("결과 ROAS", 0)))
            if roas_cand > 0:
                rev = round(cost * roas_cand, 0)

        records.append({
            "년도": row_year,
            "월": row_month,
            "광고비": cost,
            "노출수": int(imp),
            "클릭수": int(clicks),
            "전환수": int(convs),
            "전환매출": rev,
            "캠페인 이름": camp_name
        })

    return pd.DataFrame(records)


def parse_coupang_ad_report(df_raw: pd.DataFrame, default_year: int, default_month: str) -> pd.DataFrame:
    """
    쿠팡 광고 관리시스템 보고서 원본을 정제하여 DB_쿠팡 포맷으로 변환합니다.

    원본 헤더 지원:
    광고일자(날짜), 캠페인명, 캠페인종류, 광고비(집행금액), 노출수, 클릭수, 전환수(구매건수), 전환매출액(총 전환매출액)
    """
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    records: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        camp_name = str(row.get("캠페인명", row.get("캠페인 이름", row.get("캠페인", "")))).strip()
        if "합계" in camp_name or "총계" in camp_name or "소계" in camp_name:
            continue

        # 날짜 컬럼 확인
        row_year = default_year
        row_month = default_month
        date_val = str(row.get("광고일자", row.get("날짜", row.get("일자", "")))).strip()
        if re.match(r"^\d{4}-\d{2}", date_val):
            parts = date_val.split("-")
            row_year = int(parts[0])
            row_month = f"{parts[0]}-{parts[1]}"
        elif re.match(r"^\d{4}\.\d{2}", date_val):
            parts = date_val.split(".")
            row_year = int(parts[0])
            row_month = f"{parts[0]}-{parts[1]}"

        # 광고비 매핑
        cost = 0.0
        for col_cand in ["광고비", "집행금액", "광고비(원)", "총비용", "지출금액"]:
            if col_cand in row:
                cost = clean_numeric_val(row[col_cand])
                break

        imp = clean_numeric_val(row.get("노출수", row.get("노출", 0)))
        clicks = clean_numeric_val(row.get("클릭수", row.get("클릭", 0)))

        # 전환수 매핑
        convs = 0.0
        for col_cand in ["전환수", "구매건수", "전환수(구매건수)", "총 전환수", "직접전환수"]:
            if col_cand in row:
                convs = clean_numeric_val(row[col_cand])
                break

        # 전환매출 매핑
        rev = 0.0
        for col_cand in ["전환매출", "전환매출액", "총 전환매출액", "전환매출액(원)", "직접전환매출"]:
            if col_cand in row:
                rev = clean_numeric_val(row[col_cand])
                break

        records.append({
            "년도": row_year,
            "월": row_month,
            "광고비": cost,
            "노출수": int(imp),
            "클릭수": int(clicks),
            "전환수": int(convs),
            "전환매출": rev,
            "캠페인 이름": camp_name
        })

    return pd.DataFrame(records)


def detect_file_format(df: pd.DataFrame, filename: str = "") -> str:
    """
    DataFrame의 헤더 컬럼 및 파일명을 분석하여 보고서 유형을 자동 판별합니다.
    (한글/영문 파일명 및 컬럼 시그니처를 모두 지원)

    Returns:
        str: 'DB_네이버', 'DB_메타', 'DB_쿠팡', 'DB_매출' 또는 'STANDARD'
    """
    fn_lower = filename.lower()
    cols = [str(c).strip().lower() for c in df.columns]

    # 1. 파일명 기반 판별 (영문/한글 무관)
    if any(k in fn_lower for k in ["naver", "네이버", "gfa"]):
        return "DB_네이버"
    if any(k in fn_lower for k in ["coupang", "쿠팡"]):
        return "DB_쿠팡"
    if any(k in fn_lower for k in ["meta", "facebook", "fb", "instagram", "메타", "페이스북", "인스타"]):
        return "DB_메타"
    if any(k in fn_lower for k in ["sales", "revenue", "order", "매출", "판매", "정산", "스마트스토어"]):
        return "DB_매출"

    # 2. DataFrame 컬럼 기반 판별 (한글/영문 시그니처)
    # 네이버 원본 보고서 시그니처
    if any(k in c for c in cols for k in ["총비용", "캠페인 분류", "on/off", "총 전환매출액"]) or (
        any("campaign" in c for c in cols) and any("cost" in c or "spend" in c for c in cols) and any("naver" in c or "keyword" in c or "키워드" in c for c in cols)
    ):
        return "DB_네이버"

    # 메타 원본 보고서 시그니처
    if any(k in c for c in cols for k in ["보고 시작", "지출 금액", "meta 구매 전환값", "결과 기준", "게재결과"]) or (
        any("reporting starts" in c or "amount spent" in c or "ad set" in c or "purchase roas" in c for c in cols)
    ):
        return "DB_메타"

    # 쿠팡 원본 보고서 시그니처
    if any(k in c for c in cols for k in ["광고일자", "집행금액", "캠페인종류", "광고전환매출액", "광고수익률"]) or (
        any("ad spend" in c or "coupang" in c for c in cols)
    ):
        return "DB_쿠팡"

    # 매출 원본 시그니처
    if any(k in c for c in cols for k in ["판매수량", "배송건수", "매출분류", "결제금액", "정산금액", "수량", "매출액"]) or (
        any("quantity" in c or "item count" in c for c in cols)
    ):
        return "DB_매출"

    # 기본 컬럼 매칭
    if any("광고비" in c for c in cols) and any("노출수" in c for c in cols):
        return "DB_네이버"

    return "STANDARD"


def process_uploaded_file_v2(
    file_bytes: bytes,
    filename: str,
    selected_db_type: str,
    target_year: int,
    target_month: str,
    import_mode: str = "overwrite"
) -> Tuple[bool, str, Dict[str, int]]:
    """
    업로드된 엑셀 또는 CSV 파일을 분석하여 네이버/메타/쿠팡/매출 양식을 자동 변환 후 적재합니다.

    Args:
        file_bytes: 파일 바이트
        filename: 파일 이름 (한글, 영문 무관)
        selected_db_type: 사용자 선택 DB ('자동 감지', 'DB_네이버', 'DB_쿠팡', 'DB_메타', 'DB_매출')
        target_year: 날짜 없는 파일에 적용할 기본 연도 (예: 2026)
        target_month: 날짜 없는 파일에 적용할 기본 월 (예: '2026-10')
        import_mode: 'overwrite' (덮어쓰기) 또는 'append' (누적 추가)

    Returns:
        Tuple[bool, str, Dict[str, int]]: (성공 여부, 메시지, 반영 건수)
    """
    updated_counts: Dict[str, int] = {}

    try:
        # 1. 엑셀 파일인 경우
        if filename.lower().endswith((".xlsx", ".xls")):
            excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
            sheet_names = excel_file.sheet_names

            # 다중 시트 통합 엑셀 처리
            matched_sheets: Dict[str, str] = {}
            for sheet in sheet_names:
                for db_key in DB_KEYS:
                    if db_key == sheet or db_key.replace("DB_", "") in sheet:
                        matched_sheets[sheet] = db_key

            if len(matched_sheets) > 1 and selected_db_type in ["자동 감지 (통합 엑셀 / 원본 보고서)", "자동 감지"]:
                for sheet, db_key in matched_sheets.items():
                    df_sheet = excel_file.parse(sheet)
                    cleaned_df = transform_dataframe_to_db(df_sheet, db_key, target_year, target_month)
                    apply_to_db(db_key, cleaned_df, import_mode)
                    updated_counts[db_key] = len(cleaned_df)
                return True, f"통합 엑셀에서 {len(matched_sheets)}개 시트를 자동으로 연동했습니다.", updated_counts

            # 단일 시트 처리
            first_df = excel_file.parse(sheet_names[0])
            return _process_single_df(first_df, selected_db_type, target_year, target_month, import_mode, filename)

        # 2. CSV 파일인 경우
        elif filename.lower().endswith(".csv"):
            try:
                csv_df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8-sig")
            except UnicodeDecodeError:
                csv_df = pd.read_csv(io.BytesIO(file_bytes), encoding="cp949")
            return _process_single_df(csv_df, selected_db_type, target_year, target_month, import_mode, filename)

        else:
            return False, "지원하지 않는 파일 형식입니다. .xlsx 또는 .csv 파일을 업로드해주세요.", {}

    except Exception as e:
        return False, f"파일 처리 중 오류가 발생했습니다: {str(e)}", {}


# 하위 호환성을 위한 함수 별칭
process_uploaded_file = process_uploaded_file_v2


def _process_single_df(
    df: pd.DataFrame,
    selected_db_type: str,
    target_year: int,
    target_month: str,
    import_mode: str,
    filename: str = ""
) -> Tuple[bool, str, Dict[str, int]]:
    """단일 DataFrame에 대해 포맷 감지 및 DB 적재를 수행합니다."""
    # 자동 감지 모드인 경우 파일명 및 헤더 종합 판별
    if "자동 감지" in selected_db_type:
        detected = detect_file_format(df, filename)
        target_db = detected if detected != "STANDARD" else "DB_네이버"
    else:
        target_db = selected_db_type

    transformed_df = transform_dataframe_to_db(df, target_db, target_year, target_month)
    if transformed_df.empty:
        return False, "유효한 데이터 행을 찾을 수 없습니다.", {}

    apply_to_db(target_db, transformed_df, import_mode)
    return True, f"'{target_db}'에 {len(transformed_df)}건의 데이터가 성공적으로 반영되었습니다.", {target_db: len(transformed_df)}


def transform_dataframe_to_db(
    df: pd.DataFrame,
    target_db: str,
    target_year: int,
    target_month: str
) -> pd.DataFrame:
    """대상 DB에 맞게 원본 DataFrame을 정제 및 변환합니다."""
    detected = detect_file_format(df)

    if target_db == "DB_네이버" or detected == "DB_네이버":
        if any("총비용" in c for c in df.columns):
            return parse_naver_ad_report(df, target_year, target_month)
    elif target_db == "DB_메타" or detected == "DB_메타":
        if any("지출 금액" in c or "보고 시작" in c for c in df.columns):
            return parse_meta_ad_report(df, target_year, target_month)
    elif target_db == "DB_쿠팡" or detected == "DB_쿠팡":
        if any("집행금액" in c or "광고일자" in c or "캠페인종류" in c for c in df.columns):
            return parse_coupang_ad_report(df, target_year, target_month)

    # 표준 컬럼 정제 처리
    expected_cols = DB_COLUMNS.get(target_db, [])
    cleaned = df.copy()
    cleaned.columns = [str(c).strip() for c in cleaned.columns]

    # 유사 컬럼 매핑
    alias = {
        "총비용": "광고비", "지출 금액 (KRW)": "광고비", "지출 금액": "광고비", "집행금액": "광고비",
        "노출": "노출수", "조회": "클릭수", "링크 클릭": "클릭수", "클릭": "클릭수",
        "구매": "전환수", "결과": "전환수", "구매건수": "전환수", "총 전환수": "전환수",
        "총 전환매출액": "전환매출", "구매 전환값": "전환매출", "전환매출액": "전환매출",
        "캠페인 분류": "마케팅 분류", "구분": "마케팅 분류", "분류": "마케팅 분류"
    }
    for old_col, new_col in alias.items():
        if old_col in cleaned.columns and new_col not in cleaned.columns:
            cleaned[new_col] = cleaned[old_col]

    if "년도" not in cleaned.columns:
        cleaned["년도"] = target_year
    if "월" not in cleaned.columns:
        cleaned["월"] = target_month

    for col in expected_cols:
        if col not in cleaned.columns:
            cleaned[col] = 0 if col in NUMERIC_COLUMNS else ""

    for col in NUMERIC_COLUMNS:
        if col in cleaned.columns:
            cleaned[col] = cleaned[col].apply(clean_numeric_val)

    cleaned["년도"] = pd.to_numeric(cleaned["년도"], errors="coerce").fillna(target_year).astype(int)
    cleaned["월"] = cleaned["월"].astype(str).str.strip()

    keep_cols = [c for c in expected_cols if c in cleaned.columns]
    return cleaned[keep_cols]


def apply_to_db(db_name: str, new_df: pd.DataFrame, mode: str = "overwrite") -> None:
    """새로운 DataFrame을 데이터베이스에 반영합니다 (overwrite 또는 append)."""
    if mode == "append":
        existing = load_db(db_name)
        combined = pd.concat([existing, new_df], ignore_index=True)
        # 중복 방지 (년도, 월, 캠페인 이름이 동일한 경우 최신 데이터로 업데이트)
        if "캠페인 이름" in combined.columns and "월" in combined.columns:
            combined = combined.drop_duplicates(subset=["년도", "월", "캠페인 이름"], keep="last")
        save_db(db_name, combined)
    else:
        save_db(db_name, new_df)


# ----------------------------------------------------------------------
# 데이터베이스 삭제 및 관리 함수 (요청사항 반영)
# ----------------------------------------------------------------------

def delete_period_data(db_name: str, target_year: int, target_month: Optional[str] = None) -> int:
    """
    지정된 연도 또는 연도+월에 해당하는 레코드를 해당 DB에서 삭제합니다.

    Args:
        db_name: 대상 DB 키
        target_year: 대상 연도
        target_month: 대상 월 (None 또는 '전체'일 시 해당 연도 전체 삭제)

    Returns:
        int: 삭제된 행 수
    """
    df = load_db(db_name)
    if df.empty:
        return 0

    initial_len = len(df)
    if target_month and target_month != "전체":
        df = df[~((df["년도"] == target_year) & (df["월"] == target_month))]
    else:
        df = df[df["년도"] != target_year]

    deleted_count = initial_len - len(df)
    save_db(db_name, df)
    return deleted_count


def clear_db(db_name: str) -> None:
    """지정된 DB의 모든 데이터를 완전히 비우고 빈 테이블로 초기화합니다."""
    expected_cols = DB_COLUMNS.get(db_name, [])
    empty_df = pd.DataFrame(columns=expected_cols)
    save_db(db_name, empty_df)


def clear_all_dbs() -> None:
    """모든 DB를 완전히 비웁니다."""
    for db_key in DB_KEYS:
        clear_db(db_key)


def get_available_filters(data_dict: Dict[str, pd.DataFrame]) -> Tuple[List[int], List[str]]:
    """모든 DB에서 유효한 연도와 월 목록을 반환합니다."""
    years_set = set()
    months_set = set()

    for df in data_dict.values():
        if "년도" in df.columns:
            years_set.update(df["년도"].dropna().astype(int).tolist())
        if "월" in df.columns:
            months_set.update(df["월"].dropna().astype(str).tolist())

    sorted_years = sorted(list(years_set)) if years_set else [2026, 2027]
    sorted_months = sorted([m for m in months_set if m and m != "nan" and len(m) >= 6])
    if not sorted_months:
        sorted_months = ["2026-10", "2027-01"]

    return sorted_years, sorted_months


def export_all_to_excel() -> bytes:
    """전체 DB를 통합 엑셀 파일로 내보냅니다."""
    output = io.BytesIO()
    all_dbs = load_all_dbs()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for db_key, df in all_dbs.items():
            df.to_excel(writer, sheet_name=db_key, index=False)
    return output.getvalue()
