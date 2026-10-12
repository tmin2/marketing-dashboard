"""
마케팅 통합 성과 대시보드 메인 프로그램입니다.
- 상단 수평 네비게이션 메뉴바 (Google Material Top Nav Tabs) 적용
- 타이틀 및 위계 폰트 크기 전체 2포인트 확대
- 회계 표시 서식 적용 (천 단위 콤마, 0원은 '-' 회계 처리)
- 구글 머티리얼 디자인 3 가이드라인 전면 적용
"""

from typing import Dict, List, Optional, Tuple, Any, Union
import os
import sys
import pandas as pd
import numpy as np
import streamlit as st

# 모듈 경로 추가 및 핫 리로드 보장
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.append(CURRENT_DIR)

import importlib
import src.data_manager
import src.metrics_calculator
import src.visualizer
importlib.reload(src.data_manager)
importlib.reload(src.metrics_calculator)
importlib.reload(src.visualizer)

from src.data_manager import (
    DB_KEYS,
    DB_COLUMNS,
    NUMERIC_COLUMNS,
    load_all_dbs,
    load_db,
    save_db,
    clean_numeric_val,
    process_uploaded_file_v2,
    delete_period_data,
    clear_db,
    clear_all_dbs,
    get_available_filters,
    export_all_to_excel,
    TEMPLATES_DIR
)
from src.metrics_calculator import (
    filter_dataframe,
    calculate_summary_metrics,
    get_channel_performance_table,
    get_sales_channel_summary_table,
    get_monthly_trend_table,
    get_yearly_trend_table
)
from src.visualizer import (
    plot_monthly_revenue_and_cost,
    plot_channel_cost_share,
    plot_sales_channel_share,
    plot_pie_category,
    plot_channel_trend_line,
    COLORS
)

# 전역 채널 명칭 매핑 (DB 단어 제거)
CHANNEL_NAMES: Dict[str, str] = {
    "DB_네이버": "네이버 광고",
    "DB_쿠팡": "쿠팡 광고",
    "DB_메타": "메타 광고",
    "DB_매출": "전체 매출"
}

# 페이지 기본 설정 (사이드바 기본 접힘으로 상단 메뉴바에 화면 100% 집중)
st.set_page_config(
    page_title="마케팅 통합 성과 대시보드",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ----------------------------------------------------------------------
# 구글 디자인 가이드 & 상단 메뉴바 CSS 스타일링 (폰트 2포인트 확대 반영)
# ----------------------------------------------------------------------
st.markdown("""
<style>
    @import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css");

    /* 이숲(esoop.co.kr) 감성의 세련된 Pretendard 타이포그래피 및 캔버스 배경 */
    html, body, [class*="css"], .stApp {
        font-family: "Pretendard Variable", Pretendard, -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif !important;
        background-color: #F8FAFB !important;
        color: #111827;
        letter-spacing: -0.3px;
    }

    /* 메인 컨테이너 여백 최적화 (가로 폭 100% 활용) */
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2.5rem;
        padding-left: 2.0rem;
        padding-right: 2.0rem;
    }

    /* 이숲 스타일 타이포그래피 (딥 포레스트 컬러 & 대제목 44px 대형 폰트 완벽 유지) */
    h1 {
        font-family: "Pretendard Variable", Pretendard, sans-serif !important;
        font-size: 1.72rem !important;
        font-weight: 700 !important;
        color: #06291B !important;
        margin-bottom: 0.35rem !important;
        letter-spacing: -0.4px !important;
    }
    /* 최상단 메인 대시보드 타이틀 (약 38px / 2.35rem 시원하게 확대 & 딥 포레스트 포인트) */
    .main-dashboard-title,
    h1.main-dashboard-title,
    h2.main-dashboard-title,
    h3.main-dashboard-title {
        font-family: "Pretendard Variable", Pretendard, sans-serif !important;
        font-size: 2.35rem !important;
        font-weight: 800 !important;
        color: #06291B !important;
        margin-top: 0 !important;
        margin-bottom: 0.35rem !important;
        letter-spacing: -0.8px !important;
        line-height: 1.25 !important;
        display: block !important;
    }
    h2 {
        font-family: "Pretendard Variable", Pretendard, sans-serif !important;
        font-size: 1.42rem !important;
        font-weight: 700 !important;
        color: #0A3322 !important;
        margin-top: 1.2rem !important;
        margin-bottom: 0.5rem !important;
        letter-spacing: -0.3px !important;
    }
    h3 {
        font-family: "Pretendard Variable", Pretendard, sans-serif !important;
        font-size: 1.22rem !important;
        font-weight: 600 !important;
        color: #153E2B !important;
        margin-top: 0.8rem !important;
        margin-bottom: 0.4rem !important;
        letter-spacing: -0.2px !important;
    }
    .sub-caption {
        font-size: 1.02rem !important;
        color: #4B6356 !important;
        margin-bottom: 0.6rem !important;
        font-weight: 500 !important;
    }
    p, span, label {
        font-size: 0.90rem !important;
        color: #374151;
    }

    /* ======================================================= */
    /* 🚀 상단 고정 반응형 스마트 네비게이션 헤더 (Fixed GNB)      */
    /* ======================================================= */
    /* Streamlit 기본 헤더 투명화 & 메뉴 버튼 기능 유지 */
    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 0 !important;
        min-height: 0 !important;
        z-index: 1000000 !important;
        pointer-events: none !important;
    }
    header[data-testid="stHeader"] * {
        pointer-events: auto !important;
    }

    /* 상단바 컨테이너를 뷰포트 상단에 영구 고정 (스크롤 시 무조건 100% 따라옴) */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor),
    div.block-container > div[data-testid="stVerticalBlock"] > div:has(.fixed-top-nav-anchor) {
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        right: 0 !important;
        width: 100% !important;
        z-index: 999990 !important;
        background: rgba(248, 250, 251, 0.96) !important;
        backdrop-filter: blur(16px) !important;
        -webkit-backdrop-filter: blur(16px) !important;
        border: none !important;
        border-bottom: 1px solid #DFE7E2 !important;
        box-shadow: 0 4px 20px rgba(6, 41, 27, 0.08) !important;
        padding: 6px 2.0rem 6px 2.0rem !important;
        box-sizing: border-box !important;
        border-radius: 0 !important;
    }

    /* 상단 고정 앵커 컨테이너 자체를 완전 무높이 처리 (위쪽 빈 공간 원천 제거) */
    div[data-testid="element-container"]:has(.fixed-top-nav-anchor) {
        display: none !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    /* 헤더 내부 테두리 및 패딩 초기화 & 간격 초밀착 */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor) {
        border: none !important;
        border-bottom: 1px solid #DFE7E2 !important;
        border-radius: 0 !important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor) > div {
        padding: 0 !important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor) [data-testid="stVerticalBlock"] {
        gap: 0 !important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor) [data-testid="element-container"] {
        margin: 0 !important;
    }

    /* ======================================================= */
    /* 🌿 라디오 원형 점(circle)만 핀포인트 제거 (텍스트 100% 보존!) */
    /* ======================================================= */
    div[data-testid="stRadio"] input[type="radio"] + div,
    div[role="radiogroup"] input[type="radio"] + div,
    div[data-testid="stRadio"] [data-baseweb="radio"] input[type="radio"] + div,
    div[role="radiogroup"] [data-baseweb="radio"] input[type="radio"] + div {
        display: none !important;
        width: 0 !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        opacity: 0 !important;
        visibility: hidden !important;
        position: absolute !important;
        pointer-events: none !important;
    }
    div[data-testid="stRadio"] input[type="radio"],
    div[role="radiogroup"] input[type="radio"] {
        display: none !important;
        width: 0 !important;
        height: 0 !important;
        opacity: 0 !important;
        visibility: hidden !important;
        position: absolute !important;
        pointer-events: none !important;
    }
    /* 라디오 텍스트 레이블 가시성 100% 강제 보장 */
    div[data-testid="stRadio"] label p,
    div[role="radiogroup"] label p,
    div[data-testid="stRadio"] label span,
    div[role="radiogroup"] label span,
    div[data-testid="stRadio"] [data-baseweb="radio"] p,
    div[role="radiogroup"] [data-baseweb="radio"] p {
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
    }

    /* ======================================================= */
    /* 🎨 상단 수평 네비게이션 탭 바 & 각 탭별 고유 브랜드 색상   */
    /* ======================================================= */
    div[role="radiogroup"][aria-label="메뉴 선택"],
    div[data-testid="stRadio"]:first-child > div[role="radiogroup"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: wrap !important;
        background-color: #EEF3F0 !important;
        padding: 5px 6px !important;
        border-radius: 12px !important;
        gap: 6px !important;
        margin-top: 0.3rem !important;
        margin-bottom: 0.6rem !important;
        border: 1px solid #DFE7E2 !important;
        box-shadow: inset 0 1px 2px rgba(6, 41, 27, 0.04) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label {
        background-color: #FFFFFF !important;
        padding: 8px 18px !important;
        border-radius: 9px !important;
        margin: 0 !important;
        cursor: pointer !important;
        transition: all 0.22s cubic-bezier(0.16, 1, 0.3, 1) !important;
        border: 1.5px solid #E2E8F0 !important;
        box-shadow: 0 1px 3px rgba(6, 41, 27, 0.04) !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label p {
        font-size: 0.98rem !important;
        font-weight: 600 !important;
        color: #4B6356 !important;
        letter-spacing: -0.2px !important;
        margin: 0 !important;
        display: block !important;
        visibility: visible !important;
    }

    /* 1. 🏠 통합 대시보드 (이숲 에메랄드) */
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="통합 대시보드"]):hover {
        border-color: #00C56F !important;
        background-color: #F0FDF4 !important;
        transform: translateY(-1px) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="통합 대시보드"]:checked) {
        background: linear-gradient(135deg, #00D87B 0%, #00A35C 100%) !important;
        border-color: #00A35C !important;
        box-shadow: 0 3px 12px rgba(0, 197, 111, 0.35) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="통합 대시보드"]:checked) p {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* 2. 💚 네이버 광고 (네이버 그린) */
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="네이버"]):hover {
        border-color: #03C75A !important;
        background-color: #F0FDF4 !important;
        transform: translateY(-1px) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="네이버"]:checked) {
        background: linear-gradient(135deg, #03C75A 0%, #029E47 100%) !important;
        border-color: #029E47 !important;
        box-shadow: 0 3px 12px rgba(3, 199, 90, 0.35) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="네이버"]:checked) p {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* 3. 🚀 쿠팡 광고 (쿠팡 스칼렛 레드) */
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="쿠팡"]):hover {
        border-color: #FF334B !important;
        background-color: #FFF1F2 !important;
        transform: translateY(-1px) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="쿠팡"]:checked) {
        background: linear-gradient(135deg, #FF334B 0%, #E01A33 100%) !important;
        border-color: #E01A33 !important;
        box-shadow: 0 3px 12px rgba(255, 51, 75, 0.35) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="쿠팡"]:checked) p {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* 4. 🔵 메타 광고 (메타 로열 블루) */
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="메타"]):hover {
        border-color: #0084FF !important;
        background-color: #EFF6FF !important;
        transform: translateY(-1px) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="메타"]:checked) {
        background: linear-gradient(135deg, #0084FF 0%, #0066CC 100%) !important;
        border-color: #0066CC !important;
        box-shadow: 0 3px 12px rgba(0, 132, 255, 0.35) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="메타"]:checked) p {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* 5. 💰 전체 매출 (골드 앰버) */
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="전체 매출"]):hover {
        border-color: #F59E0B !important;
        background-color: #FFFBEB !important;
        transform: translateY(-1px) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="전체 매출"]:checked) {
        background: linear-gradient(135deg, #F59E0B 0%, #D97706 100%) !important;
        border-color: #D97706 !important;
        box-shadow: 0 3px 12px rgba(245, 158, 11, 0.35) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="전체 매출"]:checked) p {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* 6. 🛠️ 데이터 관리 (슬레이트 퍼플) */
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="데이터 관리"]):hover {
        border-color: #7C3AED !important;
        background-color: #F5F3FF !important;
        transform: translateY(-1px) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="데이터 관리"]:checked) {
        background: linear-gradient(135deg, #7C3AED 0%, #6D28D9 100%) !important;
        border-color: #6D28D9 !important;
        box-shadow: 0 3px 12px rgba(124, 58, 237, 0.35) !important;
    }
    div[role="radiogroup"][aria-label="메뉴 선택"] label:has(input[value*="데이터 관리"]:checked) p {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* ======================================================= */
    /* 📱 슬림(접힌) 모드 상단바 및 우측 펼치기 버튼 컴팩트 정렬 */
    /* ======================================================= */
    div[data-testid="stColumn"] div.stButton > button[key="btn_expand_nav"] {
        height: 36px !important;
        min-height: 36px !important;
        max-height: 36px !important;
        font-size: 0.86rem !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        margin: 0 !important;
        padding: 0 12px !important;
        background-color: #FFFFFF !important;
        border: 1px solid #DFE7E2 !important;
        color: #06291B !important;
        box-shadow: 0 1px 3px rgba(6, 41, 27, 0.04) !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stColumn"] div.stButton > button[key="btn_expand_nav"]:hover {
        border-color: #00C56F !important;
        color: #00A35C !important;
        background-color: #F0FDF4 !important;
    }

    /* ======================================================= */
    /* 📱 모바일 환경 극대화 최적화 (화면의 반 이하로 헤더 압축) */
    /* ======================================================= */
    @media (max-width: 768px) {
        /* PC용 빈 공간(28px spacer) 모바일에서 완전 숨김 */
        .pc-spacer {
            display: none !important;
            height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        /* 메인 컨테이너 최상단 및 좌우 여백 모바일 슬림화 */
        .block-container {
            padding-top: 4px !important;
            padding-left: 0.4rem !important;
            padding-right: 0.4rem !important;
        }
        /* 모바일에서 상단바가 화면의 절반을 가리지 않도록 relative 배치 및 상하 여백 극소화 */
        div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor) {
            position: relative !important;
            padding: 4px 6px !important;
            margin-bottom: 4px !important;
            border-radius: 8px !important;
            box-shadow: 0 1px 4px rgba(6, 41, 27, 0.04) !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor) div[data-testid="stVerticalBlock"] {
            gap: 2px !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor) div[data-testid="stHorizontalBlock"] {
            gap: 2px !important;
            margin: 0 !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:has(.fixed-top-nav-anchor) div[data-testid="stColumn"] {
            padding: 0 !important;
            margin: 0 !important;
        }

        /* 📌 타이틀 위아래 여백 완전 최소화 (여백 낭비 박멸) */
        h1.main-dashboard-title {
            font-size: 1.05rem !important;
            line-height: 1.15 !important;
            margin: 0 !important;
            padding: 0 !important;
            letter-spacing: -0.4px !important;
        }
        p.sub-caption {
            display: none !important;
        }

        /* 📌 상단 컨트롤 4개 (연도, 월, 백업, 접기) 모바일 2x2 그리드 칼배치 */
        div[data-testid="stColumn"]:has(.top-controls-wrapper-anchor) div[data-testid="stHorizontalBlock"] {
            display: grid !important;
            grid-template-columns: 1fr 1fr !important;
            gap: 4px !important;
            width: 100% !important;
            margin: 2px 0 !important;
        }
        div[data-testid="stColumn"]:has(.top-controls-wrapper-anchor) div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
            width: 100% !important;
            min-width: 0 !important;
            flex: none !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        div[data-testid="stColumn"]:has(.top-controls-wrapper-anchor) label {
            font-size: 0.70rem !important;
            font-weight: 600 !important;
            margin-bottom: 1px !important;
            min-height: auto !important;
            height: auto !important;
            line-height: 1.1 !important;
            color: #4B6356 !important;
        }
        div[data-testid="stColumn"]:has(.top-controls-wrapper-anchor) div[data-testid="stSelectbox"] div[data-baseweb="select"] {
            min-height: 30px !important;
            height: 30px !important;
            padding: 0 4px !important;
        }
        div[data-testid="stColumn"]:has(.top-controls-wrapper-anchor) div[data-testid="stSelectbox"] div[data-baseweb="select"] * {
            font-size: 0.78rem !important;
        }
        div[data-testid="stColumn"]:has(.top-controls-wrapper-anchor) button {
            min-height: 30px !important;
            height: 30px !important;
            font-size: 0.76rem !important;
            padding: 0 4px !important;
            margin: 0 !important;
        }

        /* 📌 모바일 메뉴 탭: 3열 2행 앱 스타일 칩 버튼 & 네이버/쿠팡/메타 광고 2번째 줄 배치 */
        div[role="radiogroup"][aria-label="메뉴 선택"],
        div[data-testid="stRadio"] > div[role="radiogroup"] {
            display: grid !important;
            grid-template-columns: repeat(3, 1fr) !important;
            gap: 3px !important;
            padding: 3px !important;
            margin-top: 3px !important;
            margin-bottom: 2px !important;
            background-color: #F1F5F3 !important;
            border-radius: 8px !important;
        }
        div[role="radiogroup"][aria-label="메뉴 선택"] label,
        div[data-testid="stRadio"] > div[role="radiogroup"] label,
        div[role="radiogroup"][aria-label="메뉴 선택"] > *,
        div[data-testid="stRadio"] > div[role="radiogroup"] > * {
            padding: 4px 1px !important;
            margin: 0 !important;
            min-height: 29px !important;
            height: 29px !important;
            border-radius: 6px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            text-align: center !important;
            box-shadow: none !important;
        }
        /* 라디오 원형 체크 아이콘 숨김 */
        div[role="radiogroup"] label div:first-child {
            display: none !important;
        }
        div[role="radiogroup"] label p {
            font-size: 0.70rem !important;
            font-weight: 700 !important;
            margin: 0 !important;
            line-height: 1.15 !important;
            letter-spacing: -0.5px !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
        }

        /* [모바일 전용 정렬] 1번째 줄: 종합/매출/관리, 2번째 줄: 3대 광고 매체 (네이버, 쿠팡, 메타) */
        div[role="radiogroup"] > *:nth-child(1),
        div[role="radiogroup"] label:nth-of-type(1),
        div[role="radiogroup"] label:has(input[value*="통합"]) {
            grid-row: 1 !important;
            grid-column: 1 !important;
            order: 1 !important;
        }
        div[role="radiogroup"] > *:nth-child(5),
        div[role="radiogroup"] label:nth-of-type(5),
        div[role="radiogroup"] label:has(input[value*="매출"]) {
            grid-row: 1 !important;
            grid-column: 2 !important;
            order: 2 !important;
        }
        div[role="radiogroup"] > *:nth-child(6),
        div[role="radiogroup"] label:nth-of-type(6),
        div[role="radiogroup"] label:has(input[value*="관리"]) {
            grid-row: 1 !important;
            grid-column: 3 !important;
            order: 3 !important;
        }
        div[role="radiogroup"] > *:nth-child(2),
        div[role="radiogroup"] label:nth-of-type(2),
        div[role="radiogroup"] label:has(input[value*="네이버"]) {
            grid-row: 2 !important;
            grid-column: 1 !important;
            order: 4 !important;
        }
        div[role="radiogroup"] > *:nth-child(3),
        div[role="radiogroup"] label:nth-of-type(3),
        div[role="radiogroup"] label:has(input[value*="쿠팡"]) {
            grid-row: 2 !important;
            grid-column: 2 !important;
            order: 5 !important;
        }
        div[role="radiogroup"] > *:nth-child(4),
        div[role="radiogroup"] label:nth-of-type(4),
        div[role="radiogroup"] label:has(input[value*="메타"]) {
            grid-row: 2 !important;
            grid-column: 3 !important;
            order: 6 !important;
        }

        /* 스마트 안내 캡션 박스 숨김 (모바일 화면 대폭 절약) */
        .smart-tab-guide-box {
            display: none !important;
        }
    }

    /* ======================================================= */
    /* 📊 모든 표(DataFrame / DataEditor) 전역 오토사이즈 스타일 */
    /* ======================================================= */
    div[data-testid="stDataFrame"],
    div[data-testid="stDataEditor"] {
        width: 100% !important;
    }
    div[data-testid="stDataFrame"] > div,
    div[data-testid="stDataEditor"] > div {
        width: 100% !important;
    }

    /* Expander 모던 카드 스타일링 */
    div[data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 1px solid #E5EBE7 !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 8px rgba(6, 41, 27, 0.03) !important;
    }

    /* ======================================================= */
    /* 📐 파일 업로드 영역 줄 맞춤 & 정렬 최적화 (1행 & 2행)   */
    /* ======================================================= */
    /* 1행 옵션 설정: 라벨 높이 24px 칼일치 */
    div[data-testid="stExpander"] div[data-testid="stSelectbox"] label,
    div[data-testid="stExpander"] div[data-testid="stTextInput"] label,
    div[data-testid="stExpander"] div[data-testid="stRadio"] label[data-testid="stWidgetLabel"],
    div[data-testid="stExpander"] div[data-testid="stRadio"] > label {
        height: 24px !important;
        min-height: 24px !important;
        max-height: 24px !important;
        line-height: 24px !important;
        margin-bottom: 6px !important;
        margin-top: 0 !important;
        padding: 0 !important;
        font-size: 0.90rem !important;
        font-weight: 600 !important;
        color: #374151 !important;
        display: flex !important;
        align-items: center !important;
        box-sizing: border-box !important;
    }

    /* 1행 입력 컨트롤 박스 높이 42px 칼정렬 (셀렉트, 텍스트인풋, 라디오) */
    div[data-testid="stExpander"] div[data-baseweb="select"] > div,
    div[data-testid="stExpander"] div[data-baseweb="input"],
    div[data-testid="stExpander"] div[data-baseweb="input"] > div {
        height: 42px !important;
        min-height: 42px !important;
        max-height: 42px !important;
        border-radius: 9px !important;
        border-color: #DFE7E2 !important;
        box-sizing: border-box !important;
        margin: 0 !important;
    }
    div[data-testid="stExpander"] div[data-testid="stTextInput"] input {
        height: 100% !important;
        min-height: 100% !important;
        max-height: 100% !important;
        line-height: 40px !important;
        padding: 0 12px !important;
        box-sizing: border-box !important;
        margin: 0 !important;
    }

    /* 1행 반영 방식 라디오 그룹 */
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] {
        height: 42px !important;
        min-height: 42px !important;
        max-height: 42px !important;
        margin: 0 !important;
        margin-top: 0 !important;
        padding: 3px !important;
        border-radius: 9px !important;
        background-color: #EEF3F0 !important;
        border: 1px solid #DFE7E2 !important;
        display: flex !important;
        flex-direction: row !important;
        align-items: stretch !important;
        justify-content: space-around !important;
        box-sizing: border-box !important;
        box-shadow: none !important;
        gap: 4px !important;
    }
    /* 1행 라디오 원형 동그라미 제거 (텍스트는 100% 보존) */
    div[data-testid="stExpander"] div[data-testid="stRadio"] input[type="radio"] + div,
    div[data-testid="stExpander"] div[data-testid="stRadio"] input[type="radio"] {
        display: none !important;
        width: 0 !important;
        height: 0 !important;
        opacity: 0 !important;
        visibility: hidden !important;
        position: absolute !important;
        pointer-events: none !important;
    }
    /* 1행 라디오 탭 아이템 버튼 스타일 */
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] label,
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] [data-baseweb="radio"] {
        flex: 1 1 0% !important;
        height: 34px !important;
        min-height: 34px !important;
        max-height: 34px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        padding: 0 10px !important;
        margin: 0 !important;
        border-radius: 7px !important;
        cursor: pointer !important;
        background-color: transparent !important;
        border: 1px solid transparent !important;
        transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1) !important;
        box-sizing: border-box !important;
    }
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] label p,
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] [data-baseweb="radio"] p,
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] span {
        font-size: 0.90rem !important;
        font-weight: 600 !important;
        color: #4B6356 !important;
        margin: 0 !important;
        line-height: 1 !important;
        white-space: nowrap !important;
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
    }
    /* 덮어쓰기 선택 시 (경고 오렌지) */
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input[value*="덮어쓰기"]:checked) {
        background-color: #FFF7ED !important;
        border: 1.5px solid #F97316 !important;
        box-shadow: 0 1px 4px rgba(249, 115, 22, 0.2) !important;
    }
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input[value*="덮어쓰기"]:checked) p {
        color: #EA580C !important;
        font-weight: 700 !important;
    }
    /* 누적 추가 선택 시 (안전 에메랄드) */
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input[value*="누적 추가"]:checked) {
        background-color: #F0FDF4 !important;
        border: 1.5px solid #00C56F !important;
        box-shadow: 0 1px 4px rgba(0, 197, 111, 0.2) !important;
    }
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input[value*="누적 추가"]:checked) p {
        color: #00A35C !important;
        font-weight: 700 !important;
    }
    /* 선택된 활성 탭 (화이트 카드 + 이숲 에메랄드) */
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked),
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] [data-baseweb="radio"]:has(input:checked) {
        background-color: #ffffff !important;
        border-color: #D1E7DD !important;
        box-shadow: 0 1px 4px rgba(0, 197, 111, 0.14) !important;
    }
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) p,
    div[data-testid="stExpander"] div[data-testid="stRadio"] div[role="radiogroup"] [data-baseweb="radio"]:has(input:checked) p {
        color: #00A35C !important;
        font-weight: 700 !important;
    }

    /* 2행 라벨 높이 22px 일치 */
    div[data-testid="stFileUploader"] label[data-testid="stWidgetLabel"],
    .upload-apply-label {
        height: 22px !important;
        line-height: 22px !important;
        margin-bottom: 4px !important;
        padding: 0 !important;
        font-size: 0.90rem !important;
        font-weight: 600 !important;
        color: #374151 !important;
        display: flex !important;
        align-items: center !important;
    }

    /* 2행 파일 드롭존 박스 높이 68px 칼정렬 (모던 에메랄드 대시보드 테두리) */
    div[data-testid="stFileUploader"] section[data-testid="stFileUploaderDropzone"] {
        height: 68px !important;
        min-height: 68px !important;
        padding-top: 10px !important;
        padding-bottom: 10px !important;
        border-radius: 10px !important;
        border: 1.5px dashed #B8E4CD !important;
        background-color: #FDFEFE !important;
        box-sizing: border-box !important;
        transition: border-color 0.2s ease, background-color 0.2s ease !important;
    }
    div[data-testid="stFileUploader"] section[data-testid="stFileUploaderDropzone"]:hover {
        border-color: #00C56F !important;
        background-color: #F6FCF9 !important;
    }
    /* 파일 드롭존 내부 기본 200MB 및 형식 안내 텍스트 완전 숨김 (업로드 버튼 옆으로 이전) */
    div[data-testid="stFileUploader"] section[data-testid="stFileUploaderDropzone"] small,
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"],
    div[data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] * {
        display: none !important;
        visibility: hidden !important;
        font-size: 0 !important;
        height: 0 !important;
    }
    /* 업로드 버튼 옆에 연하고 작은 글씨로 (최대 200MB) 표시 */
    div[data-testid="stFileUploader"] section[data-testid="stFileUploaderDropzone"] button::after {
        content: " (최대 200MB)";
        font-size: 0.74rem !important;
        color: #94A3B8 !important;
        font-weight: 400 !important;
        margin-left: 6px !important;
    }

    /* 2행 데이터 즉시 적용 버튼: 이숲 시그니처 에메랄드 그라데이션 & 글씨 순백색(#ffffff) 강제 적용 */
    div[data-testid="stExpander"] div.stButton > button {
        height: 68px !important;
        min-height: 68px !important;
        background: linear-gradient(135deg, #00D87B 0%, #00C56F 100%) !important;
        border: none !important;
        border-radius: 10px !important;
        box-shadow: 0 3px 12px rgba(0, 197, 111, 0.32) !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1) !important;
        margin: 0 !important;
    }
    /* 버튼 텍스트 및 아이콘 완전한 흰색(#ffffff) 보장 */
    div[data-testid="stExpander"] div.stButton > button * {
        color: #ffffff !important;
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        letter-spacing: -0.2px !important;
    }
    div[data-testid="stExpander"] div.stButton > button:hover {
        background: linear-gradient(135deg, #00C56F 0%, #00B868 100%) !important;
        box-shadow: 0 5px 18px rgba(0, 197, 111, 0.42) !important;
        transform: translateY(-1px) !important;
    }
    div[data-testid="stExpander"] div.stButton > button:hover * {
        color: #ffffff !important;
    }

    /* 2행 우측 연동 방식 안내 박스 (높이 68px 일치 & 이숲 소프트 민트 배경) */
    .upload-info-box {
        height: 68px !important;
        min-height: 68px !important;
        background-color: #F4FBF7;
        border: 1px solid #D0EFE0;
        border-radius: 10px;
        padding: 8px 12px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        box-sizing: border-box;
    }

    /* 이숲(esoop.co.kr) 모던 KPI 카드 (부드러운 라운딩 & 호버 인터랙션) */
    .g-kpi-card {
        background-color: #ffffff;
        border: 1px solid #E5EBE7;
        border-radius: 12px;
        padding: 12px 16px;
        margin-bottom: 10px;
        box-shadow: 0 2px 6px rgba(6, 41, 27, 0.03);
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        display: flex;
        flex-direction: column;
        justify-content: center;
        min-height: 80px;
    }
    .g-kpi-card:hover {
        border-color: #00C56F;
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0, 197, 111, 0.12);
    }
    .g-kpi-title {
        font-size: 0.78rem !important;
        font-weight: 600;
        color: #52635B;
        margin-bottom: 3px;
        letter-spacing: 0.1px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .g-kpi-value {
        font-family: "Pretendard Variable", Pretendard, sans-serif !important;
        font-size: 1.25rem !important;
        font-weight: 800;
        color: #06291B;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        line-height: 1.25;
        letter-spacing: -0.3px;
    }
    .g-kpi-sub {
        font-size: 0.72rem !important;
        color: #00A35C;
        background-color: #F0F9F4;
        padding: 2px 6px;
        border-radius: 4px;
        margin-top: 4px;
        display: inline-block;
        width: fit-content;
        font-weight: 600;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    /* 이숲 시그니처 소프트 민트 가이드 안내 박스 */
    .g-guide-box {
        background-color: #F4FBF7;
        border: 1px solid #D0EFE0;
        border-left: 5px solid #00C56F;
        padding: 14px 18px;
        border-radius: 10px;
        margin-top: 18px;
        font-size: 0.86rem !important;
        color: #24382E;
        line-height: 1.6;
        box-shadow: 0 2px 6px rgba(0, 197, 111, 0.04);
    }
    .g-guide-box code {
        background-color: #E3F8ED !important;
        color: #00874C !important;
        border: 1px solid #C4EED8 !important;
        padding: 2px 6px !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
    }

    /* 전역 모던 버튼 스타일링 */
    div.stButton > button {
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.90rem;
        padding: 7px 18px;
        transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
    }
    /* Primary 버튼 (이숲 에메랄드 그린 테마) */
    div.stButton > button[kind="primary"] {
        background-color: #00C56F !important;
        border-color: #00C56F !important;
        color: #ffffff !important;
        box-shadow: 0 2px 8px rgba(0, 197, 111, 0.25) !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #00B868 !important;
        border-color: #00B868 !important;
        box-shadow: 0 4px 12px rgba(0, 197, 111, 0.35) !important;
        transform: translateY(-1px);
    }

    /* Streamlit 테이블 & 데이터 에디터 카드 스타일링 */
    [data-testid="stDataFrame"] {
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(6, 41, 27, 0.02);
    }
</style>
""", unsafe_allow_html=True)


# ----------------------------------------------------------------------
# 회계 서식 (Accounting Format) 변환 함수
# ----------------------------------------------------------------------

def format_accounting(
    val: Any,
    is_currency: bool = False,
    unit: str = "",
    decimals: Optional[int] = None
) -> str:
    """
    회계 처리 서식 변환 함수:
    - 0 또는 빈값 -> '-' (대시 처리)
    - 양수/음수 -> 천 단위 콤마(,) 적용 (예: 1,404,600)
    - 통화 처리 -> ₩1,404,600 (단, 0원일 때는 '-' 단독 표기)
    - 원화(KRW) 금액(광고비, 매출, CPA, CPC, 객단가 등)은 소수점 없는 정수 원(₩) 표기
    """
    if pd.isna(val) or val is None or str(val).strip() in ["", "nan", "None"]:
        return "-"
    try:
        clean_str = str(val).replace(",", "").replace("₩", "").replace("%", "").strip()
        num = float(clean_str)
        if abs(num) < 1e-9:
            return "-"

        if decimals is not None:
            if decimals == 0:
                formatted = f"{int(round(num)):,}"
            else:
                formatted = f"{num:,.{decimals}f}"
        elif is_currency:
            # 원화(KRW)는 센트가 없으므로 정수 반올림 표기 (CPA 소수점 제거 요청 반영)
            formatted = f"{int(round(num)):,}"
        elif num.is_integer():
            formatted = f"{int(num):,}"
        else:
            formatted = f"{num:,.2f}"

        if is_currency:
            return f"₩{formatted}"
        if unit:
            return f"{formatted}{unit}"
        return formatted
    except (ValueError, TypeError):
        return str(val) if str(val).strip() else "-"


def format_table_for_accounting_view(df: pd.DataFrame) -> pd.DataFrame:
    """데이터프레임 내의 모든 수치 컬럼을 회계 서식(천단위 콤마, 0은 '-') 문자열로 변환합니다."""
    display_df = df.copy()
    for col in display_df.columns:
        if col in ["광고비", "전환매출", "배송비"]:
            display_df[col] = display_df[col].apply(lambda x: format_accounting(x, is_currency=False, decimals=0))
        elif col in ["노출수", "클릭수", "전환수", "판매수량", "배송건수"]:
            display_df[col] = display_df[col].apply(lambda x: format_accounting(x, is_currency=False, decimals=0))
        elif col == "ROAS":
            display_df[col] = display_df[col].apply(lambda x: format_accounting(x, unit="x", decimals=2))
        elif col in ["CPA", "CPC", "객단가"]:
            display_df[col] = display_df[col].apply(lambda x: format_accounting(x, is_currency=False, decimals=0))
        elif col in ["CTR", "CVR"]:
            display_df[col] = display_df[col].apply(lambda x: format_accounting(x, unit="%", decimals=2))
    return display_df


def render_google_kpi_card(title: str, value: str, sub: str = "", accent_color: str = "#00C56F") -> None:
    """이숲(esoop.co.kr) 스타일의 모던 인터랙티브 KPI 카드를 렌더링합니다."""
    sub_html = f"<div class='g-kpi-sub'>{sub}</div>" if sub else ""
    st.markdown(f"""
    <div class="g-kpi-card" style="border-top: 4.5px solid {accent_color};" title="{title}: {value}">
        <div class="g-kpi-title">{title}</div>
        <div class="g-kpi-value">{value}</div>
        {sub_html}
    </div>
    """, unsafe_allow_html=True)


# 각 성과 지표별 고유 컬러 매핑 (KPI 카드 및 시각화 테마와 1:1 완벽 일치)
INDICATOR_HEADER_COLORS: Dict[str, str] = {
    "연도": "#06291B",        # 이숲 딥 포레스트
    "월": "#06291B",          # 이숲 딥 포레스트
    "채널": "#0084FF",        # 메타 블루
    "판매 채널": "#00C56F",    # 에메랄드 그린
    "광고비": "#FF5E57",       # 코랄 레드
    "총 광고비": "#FF5E57",     # 코랄 레드
    "매출": "#00C56F",         # 이숲 에메랄드 그린
    "총 매출": "#00C56F",       # 이숲 에메랄드 그린
    "전환매출": "#00C56F",     # 이숲 에메랄드 그린
    "매출액": "#00C56F",       # 이숲 에메랄드 그린
    "ROAS": "#3B82F6",        # 로열 블루 (선명한 색상)
    "CPA": "#F59E0B",         # 웜 앰버 골드
    "객단가": "#8B5CF6",       # 모던 바이올렛
    "CPC": "#EC4899",         # 로즈 핑크
    "CTR": "#10B981",         # 브라이트 민트
    "CVR": "#0284C7",         # 스카이 블루
    "노출수": "#64748B",       # 슬레이트
    "클릭수": "#3B82F6",       # 블루
    "전환수": "#10B981",       # 에메랄드
    "판매수량": "#059669",      # 딥 에메랄드
    "배송건수": "#6366F1",      # 인디고
    "배송비": "#F97316",        # 오렌지
}


def render_styled_accounting_table(
    df: pd.DataFrame,
    highlight_col: Optional[str] = None,
    highlight_val: Optional[Union[str, List[str]]] = None,
    total_row: Optional[Dict[str, Any]] = None
) -> None:
    """
    각 지표의 회계 서식이 적용된 모던 이숲 스타일 테이블을 렌더링합니다.
    - highlight_val: 단일 문자열 또는 문자열 리스트를 지원하여 여러 행을 동시에 조건부 서식 하이라이트합니다.
    - 선택된 행의 첫 번째 컬럼 앞에는 '✔ ' 체크 표시가 적용됩니다.
    - total_row: 하단에 합계(선택 합계 또는 전체 합계) 행을 명확한 볼드체와 에메랄드 상단 라인으로 표기합니다.
    """
    if df.empty:
        st.info("표시할 데이터가 없습니다.")
        return

    # 타겟 하이라이트 값 리스트 정규화
    target_vals: List[str] = []
    if highlight_val is not None:
        if isinstance(highlight_val, list):
            target_vals = [str(v).strip() for v in highlight_val if str(v).strip()]
        elif str(highlight_val).strip():
            target_vals = [str(highlight_val).strip()]

    # 테이블 헤더 HTML 구성 (플랫폼 확인용 외 불필요한 색상 라인 제거, 깔끔한 헤더)
    th_html_list = []
    for i, col in enumerate(df.columns):
        align = "left" if i == 0 else "right"
        th_html = (
            f"<th style='"
            f"background-color: #F8FAFB; "
            f"padding: 12px 14px; "
            f"text-align: {align}; "
            f"border-bottom: 2px solid #E2E8F0; "
            f"border-right: 1px solid #F1F5F9; "
            f"white-space: nowrap; "
            f"font-size: 0.88rem; "
            f"font-weight: 700; "
            f"color: #1E293B;'>"
            f"{col}"
            f"</th>"
        )
        th_html_list.append(th_html)
    th_rows = "".join(th_html_list)

    # 테이블 바디 HTML 구성 (조건부 서식 하이라이트 적용)
    tr_html_list = []
    for _, row in df.iterrows():
        is_highlighted = False
        if highlight_col and target_vals and highlight_col in df.columns:
            cell_val = str(row[highlight_col]).strip()
            c_clean = cell_val.replace(" ", "").replace("년", "").replace("월", "").lower()
            for tv in target_vals:
                t_clean = tv.replace(" ", "").replace("년", "").replace("월", "").lower()
                if c_clean == t_clean or cell_val == tv or tv in cell_val or cell_val in tv:
                    is_highlighted = True
                    break

        row_bg = "#E6FBF1" if is_highlighted else "transparent"
        row_border_left = "5.5px solid #00C56F" if is_highlighted else "5.5px solid transparent"
        row_shadow = "box-shadow: inset 0 1px 0 #00C56F, inset 0 -1px 0 #00C56F;" if is_highlighted else ""

        td_html_list = []
        for i, col in enumerate(df.columns):
            val = str(row[col])
            # 선택된 행의 첫 번째 열에는 사용자가 손으로 그린 체크(V) 형태의 마크 부여
            if i == 0 and is_highlighted:
                val = f"<span style='color: #00A35C; font-weight: 800; margin-right: 4px;'>✔</span> {val}"
            align = "left" if i == 0 else "right"
            cell_weight = "700" if (i == 0 or is_highlighted) else "400"
            cell_color = "#06291B" if is_highlighted else "#334155"
            cell_border_l = "border-left: 5.5px solid #00C56F; " if (i == 0 and is_highlighted) else ""
            td_html = (
                f"<td style='"
                f"background-color: {row_bg} !important; "
                f"{cell_border_l}"
                f"padding: 11px 14px; "
                f"text-align: {align}; "
                f"font-size: 0.88rem; "
                f"font-weight: {cell_weight}; "
                f"color: {cell_color}; "
                f"border-bottom: 1px solid {'#B8E4CD' if is_highlighted else '#F1F5F9'}; "
                f"border-right: 1px solid {'#D1E7DD' if is_highlighted else '#F8FAFC'}; "
                f"font-variant-numeric: tabular-nums; "
                f"white-space: nowrap;'>"
                f"{val}"
                f"</td>"
            )
            td_html_list.append(td_html)

        tr_style = (
            f"background-color: {row_bg}; "
            f"border-left: {row_border_left}; "
            f"{row_shadow} "
            f"transition: background-color 0.15s ease;"
        )
        tr_rows = (
            f"<tr style='{tr_style}' "
            f"onmouseover=\"if(!{str(is_highlighted).lower()}) this.style.backgroundColor='#F4FBF7'\" "
            f"onmouseout=\"if(!{str(is_highlighted).lower()}) this.style.backgroundColor='transparent'\">"
            f"{''.join(td_html_list)}</tr>"
        )
        tr_html_list.append(tr_rows)

    # 🌟 [하단 합계 행] 렌더링 (선택 합계 또는 전체 합계)
    if total_row:
        is_sel_tot = "선택 합계" in str(total_row.get(df.columns[0], ""))
        tot_bg = "#E6FBF1" if is_sel_tot else "#F8FAFB"
        tot_border_top = "2.5px solid #00C56F" if is_sel_tot else "2.5px solid #CBD5E1"
        tot_td_list = []
        for i, col in enumerate(df.columns):
            align = "left" if i == 0 else "right"
            val = total_row.get(col, "-")
            tot_td_html = (
                f"<td style='"
                f"background-color: {tot_bg} !important; "
                f"padding: 13px 14px; "
                f"text-align: {align}; "
                f"font-size: 0.90rem; "
                f"font-weight: 800; "
                f"color: #06291B; "
                f"border-top: {tot_border_top}; "
                f"border-bottom: 2px solid {'#00C56F' if is_sel_tot else '#CBD5E1'}; "
                f"border-right: 1px solid {'#B8E4CD' if is_sel_tot else '#E2E8F0'}; "
                f"font-variant-numeric: tabular-nums; "
                f"white-space: nowrap;'>"
                f"{val}"
                f"</td>"
            )
            tot_td_list.append(tot_td_html)
        tot_tr_html = f"<tr style='background-color: {tot_bg};'>{''.join(tot_td_list)}</tr>"
        tr_html_list.append(tot_tr_html)

    tbody_html = "".join(tr_html_list)

    table_container_html = f"""
    <div style="width: 100%; overflow-x: auto; border: 1px solid #E2E8F0; border-radius: 10px; background-color: #FFFFFF; box-shadow: 0 2px 8px rgba(6, 41, 27, 0.03); margin-top: 6px; margin-bottom: 16px;">
        <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Pretendard Variable', Pretendard, sans-serif;">
            <thead>
                <tr>{th_rows}</tr>
            </thead>
            <tbody>
                {tbody_html}
            </tbody>
        </table>
    </div>
    """
    st.markdown(table_container_html, unsafe_allow_html=True)


def render_total_summary_bar(df_columns: List[str], total_row_dict: Dict[str, Any], is_selected: bool = False) -> None:
    """st.dataframe 바로 아래에 완벽히 정렬되는 이숲 스타일 회계 합계 요약 바를 렌더링합니다."""
    tot_bg = "#E6FBF1" if is_selected else "#F8FAFB"
    tot_border_top = "2.5px solid #00C56F" if is_selected else "2px solid #CBD5E1"

    td_list = []
    for i, col in enumerate(df_columns):
        align = "left" if i == 0 else "right"
        val = total_row_dict.get(col, "-")
        td_html = (
            f"<td style='"
            f"padding: 12px 14px; "
            f"text-align: {align}; "
            f"font-size: 0.90rem; "
            f"font-weight: 800; "
            f"color: #06291B; "
            f"font-variant-numeric: tabular-nums; "
            f"border-right: 1px solid {'#B8E4CD' if is_selected else '#E2E8F0'}; "
            f"white-space: nowrap;'>"
            f"{val}"
            f"</td>"
        )
        td_list.append(td_html)

    summary_html = f"""
    <div style="width: 100%; overflow-x: auto; border: 1px solid {'#00C56F' if is_selected else '#CBD5E1'}; border-radius: 8px; background-color: {tot_bg}; margin-top: 6px; margin-bottom: 16px; box-shadow: 0 2px 6px rgba(0, 197, 111, 0.05);">
        <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Pretendard Variable', Pretendard, sans-serif;">
            <tbody>
                <tr style="background-color: {tot_bg}; border-top: {tot_border_top};">
                    {''.join(td_list)}
                </tr>
            </tbody>
        </table>
    </div>
    """
    st.markdown(summary_html, unsafe_allow_html=True)


def extract_selected_channel(event_obj: Any, key_name: str, valid_channels: List[str]) -> Optional[str]:
    """Plotly 차트 클릭 이벤트(직접 반환값 또는 session_state)로부터 선택된 채널명을 정확히 추출합니다."""
    obj = event_obj
    if not obj and key_name in st.session_state:
        obj = st.session_state[key_name]
    if not obj:
        return None
    sel = obj.get("selection", {}) if isinstance(obj, dict) else getattr(obj, "selection", {})
    pts = sel.get("points", []) if isinstance(sel, dict) else getattr(sel, "points", [])
    if not pts:
        return None
    pt = pts[0]
    if isinstance(pt, dict):
        cd = pt.get("customdata")
        if cd and str(cd).strip() in valid_channels:
            return str(cd).strip()
        lb = pt.get("label")
        if lb and str(lb).strip() in valid_channels:
            return str(lb).strip()
        x_val = pt.get("x")
        if x_val and str(x_val).strip() in valid_channels:
            return str(x_val).strip()
        p_idx = pt.get("point_number", pt.get("pointNumber", pt.get("point_index")))
        if p_idx is not None and 0 <= p_idx < len(valid_channels):
            return valid_channels[p_idx]
    return None


# ======================================================================
# 전역 데이터 로드 및 최상단 헤더 & 컨트롤 바
# ======================================================================
all_dbs: Dict[str, pd.DataFrame] = load_all_dbs()
available_years, available_months = get_available_filters(all_dbs)

# 상단바 접기/펼치기 및 전역 상태 관리
if "show_top_navbar" not in st.session_state:
    st.session_state["show_top_navbar"] = True

# 상단바 펼침/접힘 상태에 맞춰 본문 시작 위치(Padding-Top)를 동적으로 정확히 조절
if st.session_state["show_top_navbar"]:
    st.markdown("""
    <style>
        .block-container {
            padding-top: 185px !important;
        }
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
        .block-container {
            padding-top: 66px !important;
        }
    </style>
    """, unsafe_allow_html=True)

page_options: List[str] = [
    "🏠 통합 대시보드",
    "🟩 네이버 광고",
    "🚀 쿠팡 광고",
    "♾️ 메타 광고",
    "💰 전체 매출",
    "🛠️ 데이터 관리"
]

if "current_page" not in st.session_state or st.session_state["current_page"] not in page_options:
    st.session_state["current_page"] = page_options[0]

if "selected_year_val" not in st.session_state or st.session_state["selected_year_val"] not in available_years:
    st.session_state["selected_year_val"] = available_years[0] if available_years else 2026
if "selected_month_raw_val" not in st.session_state:
    st.session_state["selected_month_raw_val"] = "전체 (연간 누적)"

# 📌 화면 상단 영구 고정 반응형 스마트 네비게이션 래퍼 컨테이너 (border=True로 DOM 독립성 확보)
nav_container = st.container(border=True)
with nav_container:
    # Fixed 고정 CSS 앵커
    st.markdown("<div class='fixed-top-nav-anchor'></div>", unsafe_allow_html=True)

    if st.session_state["show_top_navbar"]:
        # [A] 펼쳐진 풀 상단바 모드
        top_col1, top_col2 = st.columns([3.3, 2.7])

        with top_col1:
            st.markdown("""
            <h1 class='main-dashboard-title' style='font-size: 1.95rem !important; font-weight: 800 !important; color: #06291B !important; margin: 0 !important; line-height: 1.15 !important; letter-spacing: -0.6px !important; display: block !important;'>
                📈 마케팅 통합 성과 대시보드
            </h1>
            <p class='sub-caption' style='margin: 1px 0 0 0 !important; font-size: 0.88rem !important; line-height: 1.2 !important;'>
                마케팅 종합 성과 & 네이버·쿠팡·메타 원본 보고서 자동 연동 시스템
            </p>
            """, unsafe_allow_html=True)

        with top_col2:
            st.markdown("<div class='top-controls-wrapper-anchor'></div>", unsafe_allow_html=True)
            # 우측 상단 컨트롤 바 (연도, 월, 백업, 그리고 '상단바 접기' 버튼)
            f_c1, f_c2, f_c3, f_c4 = st.columns([1.0, 1.3, 1.0, 0.9])
            with f_c1:
                cur_y = st.session_state["selected_year_val"]
                cur_y_idx = available_years.index(cur_y) if cur_y in available_years else 0
                selected_year = st.selectbox("기준 연도", available_years, index=cur_y_idx, key="sb_year_ctl")
                if selected_year != st.session_state["selected_year_val"]:
                    st.session_state["selected_year_val"] = selected_year
                    st.session_state["selected_month_raw_val"] = "전체 (연간 누적)"
                    if "sb_month_ctl" in st.session_state:
                        st.session_state["sb_month_ctl"] = "전체 (연간 누적)"
                    st.rerun()

            with f_c2:
                # 선택된 연도에 속하는 월만 동적으로 추출 (2026 선택 시 2027 월 절대 노출 금지!)
                cur_y_str = str(st.session_state["selected_year_val"])
                months_for_cur_year = [m for m in available_months if str(m).startswith(f"{cur_y_str}-")]
                if not months_for_cur_year:
                    months_for_cur_year = [f"{cur_y_str}-{m:02d}" for m in range(1, 13)]
                month_options = ["전체 (연간 누적)"] + months_for_cur_year

                # 세션에 저장된 월 값이 현재 연도에 맞지 않으면 즉시 '전체 (연간 누적)'으로 자동 보정
                cur_m_val = st.session_state.get("selected_month_raw_val", month_options[0])
                if cur_m_val not in month_options:
                    cur_m_val = month_options[0]
                    st.session_state["selected_month_raw_val"] = cur_m_val
                    if "sb_month_ctl" in st.session_state:
                        st.session_state["sb_month_ctl"] = cur_m_val

                cur_m_idx = month_options.index(cur_m_val)
                selected_month_raw = st.selectbox("기준 월", month_options, index=cur_m_idx, key="sb_month_ctl")
                st.session_state["selected_month_raw_val"] = selected_month_raw
                selected_month: Optional[str] = None if selected_month_raw.startswith("전체") else selected_month_raw

            with f_c3:
                st.markdown("<div class='pc-spacer' style='height: 28px;'></div>", unsafe_allow_html=True)
                excel_backup = export_all_to_excel()
                st.download_button(
                    label="💾 백업",
                    data=excel_backup,
                    file_name="마케팅_통합데이터_백업.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with f_c4:
                st.markdown("<div class='pc-spacer' style='height: 28px;'></div>", unsafe_allow_html=True)
                if st.button("🔼 접기", key="btn_collapse_nav", use_container_width=True, help="상단바를 접어 화면을 더 넓게 봅니다"):
                    st.session_state["show_top_navbar"] = False
                    st.rerun()

        # 수평 메뉴 탭 바
        curr_page_idx = page_options.index(st.session_state["current_page"]) if st.session_state["current_page"] in page_options else 0
        selected_page = st.radio(
            "메뉴 선택",
            page_options,
            index=curr_page_idx,
            horizontal=True,
            label_visibility="collapsed"
        )
        st.session_state["current_page"] = selected_page

        # 각 탭 및 반영 방식 마우스 오버 툴팁(Tooltip) 주입 스크립트
        st.markdown("""
        <script>
        const navTooltips = {
            "통합 대시보드": "🏠 통합 대시보드: 마케팅 종합 성과 & 6대 핵심 KPI 요약",
            "네이버 광고": "💚 네이버 광고: 검색광고 및 성과형 디스플레이(GFA) 상세 실적",
            "쿠팡 광고": "🚀 쿠팡 광고: 로켓배송 매출성장 및 검색광고 집행 실적",
            "메타 광고": "🔵 메타 광고: 페이스북 및 인스타그램 캠페인별 광고 실적",
            "전체 매출": "💰 전체 매출: 스마트스토어, 쿠팡, 자사몰 등 판매채널별 매출",
            "데이터 관리": "🛠️ 데이터 관리: 원본 데이터 조회, 편집 및 백업/기간별 삭제",
            "덮어쓰기": "해당 귀속 연월의 기존 데이터를 교체합니다",
            "누적 추가": "기존 데이터를 보존하고 새 데이터를 누적으로 추가합니다"
        };
        function injectNavTooltips() {
            try {
                const labels = window.parent.document.querySelectorAll('div[data-testid="stRadio"] label');
                labels.forEach(lbl => {
                    const txt = (lbl.innerText || '').trim();
                    for (const [k, v] of Object.entries(navTooltips)) {
                        if (txt.includes(k)) {
                            lbl.setAttribute('title', v);
                        }
                    }
                });
                // 모바일 환경일 때 1행: 통합/매출/관리, 2행: 네이버/쿠팡/메타 순서 재배치 강제 적용
                if (window.innerWidth <= 768) {
                    const group = window.parent.document.querySelector('div[role="radiogroup"][aria-label="메뉴 선택"]') 
                               || window.parent.document.querySelector('div[data-testid="stRadio"] div[role="radiogroup"]');
                    if (group) {
                        Array.from(group.children).forEach(el => {
                            const txt = (el.innerText || '').trim();
                            if (txt.includes('통합')) { el.style.order = '1'; el.style.gridRow = '1'; el.style.gridColumn = '1'; }
                            else if (txt.includes('매출')) { el.style.order = '2'; el.style.gridRow = '1'; el.style.gridColumn = '2'; }
                            else if (txt.includes('관리')) { el.style.order = '3'; el.style.gridRow = '1'; el.style.gridColumn = '3'; }
                            else if (txt.includes('네이버')) { el.style.order = '4'; el.style.gridRow = '2'; el.style.gridColumn = '1'; }
                            else if (txt.includes('쿠팡')) { el.style.order = '5'; el.style.gridRow = '2'; el.style.gridColumn = '2'; }
                            else if (txt.includes('메타')) { el.style.order = '6'; el.style.gridRow = '2'; el.style.gridColumn = '3'; }
                        });
                    }
                }
            } catch(e) {}
        }
        setTimeout(injectNavTooltips, 300);
        setTimeout(injectNavTooltips, 800);
        setTimeout(injectNavTooltips, 1500);
        </script>
        """, unsafe_allow_html=True)

        # 현재 선택된 탭 스마트 안내 캡션 (색상 포인트 & 미니 가이드)
        tab_guides = {
            "🏠 통합 대시보드": ("#00C56F", "전체 마케팅 종합 성과 요약, KPI 카드, 월별 추이 및 채널별 광고·매출 성과를 한눈에 파악합니다."),
            "🟩 네이버 광고": ("#03C75A", "네이버 검색광고 및 성과형 디스플레이(GFA)의 광고비, 노출수, 클릭수, 전환 성과를 상세 분석합니다."),
            "🚀 쿠팡 광고": ("#FF334B", "쿠팡 매출성장광고 및 검색광고 집행 데이터와 광고 기여 매출액을 상세 분석합니다."),
            "♾️ 메타 광고": ("#0084FF", "페이스북 및 인스타그램 스폰서드 캠페인별 광고비, 클릭수, 전환매출 지표를 확인합니다."),
            "💰 전체 매출": ("#F59E0B", "스마트스토어, 쿠팡, 자사몰, 오픈마켓 등 전체 판매채널별 매출액과 주문 실적을 집계합니다."),
            "🛠️ 데이터 관리": ("#7C3AED", "업로드된 원본 DB를 엑셀 표 형태로 직접 수정하거나 특정 기간의 데이터를 안전하게 백업/삭제합니다.")
        }
        t_color, t_text = tab_guides.get(selected_page, ("#00C56F", ""))
        st.markdown(f"""
        <div class="smart-tab-guide-box" style="display: flex; align-items: center; gap: 8px; margin-top: 5px; padding: 5px 12px; background-color: #FFFFFF; border-left: 3.5px solid {t_color}; border-radius: 6px; font-size: 0.82rem; color: #4B6356; box-shadow: 0 1px 3px rgba(6,41,27,0.03);">
            <span style="font-weight: 700; color: {t_color}; white-space: nowrap;">💡 {selected_page}</span>
            <span style="color: #CBD5E1;">|</span>
            <span style="color: #475569;">{t_text}</span>
        </div>
        """, unsafe_allow_html=True)

    else:
        # [B] 접힌 초슬림 플로팅 바 모드 (모바일 100% 최적화: 불필요한 버튼 창 완전 제거, 단 1줄 컴팩트 바)
        cur_y_val = st.session_state['selected_year_val']
        cur_m_val = st.session_state.get('selected_month_raw_val', '')
        if cur_m_val and not cur_m_val.startswith("전체") and cur_m_val.startswith(f"{cur_y_val}-"):
            try:
                p_parts = cur_m_val.split("-")
                badge_period_text = f"{cur_y_val}년 {int(p_parts[1])}월"
            except Exception:
                badge_period_text = f"{cur_y_val}년 {cur_m_val}"
        else:
            badge_period_text = f"{cur_y_val}년 전체 (누적)"

        slim_c1, slim_c2 = st.columns([3.8, 1.4])
        with slim_c1:
            st.markdown(f"""
            <div style="display: flex; align-items: center; gap: 8px; height: 36px; padding: 1px 0;">
                <span style="font-weight: 800; font-size: 1.08rem; color: #06291B; white-space: nowrap;">📈 {st.session_state['current_page']}</span>
                <span style="background-color: #EEF3F0; color: #06291B; padding: 3px 8px; border-radius: 6px; font-size: 0.80rem; font-weight: 600; border: 1px solid #DFE7E2; white-space: nowrap;">
                    📅 {badge_period_text}
                </span>
            </div>
            """, unsafe_allow_html=True)
        with slim_c2:
            if st.button("🔽 상단바 펼치기", key="btn_expand_nav", use_container_width=True, help="상단바와 전체 탭 메뉴를 다시 펼칩니다"):
                st.session_state["show_top_navbar"] = True
                st.rerun()

        selected_page = st.session_state["current_page"]
        selected_year = st.session_state["selected_year_val"]
        selected_month_raw = st.session_state.get("selected_month_raw_val", "전체 (연간 누적)")
        selected_month = selected_month_raw if (selected_month_raw and not selected_month_raw.startswith("전체") and selected_month_raw.startswith(f"{selected_year}-")) else None


# 전역 기준 기간 텍스트 공통 산출 (모든 페이지 공통 참조)
if selected_month and selected_month.startswith(f"{selected_year}-"):
    try:
        p_parts = selected_month.split("-")
        period_title = f"{selected_year}년 {int(p_parts[1])}월"
    except Exception:
        period_title = f"{selected_year}년 {selected_month}"
else:
    period_title = f"{selected_year}년 전체 (누적)"


# ====================================================
# 1. 메인 통합 대시보드 페이지
# ====================================================
if "통합 대시보드" in selected_page:

    # ------------------------------------------------
    # [1] 핵심 KPI 요약 카드 (폰트 2포인트 확대 & 회계 서식 적용)
    # ------------------------------------------------
    if selected_month:
        # 월 선택 시 연도와 월을 함께 표시하되 중복 없이 자연스러운 한국어 표기 (예: 2027년 1월)
        try:
            p_parts = selected_month.split("-")
            p_year = p_parts[0]
            p_month = int(p_parts[1])
            period_title = f"{p_year}년 {p_month}월"
        except Exception:
            period_title = f"{selected_year}년 {selected_month}"
    else:
        period_title = f"{selected_year}년 전체 기간"

    st.markdown(f"<h2>📌 핵심 성과 지표 (KPI) - [{period_title}]</h2>", unsafe_allow_html=True)

    kpis = calculate_summary_metrics(all_dbs, target_year=selected_year, target_month=selected_month)

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        render_google_kpi_card("총 광고비", format_accounting(kpis["총 광고비"], is_currency=True), "네이버+쿠팡+메타 합산", "#FF5E57")
    with c2:
        render_google_kpi_card("총 매출", format_accounting(kpis["총 매출"], is_currency=True), "전체 판매 전환매출", "#00C56F")
    with c3:
        render_google_kpi_card("통합 ROAS", format_accounting(kpis["통합 ROAS"], unit="x"), "총 매출 / 총 광고비", "#06291B")
    with c4:
        render_google_kpi_card("통합 CPA", format_accounting(kpis["통합 CPA"], is_currency=True), "광고비 / 전환수", "#F59E0B")
    with c5:
        render_google_kpi_card("통합 CTR", format_accounting(kpis["통합 CTR"], unit="%"), "클릭수 / 노출수", "#10B981")
    with c6:
        render_google_kpi_card("통합 CVR", format_accounting(kpis["통합 CVR"], unit="%"), "전환수 / 클릭수", "#0284C7")

    # ------------------------------------------------
    # [3] 시각화 차트 (지표별 필터링 pills 및 차트 클릭 지원)
    # ------------------------------------------------
    monthly_df = get_monthly_trend_table(all_dbs, target_year=selected_year)
    channel_df = get_channel_performance_table(all_dbs, target_year=selected_year, target_month=selected_month)

    g1, g2 = st.columns([1.6, 1.0])
    with g1:
        st.markdown("<h3>월별 종합 성과 추이</h3>", unsafe_allow_html=True)
        selected_monthly_metric = st.pills(
            "차트 분석 지표 선택",
            [
                "💰 광고비 vs 매출 & ROAS",
                "👆 클릭수 vs 노출수 & CTR",
                "🎯 광고비 vs 전환수",
                "⚡ 전환수 vs CPA",
                "💳 객단가 vs 전환율 (CVR)"
            ],
            default="💰 광고비 vs 매출 & ROAS",
            key="pill_monthly_metric",
            label_visibility="collapsed"
        )
        chart_monthly_event = st.plotly_chart(
            plot_monthly_revenue_and_cost(monthly_df, metric_mode=selected_monthly_metric or "💰 광고비 vs 매출 & ROAS"),
            use_container_width=True,
            key="chart_monthly_main",
            on_select="rerun"
        )

    with g2:
        st.markdown("<h3>매체별 성과 비중</h3>", unsafe_allow_html=True)
        selected_ad_metric = st.pills(
            "매체 비중 기준 지표",
            ["💸 광고비", "💰 전환매출", "🎯 전환수", "👆 클릭수", "👁️ 노출수"],
            default="💸 광고비",
            key="pill_ad_metric",
            label_visibility="collapsed"
        )
        chart_share_event = st.plotly_chart(
            plot_channel_cost_share(channel_df, metric=selected_ad_metric or "💸 광고비"),
            use_container_width=True,
            key="chart_share_main",
            on_select="rerun"
        )

    # 그래프와 하단 표 사이 여백 확보
    st.markdown("<div style='margin-top: 40px;'></div>", unsafe_allow_html=True)

    # ------------------------------------------------
    # [4] 채널별 광고 성과 표 (행 직접 클릭/체크 & 하단 실시간 합계)
    # ------------------------------------------------
    ad_channel_df = channel_df[channel_df["채널"] != "매출(전체)"].copy()

    st.markdown("<h2>📋 채널별 광고 성과 (기준 월 연동)</h2>", unsafe_allow_html=True)
    st.caption("표에서 원하는 채널 행(체크박스)을 직접 클릭하면 다중 선택 및 하이라이트되며, 하단에 선택 항목의 합계(소계)가 실시간으로 자동 산출됩니다.")

    # 🛠️ 실시간 필터링 컨트롤 바 (표 헤더 클릭으로 직접 정렬 지원)
    ad_f_c1, ad_f_c2 = st.columns([3.2, 0.8])
    with ad_f_c1:
        ad_search_query = st.text_input(
            "채널 검색 필터",
            placeholder="🔍 채널명 검색 (예: 네이버, 쿠팡, 메타)...",
            key="txt_ad_search_filter",
            label_visibility="collapsed"
        )
    with ad_f_c2:
        ad_hide_zero = st.checkbox("0원 채널 제외", value=False, key="chk_ad_hide_zero")

    # [실제 필터링 적용]
    if ad_search_query:
        ad_channel_df = ad_channel_df[ad_channel_df["채널"].astype(str).str.contains(ad_search_query.strip(), case=False, na=False)]
    if ad_hide_zero:
        ad_channel_df = ad_channel_df[(ad_channel_df["광고비"] > 0) | (ad_channel_df["전환매출"] > 0)]

    # [기본 정렬: 광고비 내림차순, 표 제목 클릭 시 오름차순/내림차순 자유 정렬]
    if "광고비" in ad_channel_df.columns:
        ad_channel_df = ad_channel_df.sort_values(by="광고비", ascending=False).reset_index(drop=True)

    ad_channels = [str(ch).strip() for ch in ad_channel_df["채널"]]

    col_cfg_ad = {
        "채널": st.column_config.TextColumn("채널", width="small"),
        "광고비": st.column_config.NumberColumn("광고비", format="₩%,d", width="medium"),
        "노출수": st.column_config.NumberColumn("노출수", format="%,d", width="medium"),
        "클릭수": st.column_config.NumberColumn("클릭수", format="%,d", width="medium"),
        "전환수": st.column_config.NumberColumn("전환수", format="%,d", width="small"),
        "전환매출": st.column_config.NumberColumn("전환매출", format="₩%,d", width="medium"),
        "ROAS": st.column_config.NumberColumn("ROAS", format="%.2fx", width="small"),
        "CPA": st.column_config.NumberColumn("CPA", format="₩%,d", width="small"),
    }

    ad_table_event = st.dataframe(
        ad_channel_df,
        column_config=col_cfg_ad,
        use_container_width=True,
        hide_index=True,
        selection_mode="multi-row",
        on_select="rerun",
        key="df_ad_perf_interactive"
    )

    # st.dataframe 다중 선택 결과 추출
    selected_ad_idx: List[int] = []
    if ad_table_event and hasattr(ad_table_event, "selection"):
        selected_ad_idx = getattr(ad_table_event.selection, "rows", [])
    elif isinstance(ad_table_event, dict):
        selected_ad_idx = ad_table_event.get("selection", {}).get("rows", [])

    if selected_ad_idx:
        target_ad_df = ad_channel_df.iloc[selected_ad_idx]
        sel_ad_names = ", ".join([ad_channels[i] for i in selected_ad_idx if i < len(ad_channels)])
        tot_ad_label = f"🎯 선택 합계 ({sel_ad_names})"
        is_ad_selected = True
    else:
        target_ad_df = ad_channel_df
        tot_ad_label = "🏢 전체 합계"
        is_ad_selected = False

    tot_cost = float(target_ad_df["광고비"].sum()) if not target_ad_df.empty else 0.0
    tot_imp = float(target_ad_df["노출수"].sum()) if not target_ad_df.empty else 0.0
    tot_click = float(target_ad_df["클릭수"].sum()) if not target_ad_df.empty else 0.0
    tot_conv = float(target_ad_df["전환수"].sum()) if not target_ad_df.empty else 0.0
    tot_rev = float(target_ad_df["전환매출"].sum()) if not target_ad_df.empty else 0.0
    tot_roas = (tot_rev / tot_cost) if tot_cost > 0 else 0.0
    tot_cpa = float(round(tot_cost / tot_conv)) if tot_conv > 0 else 0.0

    total_ad_row = {
        "채널": tot_ad_label,
        "광고비": format_accounting(tot_cost, is_currency=True),
        "노출수": format_accounting(tot_imp, decimals=0),
        "클릭수": format_accounting(tot_click, decimals=0),
        "전환수": format_accounting(tot_conv, decimals=0),
        "전환매출": format_accounting(tot_rev, is_currency=True),
        "ROAS": format_accounting(tot_roas, unit="x"),
        "CPA": format_accounting(tot_cpa, is_currency=True)
    }

    # 하단 일체형 합계 요약 바 렌더링
    render_total_summary_bar(list(ad_channel_df.columns), total_ad_row, is_selected=is_ad_selected)

    # ------------------------------------------------
    # [5] 채널별 매출 (기준 월 연동, 판매채널별 그래프 & 표 나란히 배치)
    # ------------------------------------------------
    st.markdown("<div style='margin-top: 35px;'></div>", unsafe_allow_html=True)
    st.markdown("<h2>💰 채널별 매출 (기준 월 연동)</h2>", unsafe_allow_html=True)
    st.caption("선택된 기간 기준 판매처(네이버, 쿠팡, 자사몰, SK베네피아, GS샵 등)별 매출 및 주문 실적 (0원은 '-' 회계 표기)")

    sales_channel_df = get_sales_channel_summary_table(all_dbs, target_year=selected_year, target_month=selected_month)

    sales_col_chart, sales_col_table = st.columns([1.0, 1.45])

    with sales_col_chart:
        st.markdown("<h3>판매채널별 실적 비중</h3>", unsafe_allow_html=True)
        selected_sales_metric = st.pills(
            "판매채널 비중 기준 지표",
            ["💰 매출액", "📦 판매수량", "🚚 배송건수", "💳 배송비", "🏷️ 객단가"],
            default="💰 매출액",
            key="pill_sales_metric",
            label_visibility="collapsed"
        )
        st.plotly_chart(
            plot_sales_channel_share(sales_channel_df, metric=selected_sales_metric or "💰 매출액"),
            use_container_width=True,
            key="chart_sales_channel_share"
        )

    with sales_col_table:
        st.markdown("<h3>판매채널별 상세 실적</h3>", unsafe_allow_html=True)
        st.caption("표의 행을 직접 클릭하여 원하는 판매처를 다중 선택하시면 하단에 실시간 소계가 반영됩니다.")

        # 🛠️ 판매채널별 실시간 필터링 컨트롤 바 (표 헤더 클릭으로 직접 정렬 지원)
        s_f_c1, s_f_c2 = st.columns([3.0, 1.0])
        with s_f_c1:
            sales_search_query = st.text_input(
                "판매처 검색",
                placeholder="🔍 판매처 검색 (예: 자사몰, 쿠팡, GS샵)...",
                key="txt_sales_search_filter",
                label_visibility="collapsed"
            )
        with s_f_c2:
            sales_hide_zero = st.checkbox("0원 판매처 제외", value=False, key="chk_sales_hide_zero")

        # [실제 필터링 적용]
        if sales_search_query:
            sales_channel_df = sales_channel_df[sales_channel_df["판매 채널"].astype(str).str.contains(sales_search_query.strip(), case=False, na=False)]
        if sales_hide_zero and "매출액" in sales_channel_df.columns:
            sales_channel_df = sales_channel_df[sales_channel_df["매출액"] > 0]

        # [기본 정렬: 매출액 내림차순, 표 제목 클릭 시 오름차순/내림차순 자유 정렬]
        if "매출액" in sales_channel_df.columns:
            sales_channel_df = sales_channel_df.sort_values(by="매출액", ascending=False).reset_index(drop=True)

        sales_channels = [str(ch).strip() for ch in sales_channel_df["판매 채널"] if str(ch).strip() and str(ch).strip() != "nan"]

        col_cfg_sales = {
            "판매 채널": st.column_config.TextColumn("판매 채널", width="small"),
            "매출액": st.column_config.NumberColumn("매출액", format="₩%,d", width="medium"),
            "판매수량": st.column_config.NumberColumn("판매수량", format="%,d개", width="small"),
            "배송건수": st.column_config.NumberColumn("배송건수", format="%,d건", width="small"),
            "배송비": st.column_config.NumberColumn("배송비", format="₩%,d", width="small"),
            "객단가": st.column_config.NumberColumn("객단가", format="₩%,d", width="small"),
        }

        sales_table_event = st.dataframe(
            sales_channel_df,
            column_config=col_cfg_sales,
            use_container_width=True,
            hide_index=True,
            selection_mode="multi-row",
            on_select="rerun",
            key="df_sales_perf_interactive"
        )

        selected_sales_idx: List[int] = []
        if sales_table_event and hasattr(sales_table_event, "selection"):
            selected_sales_idx = getattr(sales_table_event.selection, "rows", [])
        elif isinstance(sales_table_event, dict):
            selected_sales_idx = sales_table_event.get("selection", {}).get("rows", [])

        if selected_sales_idx:
            target_sales_df = sales_channel_df.iloc[selected_sales_idx]
            sel_sales_names = ", ".join([sales_channels[i] for i in selected_sales_idx if i < len(sales_channels)])
            tot_sales_label = f"🎯 선택 합계 ({sel_sales_names})"
            is_sales_selected = True
        else:
            target_sales_df = sales_channel_df
            tot_sales_label = "🏬 전체 합계"
            is_sales_selected = False

        tot_rev_s = float(target_sales_df["매출액"].sum()) if not target_sales_df.empty else 0.0
        tot_qty_s = float(target_sales_df["판매수량"].sum()) if not target_sales_df.empty else 0.0
        tot_deliv_s = float(target_sales_df["배송건수"].sum()) if not target_sales_df.empty else 0.0
        tot_fee_s = float(target_sales_df["배송비"].sum()) if not target_sales_df.empty else 0.0
        tot_aov_s = float(round(tot_rev_s / tot_qty_s)) if tot_qty_s > 0 else 0.0

        total_sales_row = {
            "판매 채널": tot_sales_label,
            "매출액": format_accounting(tot_rev_s, is_currency=True),
            "판매수량": format_accounting(tot_qty_s, unit="개"),
            "배송건수": format_accounting(tot_deliv_s, unit="건"),
            "배송비": format_accounting(tot_fee_s, is_currency=True),
            "객단가": format_accounting(tot_aov_s, is_currency=True)
        }

        render_total_summary_bar(list(sales_channel_df.columns), total_sales_row, is_selected=is_sales_selected)

    # ------------------------------------------------
    # [6] 월별 추이 테이블 & 연도별 요약 표 (회계 서식 및 조건부 서식 적용)
    # ------------------------------------------------
    tab1, tab2 = st.tabs(["📅 월별 추이 성과표", "📊 연도별 성과 요약"])

    # 전체 데이터가 존재하는 모든 월을 취합한 전체 누적 테이블 생성 (target_year=None)
    all_monthly_df = get_monthly_trend_table(all_dbs, target_year=None)

    col_cfg_monthly = {
        "월": st.column_config.TextColumn("월", width="small"),
        "총 광고비": st.column_config.NumberColumn("총 광고비", format="₩%,d", width="medium"),
        "총 매출": st.column_config.NumberColumn("총 매출", format="₩%,d", width="medium"),
        "ROAS": st.column_config.NumberColumn("ROAS", format="%.2fx", width="small"),
        "CPA": st.column_config.NumberColumn("CPA", format="₩%,d", width="small"),
        "객단가": st.column_config.NumberColumn("객단가", format="₩%,d", width="small"),
        "CPC": st.column_config.NumberColumn("CPC", format="₩%,d", width="small"),
        "CTR": st.column_config.NumberColumn("CTR", format="%.2f%%", width="small"),
        "CVR": st.column_config.NumberColumn("CVR", format="%.2f%%", width="small"),
        "노출수": st.column_config.NumberColumn("노출수", format="%,d", width="medium"),
        "클릭수": st.column_config.NumberColumn("클릭수", format="%,d", width="small"),
        "전환수": st.column_config.NumberColumn("전환수", format="%,d", width="small"),
    }

    with tab1:
        # 월별 성과표 상단: 실제 필터링 컨트롤 바 (표 헤더 클릭으로 직접 정렬 지원)
        m_head_c1, m_head_c2, m_head_c3 = st.columns([1.8, 1.4, 0.8])
        with m_head_c1:
            m_search_query = st.text_input(
                "월 검색 필터",
                placeholder="🔍 연월 검색 (예: 2027, 2026-10)...",
                key="txt_m_search_filter",
                label_visibility="collapsed"
            )
        with m_head_c2:
            year_filter_options = ["전체 (누적 기간)"] + [f"{y}년" for y in available_years]
            sel_m_year_filter = st.selectbox(
                "조회 기간 필터",
                year_filter_options,
                index=0,
                key="sb_m_table_filter",
                label_visibility="collapsed"
            )
        with m_head_c3:
            m_hide_empty = st.checkbox("0원 실적 제외", value=False, key="chk_m_hide_empty")

        # [실제 연도 및 검색 필터링 적용]
        if sel_m_year_filter.startswith("전체"):
            disp_monthly = all_monthly_df.copy()
        else:
            filter_y_str = sel_m_year_filter.replace("년", "").strip()
            disp_monthly = all_monthly_df[all_monthly_df["월"].astype(str).str.startswith(filter_y_str)].copy()

        if m_search_query:
            disp_monthly = disp_monthly[disp_monthly["월"].astype(str).str.contains(m_search_query.strip(), case=False, na=False)]

        if m_hide_empty:
            disp_monthly = disp_monthly[(disp_monthly["총 광고비"] > 0) | (disp_monthly["총 매출"] > 0)]

        # [기본 정렬: 최신 월 순, 표 제목 클릭 시 오름차순/내림차순 자유 정렬]
        disp_monthly = disp_monthly.sort_values(by="월", ascending=False).reset_index(drop=True)

        monthly_table_event = st.dataframe(
            disp_monthly,
            column_config=col_cfg_monthly,
            use_container_width=True,
            hide_index=True,
            selection_mode="multi-row",
            on_select="rerun",
            key="df_monthly_trend_interactive"
        )

        # 다중 행 선택 결과 추출
        selected_m_idx: List[int] = []
        if monthly_table_event and hasattr(monthly_table_event, "selection"):
            selected_m_idx = getattr(monthly_table_event.selection, "rows", [])
        elif isinstance(monthly_table_event, dict):
            selected_m_idx = monthly_table_event.get("selection", {}).get("rows", [])

        m_months_list = [str(m).strip() for m in disp_monthly["월"]]
        if selected_m_idx:
            target_m_df = disp_monthly.iloc[selected_m_idx]
            sel_m_names = ", ".join([m_months_list[i] for i in selected_m_idx if i < len(m_months_list)])
            tot_m_label = f"🎯 선택 합계 ({sel_m_names})"
            is_m_selected = True
        else:
            target_m_df = disp_monthly
            tot_m_label = "📅 전체 누적 합계" if sel_m_year_filter.startswith("전체") else f"📅 {sel_m_year_filter} 합계"
            is_m_selected = False

        tot_m_cost = float(target_m_df["총 광고비"].sum()) if not target_m_df.empty else 0.0
        tot_m_rev = float(target_m_df["총 매출"].sum()) if not target_m_df.empty else 0.0
        tot_m_imp = float(target_m_df["노출수"].sum()) if not target_m_df.empty else 0.0
        tot_m_click = float(target_m_df["클릭수"].sum()) if not target_m_df.empty else 0.0
        tot_m_conv = float(target_m_df["전환수"].sum()) if not target_m_df.empty else 0.0

        tot_m_roas = (tot_m_rev / tot_m_cost) if tot_m_cost > 0 else 0.0
        tot_m_cpa = float(round(tot_m_cost / tot_m_conv)) if tot_m_conv > 0 else 0.0
        tot_m_cpc = float(round(tot_m_cost / tot_m_click)) if tot_m_click > 0 else 0.0
        tot_m_ctr = (tot_m_click / tot_m_imp * 100) if tot_m_imp > 0 else 0.0
        tot_m_cvr = (tot_m_conv / tot_m_click * 100) if tot_m_click > 0 else 0.0
        tot_m_aov = float(round(target_m_df["객단가"].mean())) if not target_m_df.empty and "객단가" in target_m_df.columns else 0.0

        total_monthly_row = {
            "월": tot_m_label,
            "총 광고비": format_accounting(tot_m_cost, is_currency=True),
            "총 매출": format_accounting(tot_m_rev, is_currency=True),
            "ROAS": format_accounting(tot_m_roas, unit="x"),
            "CPA": format_accounting(tot_m_cpa, is_currency=True),
            "객단가": format_accounting(tot_m_aov, is_currency=True),
            "CPC": format_accounting(tot_m_cpc, is_currency=True),
            "CTR": format_accounting(tot_m_ctr, unit="%"),
            "CVR": format_accounting(tot_m_cvr, unit="%"),
            "노출수": format_accounting(tot_m_imp, decimals=0),
            "클릭수": format_accounting(tot_m_click, decimals=0),
            "전환수": format_accounting(tot_m_conv, decimals=0),
        }

        # 하단 일체형 합계 요약 바 렌더링
        render_total_summary_bar(list(disp_monthly.columns), total_monthly_row, is_selected=is_m_selected)

    with tab2:
        st.caption("표에서 원하는 연도(행)를 직접 클릭하면 다중 선택 및 체크(`✔`)되며, 하단에 선택 항목의 합계(소계)가 실시간으로 자동 산출됩니다.")
        yearly_df = get_yearly_trend_table(all_dbs)

        # 🛠️ 연도별 실시간 필터링 컨트롤 바 (표 헤더 클릭으로 직접 정렬 지원)
        y_f_c1, y_f_c2 = st.columns([3.0, 1.0])
        with y_f_c1:
            y_search_query = st.text_input(
                "연도 검색",
                placeholder="🔍 연도 검색 (예: 2026, 2027)...",
                key="txt_yearly_search_filter",
                label_visibility="collapsed"
            )
        with y_f_c2:
            y_hide_zero = st.checkbox("0원 연도 제외", value=False, key="chk_yearly_hide_zero")

        # [실제 필터링 적용]
        if y_search_query:
            yearly_df = yearly_df[yearly_df["연도"].astype(str).str.contains(y_search_query.strip(), case=False, na=False)]
        if y_hide_zero:
            yearly_df = yearly_df[(yearly_df["총 광고비"] > 0) | (yearly_df["총 매출"] > 0)]

        # [기본 정렬: 최신 연도 순, 표 제목 클릭 시 오름차순/내림차순 자유 정렬]
        yearly_df = yearly_df.sort_values(by="연도", ascending=False).reset_index(drop=True)

        col_cfg_yearly = {
            "연도": st.column_config.TextColumn("연도", width="small"),
            "총 광고비": st.column_config.NumberColumn("총 광고비", format="₩%,d", width="medium"),
            "총 매출": st.column_config.NumberColumn("총 매출", format="₩%,d", width="medium"),
            "ROAS": st.column_config.NumberColumn("ROAS", format="%.2fx", width="small"),
            "CPA": st.column_config.NumberColumn("CPA", format="₩%,d", width="small"),
            "객단가": st.column_config.NumberColumn("객단가", format="₩%,d", width="small"),
            "CPC": st.column_config.NumberColumn("CPC", format="₩%,d", width="small"),
            "CTR": st.column_config.NumberColumn("CTR", format="%.2f%%", width="small"),
            "CVR": st.column_config.NumberColumn("CVR", format="%.2f%%", width="small"),
            "노출수": st.column_config.NumberColumn("노출수", format="%,d", width="medium"),
            "클릭수": st.column_config.NumberColumn("클릭수", format="%,d", width="small"),
            "전환수": st.column_config.NumberColumn("전환수", format="%,d", width="small"),
        }

        yearly_table_event = st.dataframe(
            yearly_df,
            column_config=col_cfg_yearly,
            use_container_width=True,
            hide_index=True,
            selection_mode="multi-row",
            on_select="rerun",
            key="df_yearly_trend_interactive"
        )

        selected_y_idx: List[int] = []
        if yearly_table_event and hasattr(yearly_table_event, "selection"):
            selected_y_idx = getattr(yearly_table_event.selection, "rows", [])
        elif isinstance(yearly_table_event, dict):
            selected_y_idx = yearly_table_event.get("selection", {}).get("rows", [])

        y_years_list = [str(y).strip() for y in yearly_df["연도"]]
        if selected_y_idx:
            target_y_df = yearly_df.iloc[selected_y_idx]
            sel_y_names = ", ".join([y_years_list[i] for i in selected_y_idx if i < len(y_years_list)])
            tot_y_label = f"🎯 선택 합계 ({sel_y_names})"
            is_y_selected = True
        else:
            target_y_df = yearly_df
            tot_y_label = "🏢 전체 누적 합계"
            is_y_selected = False

        tot_y_cost = float(target_y_df["총 광고비"].sum()) if not target_y_df.empty else 0.0
        tot_y_rev = float(target_y_df["총 매출"].sum()) if not target_y_df.empty else 0.0
        tot_y_imp = float(target_y_df["노출수"].sum()) if not target_y_df.empty else 0.0
        tot_y_click = float(target_y_df["클릭수"].sum()) if not target_y_df.empty else 0.0
        tot_y_conv = float(target_y_df["전환수"].sum()) if not target_y_df.empty else 0.0

        tot_y_roas = (tot_y_rev / tot_y_cost) if tot_y_cost > 0 else 0.0
        tot_y_cpa = float(round(tot_y_cost / tot_y_conv)) if tot_y_conv > 0 else 0.0
        tot_y_cpc = float(round(tot_y_cost / tot_y_click)) if tot_y_click > 0 else 0.0
        tot_y_ctr = (tot_y_click / tot_y_imp * 100) if tot_y_imp > 0 else 0.0
        tot_y_cvr = (tot_y_conv / tot_y_click * 100) if tot_y_click > 0 else 0.0
        tot_y_aov = float(round(target_y_df["객단가"].mean())) if not target_y_df.empty and "객단가" in target_y_df.columns else 0.0

        total_yearly_row = {
            "연도": tot_y_label,
            "총 광고비": format_accounting(tot_y_cost, is_currency=True),
            "총 매출": format_accounting(tot_y_rev, is_currency=True),
            "ROAS": format_accounting(tot_y_roas, unit="x"),
            "CPA": format_accounting(tot_y_cpa, is_currency=True),
            "객단가": format_accounting(tot_y_aov, is_currency=True),
            "CPC": format_accounting(tot_y_cpc, is_currency=True),
            "CTR": format_accounting(tot_y_ctr, unit="%"),
            "CVR": format_accounting(tot_y_cvr, unit="%"),
            "노출수": format_accounting(tot_y_imp, decimals=0),
            "클릭수": format_accounting(tot_y_click, decimals=0),
            "전환수": format_accounting(tot_y_conv, decimals=0),
        }

        render_total_summary_bar(list(yearly_df.columns), total_yearly_row, is_selected=is_y_selected)

    # 성과 지표 안내 박스 (이숲 소프트 민트 스타일 & 객단가/CPC/CTR/CVR 상세 안내)
    st.markdown("""
    <div class="g-guide-box">
        <b style="font-size: 0.98rem; color: #06291B; font-weight: 700;">💡 성과 지표 안내</b><br>
        <div style="margin-top: 8px; line-height: 1.68;">
            • <b>회계 표시 기준</b>: 모든 금액과 수량은 천 단위 콤마(,)가 적용되며, 실적이 없는 0원/0건은 '-'로 단정하게 표기됩니다.<br>
            • <b>ROAS (광고수익률)</b>: 산출 공식 = <code>총 매출 / 총 광고비</code> | 광고비 1원 투입 대비 발생한 매출 배수 (예: 3.50x = 광고비의 350% 매출 달성)<br>
            • <b>CPA (전환당비용)</b>: 산출 공식 = <code>총 광고비 / 총 전환수</code> | 고객 1건 전환(구매 등)을 유치하는 데 투입된 평균 광고비용 (원화 정수 표기)<br>
            • <b>객단가 (평균 주문금액, AOV)</b>: 산출 공식 = <code>총 매출 / 총 판매수량</code> | 주문 1건(수량 1개)당 결제된 평균 결제 금액<br>
            • <b>CPC (클릭당비용)</b>: 산출 공식 = <code>총 광고비 / 총 클릭수</code> | 잠재 고객 1회 클릭 유입당 지출된 평균 광고비<br>
            • <b>CTR (클릭률)</b>: 산출 공식 = <code>(총 클릭수 / 총 노출수) × 100%</code> | 광고 노출 횟수 중 실제 고객 클릭으로 이어진 유입 비율 (광고 소재 매력도 및 주목도 지표)<br>
            • <b>CVR (전환율)</b>: 산출 공식 = <code>(총 전환수 / 총 클릭수) × 100%</code> | 광고 클릭 유입 고객 중 최종 구매/전환으로 이어진 비율 (상세페이지 및 상품 전환 경쟁력 지표)
        </div>
    </div>
    """, unsafe_allow_html=True)


# ====================================================
# 공통 함수: 개별 상세 페이지 렌더러
# ====================================================
def render_channel_page(
    db_key: str,
    channel_display_name: str,
    theme_color: str,
    cat_column: Optional[str] = None
) -> None:
    """개별 상세 페이지(상단 지표/차트 + 하단 회계 서식 테이블 및 '변경사항 저장')를 렌더링합니다."""
    st.markdown(f"<h2>{channel_display_name}</h2>", unsafe_allow_html=True)

    df_raw = all_dbs.get(db_key, pd.DataFrame())
    fdf = filter_dataframe(df_raw, target_year=selected_year, target_month=selected_month)

    # 1. 상단 핵심 지표 요약 (회계 서식 적용)
    cost = fdf["광고비"].sum() if "광고비" in fdf.columns else 0
    rev = fdf["전환매출"].sum() if "전환매출" in fdf.columns else 0
    roas = (rev / cost) if cost > 0 else 0
    clicks = fdf["클릭수"].sum() if "클릭수" in fdf.columns else 0
    convs = fdf["전환수"].sum() if "전환수" in fdf.columns else 0
    cpa = float(round(cost / convs)) if convs > 0 else 0.0

    if db_key == "DB_매출":
        qty = fdf["판매수량"].sum() if "판매수량" in fdf.columns else 0
        deliv = fdf["배송건수"].sum() if "배송건수" in fdf.columns else 0
        deliv_fee = fdf["배송비"].sum() if "배송비" in fdf.columns else 0
        aov = float(round(rev / qty)) if qty > 0 else 0.0

        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            render_google_kpi_card("총 매출액", format_accounting(rev, is_currency=True), "전체 판매액", theme_color)
        with m2:
            render_google_kpi_card("총 판매수량", format_accounting(qty, unit="개"), "출고 상품 수량", theme_color)
        with m3:
            render_google_kpi_card("총 배송건수", format_accounting(deliv, unit="건"), "배송 접수 건수", theme_color)
        with m4:
            render_google_kpi_card("총 배송비", format_accounting(deliv_fee, is_currency=True), "배송 비용", theme_color)
        with m5:
            render_google_kpi_card("평균 객단가", format_accounting(aov, is_currency=True), "매출 / 수량", theme_color)
    else:
        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            render_google_kpi_card("광고비", format_accounting(cost, is_currency=True), "집행 비용", theme_color)
        with m2:
            render_google_kpi_card("전환매출", format_accounting(rev, is_currency=True), "광고 기여 매출", theme_color)
        with m3:
            render_google_kpi_card("ROAS", format_accounting(roas, unit="x"), "매출 / 광고비", theme_color)
        with m4:
            render_google_kpi_card("유입수 (클릭)", format_accounting(clicks, unit="회"), "총 클릭수", theme_color)
        with m5:
            render_google_kpi_card("CPA", format_accounting(cpa, is_currency=True), "전환당 비용", theme_color)

    # 2. 상단 차트 (외부 헤더 분리로 범례 겹침 원천 차단)
    clean_name = channel_display_name.replace("💚", "").replace("🚀", "").replace("🔵", "").replace("💰", "").strip()
    gc1, gc2 = st.columns([1, 1.2])
    with gc1:
        if cat_column and cat_column in fdf.columns:
            st.markdown(f"<h3>{clean_name} 분류별 비중</h3>", unsafe_allow_html=True)
            st.plotly_chart(
                plot_pie_category(fdf, category_col=cat_column, value_col="전환매출" if db_key == "DB_매출" else "광고비"),
                use_container_width=True,
                key=f"chart_pie_{db_key}"
            )
        else:
            m_grp = df_raw.groupby("월").agg({"광고비": "sum", "전환매출": "sum"}).reset_index()
            m_grp["ROAS"] = m_grp["전환매출"] / m_grp["광고비"].replace(0, 1)
            st.markdown(f"<h3>{clean_name} 월별 ROAS 추이</h3>", unsafe_allow_html=True)
            st.plotly_chart(
                plot_channel_trend_line(m_grp, metric_cols=["ROAS"]),
                use_container_width=True,
                key=f"chart_roas_{db_key}"
            )
    with gc2:
        st.markdown(f"<h3>{clean_name} 월별 실적 추이</h3>", unsafe_allow_html=True)
        st.plotly_chart(
            plot_channel_trend_line(df_raw, metric_cols=["전환매출", "광고비"] if db_key != "DB_매출" else ["전환매출", "판매수량"]),
            use_container_width=True,
            key=f"chart_trend_{db_key}"
        )

    # 3. 하단 전체 상세 데이터 테이블
    st.markdown("<div style='margin-top: 45px;'></div>", unsafe_allow_html=True)
    st.markdown("<h2>📑 전체 상세 데이터 내역</h2>", unsafe_allow_html=True)

    tab_ch_view, tab_ch_edit = st.tabs([
        "📊 성과 분석 표 (검색·정렬·선택합계)",
        "✏️ 원본 데이터 직접 수정 (에디터)"
    ])

    with tab_ch_view:
        st.caption("표에서 원하는 행(체크박스)을 직접 클릭하면 다중 선택 및 체크(`✔`)되며, 하단에 선택 항목의 합계(소계)가 실시간으로 자동 산출됩니다.")

        # 🛠️ 실시간 필터링 컨트롤 바 (표 헤더 클릭으로 직접 정렬 지원)
        ch_c1, ch_c2, ch_c3 = st.columns([1.8, 1.4, 0.8])
        with ch_c1:
            search_ph = "🔍 판매처, 분류, 연월 검색..." if db_key == "DB_매출" else "🔍 캠페인명, 분류, 연월 검색..."
            ch_search_query = st.text_input(
                "검색 필터",
                placeholder=search_ph,
                key=f"txt_search_{db_key}",
                label_visibility="collapsed"
            )
        with ch_c2:
            safe_p_title = period_title if "period_title" in globals() or "period_title" in locals() else f"{selected_year}년"
            period_options = [f"📅 현재 기준 ({safe_p_title})", "🌐 전체 기간 (누적)"] + [f"{y}년" for y in available_years]
            ch_period_filter = st.selectbox(
                "조회 기간",
                period_options,
                index=0,
                key=f"sb_period_{db_key}",
                label_visibility="collapsed"
            )
        with ch_c3:
            ch_hide_zero = st.checkbox("0원 실적 제외", value=False, key=f"chk_zero_{db_key}")

        # 1) 기본 데이터프레임 복사 및 파생 지표 산출
        df_disp = df_raw.copy()
        if db_key == "DB_매출":
            qty_col = df_disp["판매수량"] if "판매수량" in df_disp.columns else pd.Series(0, index=df_disp.index)
            rev_col = df_disp["전환매출"] if "전환매출" in df_disp.columns else pd.Series(0, index=df_disp.index)
            df_disp["객단가"] = np.where(qty_col > 0, (rev_col / qty_col).round(), 0.0)
        else:
            cost_col = df_disp["광고비"] if "광고비" in df_disp.columns else pd.Series(0, index=df_disp.index)
            rev_col = df_disp["전환매출"] if "전환매출" in df_disp.columns else pd.Series(0, index=df_disp.index)
            conv_col = df_disp["전환수"] if "전환수" in df_disp.columns else pd.Series(0, index=df_disp.index)
            click_col = df_disp["클릭수"] if "클릭수" in df_disp.columns else pd.Series(0, index=df_disp.index)
            imp_col = df_disp["노출수"] if "노출수" in df_disp.columns else pd.Series(0, index=df_disp.index)

            df_disp["ROAS"] = np.where(cost_col > 0, (rev_col / cost_col), 0.0)
            df_disp["CPA"] = np.where(conv_col > 0, (cost_col / conv_col).round(), 0.0)
            df_disp["CPC"] = np.where(click_col > 0, (cost_col / click_col).round(), 0.0)
            df_disp["CTR"] = np.where(imp_col > 0, (click_col / imp_col * 100), 0.0)
            df_disp["CVR"] = np.where(click_col > 0, (conv_col / click_col * 100), 0.0)

        # 2) 기간 필터링 적용
        if ch_period_filter.startswith("📅 현재 기준"):
            df_disp = filter_dataframe(df_disp, target_year=selected_year, target_month=selected_month)
        elif ch_period_filter.endswith("년"):
            target_filter_y = int(ch_period_filter.replace("년", "").strip())
            df_disp = df_disp[df_disp["년도"] == target_filter_y]

        # 3) 검색 필터링 적용 (진짜 행 단위 필터링)
        if ch_search_query:
            q_clean = ch_search_query.strip()
            match_mask = pd.Series(False, index=df_disp.index)
            searchable_cols = ["캠페인 이름", "마케팅 분류", "매출분류", "월", "년도"]
            for s_col in searchable_cols:
                if s_col in df_disp.columns:
                    match_mask = match_mask | df_disp[s_col].astype(str).str.contains(q_clean, case=False, na=False)
            df_disp = df_disp[match_mask]

        # 4) 0원 실적 제외 적용
        if ch_hide_zero:
            if db_key == "DB_매출":
                df_disp = df_disp[df_disp["전환매출"] > 0]
            else:
                df_disp = df_disp[(df_disp["광고비"] > 0) | (df_disp["전환매출"] > 0)]

        # 5) 기본 정렬 적용 (표 제목 클릭 시 오름차순/내림차순 자유 정렬)
        if db_key == "DB_매출" and "전환매출" in df_disp.columns:
            df_disp = df_disp.sort_values(by="전환매출", ascending=False).reset_index(drop=True)
        elif "광고비" in df_disp.columns:
            df_disp = df_disp.sort_values(by="광고비", ascending=False).reset_index(drop=True)

        # 6) 컬럼 표시 순서 정리
        if db_key == "DB_네이버":
            ordered_cols = ["월", "마케팅 분류", "캠페인 이름", "광고비", "전환매출", "ROAS", "CPA", "CPC", "CTR", "CVR", "노출수", "클릭수", "전환수"]
        elif db_key in ["DB_쿠팡", "DB_메타"]:
            ordered_cols = ["월", "캠페인 이름", "광고비", "전환매출", "ROAS", "CPA", "CPC", "CTR", "CVR", "노출수", "클릭수", "전환수"]
        elif db_key == "DB_매출":
            ordered_cols = ["월", "매출분류", "전환매출", "판매수량", "객단가", "배송건수", "배송비"]
        else:
            ordered_cols = list(df_disp.columns)

        final_show_cols = [c for c in ordered_cols if c in df_disp.columns]
        df_disp = df_disp[final_show_cols]

        # 7) 컬럼 설정 (회계 서식 및 너비)
        col_cfg_detail = {
            "월": st.column_config.TextColumn("월", width="small"),
            "마케팅 분류": st.column_config.TextColumn("분류", width="small"),
            "매출분류": st.column_config.TextColumn("판매 채널", width="medium"),
            "캠페인 이름": st.column_config.TextColumn("캠페인명", width="medium"),
            "광고비": st.column_config.NumberColumn("광고비", format="₩%,d"),
            "전환매출": st.column_config.NumberColumn("매출액" if db_key == "DB_매출" else "전환매출", format="₩%,d"),
            "ROAS": st.column_config.NumberColumn("ROAS", format="%.2fx"),
            "CPA": st.column_config.NumberColumn("CPA", format="₩%,d"),
            "CPC": st.column_config.NumberColumn("CPC", format="₩%,d"),
            "객단가": st.column_config.NumberColumn("객단가", format="₩%,d"),
            "CTR": st.column_config.NumberColumn("CTR", format="%.2f%%"),
            "CVR": st.column_config.NumberColumn("CVR", format="%.2f%%"),
            "노출수": st.column_config.NumberColumn("노출수", format="%,d"),
            "클릭수": st.column_config.NumberColumn("클릭수", format="%,d"),
            "전환수": st.column_config.NumberColumn("전환수", format="%,d"),
            "판매수량": st.column_config.NumberColumn("판매수량", format="%,d개"),
            "배송건수": st.column_config.NumberColumn("배송건수", format="%,d건"),
            "배송비": st.column_config.NumberColumn("배송비", format="₩%,d"),
        }

        # 8) 인터랙티브 데이터프레임 렌더링
        ch_table_event = st.dataframe(
            df_disp,
            column_config=col_cfg_detail,
            use_container_width=True,
            hide_index=True,
            selection_mode="multi-row",
            on_select="rerun",
            key=f"df_interactive_{db_key}"
        )

        # 9) 행 선택 및 실시간 합계 / 소계 바 연동
        selected_ch_idx: List[int] = []
        if ch_table_event and hasattr(ch_table_event, "selection"):
            selected_ch_idx = getattr(ch_table_event.selection, "rows", [])
        elif isinstance(ch_table_event, dict):
            selected_ch_idx = ch_table_event.get("selection", {}).get("rows", [])

        if selected_ch_idx:
            target_ch_tot = df_disp.iloc[selected_ch_idx]
            tot_label = f"🎯 선택 합계 ({len(selected_ch_idx)}건 선택)"
            is_ch_sel = True
        else:
            target_ch_tot = df_disp
            tot_label = "🏢 전체 누적 합계" if "전체" in ch_period_filter else f"📅 {ch_period_filter} 합계"
            is_ch_sel = False

        # 합계 계산
        tot_c_cost = float(target_ch_tot["광고비"].sum()) if not target_ch_tot.empty and "광고비" in target_ch_tot.columns else 0.0
        tot_c_rev = float(target_ch_tot["전환매출"].sum()) if not target_ch_tot.empty and "전환매출" in target_ch_tot.columns else 0.0
        tot_c_imp = float(target_ch_tot["노출수"].sum()) if not target_ch_tot.empty and "노출수" in target_ch_tot.columns else 0.0
        tot_c_click = float(target_ch_tot["클릭수"].sum()) if not target_ch_tot.empty and "클릭수" in target_ch_tot.columns else 0.0
        tot_c_conv = float(target_ch_tot["전환수"].sum()) if not target_ch_tot.empty and "전환수" in target_ch_tot.columns else 0.0
        tot_c_qty = float(target_ch_tot["판매수량"].sum()) if not target_ch_tot.empty and "판매수량" in target_ch_tot.columns else 0.0
        tot_c_deliv = float(target_ch_tot["배송건수"].sum()) if not target_ch_tot.empty and "배송건수" in target_ch_tot.columns else 0.0
        tot_c_fee = float(target_ch_tot["배송비"].sum()) if not target_ch_tot.empty and "배송비" in target_ch_tot.columns else 0.0

        tot_c_roas = (tot_c_rev / tot_c_cost) if tot_c_cost > 0 else 0.0
        tot_c_cpa = float(round(tot_c_cost / tot_c_conv)) if tot_c_conv > 0 else 0.0
        tot_c_cpc = float(round(tot_c_cost / tot_c_click)) if tot_c_click > 0 else 0.0
        tot_c_ctr = (tot_c_click / tot_c_imp * 100) if tot_c_imp > 0 else 0.0
        tot_c_cvr = (tot_c_conv / tot_c_click * 100) if tot_c_click > 0 else 0.0
        tot_c_aov = float(round(tot_c_rev / tot_c_qty)) if tot_c_qty > 0 else 0.0

        total_ch_row = {
            "월": tot_label,
            "마케팅 분류": "-",
            "매출분류": "-",
            "캠페인 이름": "-",
            "광고비": format_accounting(tot_c_cost, is_currency=True),
            "전환매출": format_accounting(tot_c_rev, is_currency=True),
            "ROAS": format_accounting(tot_c_roas, unit="x"),
            "CPA": format_accounting(tot_c_cpa, is_currency=True),
            "CPC": format_accounting(tot_c_cpc, is_currency=True),
            "CTR": format_accounting(tot_c_ctr, unit="%"),
            "CVR": format_accounting(tot_c_cvr, unit="%"),
            "노출수": format_accounting(tot_c_imp, decimals=0),
            "클릭수": format_accounting(tot_c_click, decimals=0),
            "전환수": format_accounting(tot_c_conv, decimals=0),
            "판매수량": format_accounting(tot_c_qty, decimals=0),
            "객단가": format_accounting(tot_c_aov, is_currency=True),
            "배송건수": format_accounting(tot_c_deliv, decimals=0),
            "배송비": format_accounting(tot_c_fee, is_currency=True),
        }

        render_total_summary_bar(list(df_disp.columns), total_ch_row, is_selected=is_ch_sel)

    with tab_ch_edit:
        st.caption("천 단위 쉼표(,) 및 0원은 '-' 회계 처리가 적용된 편집 전용 표입니다. 셀을 직접 더블 클릭하여 수정 후 저장 버튼을 누르세요.")
        display_editor_df = format_table_for_accounting_view(df_raw)

        edited_df = st.data_editor(
            display_editor_df,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_{db_key}"
        )

        btn_col1, btn_col2 = st.columns([1.5, 3])
        with btn_col1:
            if st.button("💾 변경사항 저장", type="primary", use_container_width=True, key=f"btn_save_{db_key}"):
                save_ready_df = edited_df.copy()
                for col in NUMERIC_COLUMNS:
                    if col in save_ready_df.columns:
                        save_ready_df[col] = save_ready_df[col].apply(clean_numeric_val)
                if "년도" in save_ready_df.columns:
                    save_ready_df["년도"] = pd.to_numeric(save_ready_df["년도"], errors="coerce").fillna(2026).astype(int)
                save_db(db_key, save_ready_df)
                st.success("변경사항이 성공적으로 저장되었습니다.")
                st.rerun()

    # 4. 데이터 삭제 및 관리 도구
    with st.expander("⚠️ **데이터 삭제 및 관리 도구**", expanded=False):
        st.markdown(f"**현재 등록된 전체 행 수**: `{len(df_raw)}건`")
        del_col1, del_col2 = st.columns(2)

        with del_col1:
            st.markdown("##### 🗑️ 특정 기간 데이터 삭제")
            st.caption("원하는 연도와 월의 데이터만 골라서 삭제합니다.")
            del_year = st.selectbox("삭제할 연도", available_years, key=f"del_y_{db_key}")
            del_month = st.selectbox("삭제할 월 (선택)", ["전체 (해당 연도 모두)"] + available_months, key=f"del_m_{db_key}")
            del_m_val = None if del_month.startswith("전체") else del_month

            if st.button(f"🗑️ {del_year}년 {del_month} 데이터 삭제", key=f"btn_del_p_{db_key}"):
                deleted = delete_period_data(db_key, del_year, del_m_val)
                st.warning(f"{deleted}건의 데이터가 삭제되었습니다.")
                st.rerun()

        with del_col2:
            st.markdown("##### 🧹 전체 데이터 비우기 (초기화)")
            st.caption("이 항목의 모든 데이터를 완전히 비웁니다.")
            confirm_clear = st.checkbox("정말로 모든 데이터를 삭제하시겠습니까?", key=f"chk_c_{db_key}")
            if confirm_clear:
                if st.button("🚨 전체 비우기 실행", type="secondary", key=f"btn_c_{db_key}"):
                    clear_db(db_key)
                    st.error("모든 데이터가 삭제되었습니다.")
                    st.rerun()


# ====================================================
# 2. 네이버 광고 페이지
# ====================================================
if "네이버" in selected_page:
    render_channel_page("DB_네이버", "🟩 네이버 광고", COLORS["네이버"], cat_column="마케팅 분류")

# ====================================================
# 3. 쿠팡 광고 페이지
# ====================================================
elif "쿠팡" in selected_page:
    render_channel_page("DB_쿠팡", "🚀 쿠팡 광고", COLORS["쿠팡"], cat_column=None)

# ====================================================
# 4. 메타 광고 페이지
# ====================================================
elif "메타" in selected_page:
    render_channel_page("DB_메타", "♾️ 메타 광고", COLORS["메타"], cat_column=None)

# ====================================================
# 5. 전체 매출 페이지
# ====================================================
elif "매출" in selected_page:
    render_channel_page("DB_매출", "💰 전체 매출", COLORS["매출(전체)"], cat_column="매출분류")

# ====================================================
# 6. 데이터 통합 관리 및 초기화 페이지
# ====================================================
elif "데이터 관리" in selected_page:
    st.markdown("<h2>🛠️ 데이터 관리</h2>", unsafe_allow_html=True)
    st.caption("새로운 광고/매출 보고서 파일 업로드 및 자동 반영, 등록 현황 확인, 일괄 삭제 및 복원을 관리합니다.")

    # ------------------------------------------------
    # [1] 스마트 파일 업로드 & 자동 연동 영역 (데이터 관리 최상단)
    # ------------------------------------------------
    with st.expander("📂 **엑셀 / CSV 보고서 파일 업로드 및 자동 반영**", expanded=True):
        st.markdown("""
        **💡 광고 및 매출 보고서 자동 연동 안내**
        * **보고서 원본 호환**: 네이버 검색광고/GFA, 쿠팡 매출성장/검색광고, 메타 광고관리자, 스마트스토어/오픈마켓 정산 엑셀을 수정 없이 원본 그대로 업로드하시면 자동으로 양식을 감지하여 변환합니다.
        * **반영 방식 안내**: 
          - **➕ 누적 추가 (기본)**: 기존 데이터를 안전하게 보존하면서 새로운 데이터를 덧붙여 누적 집계합니다.
          - **🔄 덮어쓰기**: 선택하신 귀속 연월의 해당 채널 데이터를 새로 올린 파일로 완전히 교체합니다.
        """)

        # 1행: 설정 옵션 (파일 종류, 귀속 연월, 반영 방식 - 칼정렬 50% : 25% : 25%)
        opt_c1, opt_c2, opt_c3 = st.columns([2.0, 1.0, 1.0])

        with opt_c1:
            db_type_choices = [
                "자동 감지 (통합 엑셀 / 원본 보고서)",
                "네이버 광고",
                "쿠팡 광고",
                "메타 광고",
                "전체 매출"
            ]
            db_key_mapping = {
                "자동 감지 (통합 엑셀 / 원본 보고서)": "자동 감지",
                "네이버 광고": "DB_네이버",
                "쿠팡 광고": "DB_쿠팡",
                "메타 광고": "DB_메타",
                "전체 매출": "DB_매출"
            }
            selected_choice = st.selectbox("파일 종류 선택", db_type_choices)
            selected_db_type = db_key_mapping[selected_choice]

        with opt_c2:
            default_m_str = selected_month if (selected_month and selected_month.startswith(f"{selected_year}-")) else f"{selected_year}-01"
            input_period = st.text_input("귀속 연월 (YYYY-MM)", value=default_m_str)
            try:
                p_parts = input_period.strip().split("-")
                attr_year = int(p_parts[0])
                attr_month = f"{p_parts[0]}-{p_parts[1].zfill(2)}"
            except Exception:
                attr_year = selected_year
                attr_month = f"{selected_year}-01"

        with opt_c3:
            import_mode_choice = st.radio("반영 방식", ["🔄 덮어쓰기", "➕ 누적 추가"], index=1, horizontal=True)
            mode_param = "append" if "누적" in import_mode_choice else "overwrite"

        st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)

        # 2행: 파일 선택 드래그앤드롭 + 즉시 적용 버튼 + 연동 방식 안내 (칼정렬 50% : 25% : 25%)
        file_c1, file_c2, file_c3 = st.columns([2.0, 1.0, 1.0])
        with file_c1:
            uploaded_file = st.file_uploader(
                "보고서 파일 선택 (.xlsx, .xls, .csv)",
                type=["xlsx", "xls", "csv"],
                key="main_uploader"
            )

        with file_c2:
            st.markdown('<div class="upload-apply-label">즉시 실행</div>', unsafe_allow_html=True)
            apply_clicked = st.button("🚀 데이터 즉시 적용", type="primary", use_container_width=True, key="btn_apply_upload")

        with file_c3:
            st.markdown('<div class="upload-apply-label">연동 방식 안내</div>', unsafe_allow_html=True)
            st.markdown("""
            <div class="upload-info-box">
                <div style="font-size: 0.85rem; color: #444746; line-height: 1.45;">
                    • <b>덮어쓰기</b>: 해당 연월 데이터 교체<br>
                    • <b>누적 추가</b>: 기존 데이터 보존 누적
                </div>
            </div>
            """, unsafe_allow_html=True)

        if apply_clicked:
            if uploaded_file is None:
                st.warning("⚠️ 업로드할 보고서 파일(.xlsx, .csv)을 먼저 선택해주세요.")
            else:
                with st.spinner("보고서를 분석하여 데이터에 반영 중입니다..."):
                    f_bytes = uploaded_file.read()
                    success, msg, counts = process_uploaded_file_v2(
                        file_bytes=f_bytes,
                        filename=uploaded_file.name,
                        selected_db_type=selected_db_type,
                        target_year=attr_year,
                        target_month=attr_month,
                        import_mode=mode_param
                    )
                    if success:
                        st.success(f"✅ {msg}")
                        st.toast("데이터가 성공적으로 업데이트되었습니다!", icon="🎉")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")

    st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)

    # 현황 요약
    st.markdown("<h3>📊 현재 데이터 등록 현황</h3>", unsafe_allow_html=True)
    status_cols = st.columns(4)
    channel_labels = ["네이버 광고", "쿠팡 광고", "메타 광고", "전체 매출"]
    channel_colors = [
        COLORS.get("네이버", "#03C75A"),
        COLORS.get("쿠팡", "#EA1D2C"),
        COLORS.get("메타", "#0084FF"),
        COLORS.get("매출(전체)", "#00C56F")
    ]
    for i, db_key in enumerate(DB_KEYS):
        cnt = len(all_dbs.get(db_key, pd.DataFrame()))
        with status_cols[i]:
            render_google_kpi_card(channel_labels[i], f"{cnt:,} 건", "등록된 데이터 행 수", channel_colors[i])

    st.markdown("---")

    col_m1, col_m2 = st.columns(2)

    with col_m1:
        st.markdown("<h3>🗑️ 기간별 데이터 일괄 삭제</h3>", unsafe_allow_html=True)
        st.caption("선택한 연도와 월에 해당하는 모든 채널(네이버, 쿠팡, 메타, 매출)의 데이터를 한 번에 제거합니다.")
        batch_year = st.selectbox("일괄 삭제 대상 연도", available_years, key="batch_del_y")
        batch_month = st.selectbox("일괄 삭제 대상 월", ["전체 (연간 전체)"] + available_months, key="batch_del_m")
        batch_m_val = None if batch_month.startswith("전체") else batch_month

        if st.button("🗑️ 선택 기간 데이터 일괄 삭제", type="secondary"):
            total_del = 0
            for k in DB_KEYS:
                total_del += delete_period_data(k, batch_year, batch_m_val)
            st.warning(f"총 {total_del}건의 데이터가 삭제되었습니다.")
            st.rerun()

    with col_m2:
        st.markdown("<h3>🔄 데이터 초기화 및 복원</h3>", unsafe_allow_html=True)
        st.caption("데이터를 테스트 중 초기화하거나 기본 2027년 샘플 상태로 되돌립니다.")

        if st.button("🔄 기본 샘플 데이터로 복원하기", type="primary", use_container_width=True):
            from src.generate_samples import generate_all_samples
            generate_all_samples()
            st.success("기본 샘플 데이터로 복원되었습니다.")
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("##### 🚨 모든 데이터 완전 삭제 (새 데이터 등록용)")
        confirm_all_clear = st.checkbox("정말로 모든 채널의 데이터를 완전히 비우시겠습니까? (복구 불가)")
        if confirm_all_clear:
            if st.button("🚨 모든 데이터 비우기 실행"):
                clear_all_dbs()
                st.error("모든 데이터가 비워졌습니다. 이제 새로운 엑셀 보고서를 업로드해 보세요.")
                st.rerun()
