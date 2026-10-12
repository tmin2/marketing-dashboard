"""
Plotly 기반 인터랙티브 시각화 차트 모듈입니다.
- 차트 내부 타이틀을 제거하고 외부 헤더로 분리하여 범례와의 텍스트 겹침을 원천 차단합니다.
- 날짜 축 깨짐 버그를 방지하기 위해 한글 범주형 레이블('YYYY년 MM월')을 적용합니다.
- 모던하고 깔끔한 구글 머티리얼 디자인 컬러 팔레트를 적용합니다.
모든 함수는 명확한 타입 힌트와 상세 주석을 포함합니다.
"""

from typing import Dict, List, Optional
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

# 브랜드 및 메트릭 테마 색상 정의 (이숲 esoop.co.kr 디자인 시스템 팔레트)
COLORS: Dict[str, str] = {
    "네이버": "#03C75A",       # 네이버 그린
    "쿠팡": "#EA1D2C",         # 쿠팡 레드
    "메타": "#0084FF",         # 메타 블루
    "매출(전체)": "#00C56F",    # 이숲 시그니처 에메랄드 그린
    "광고비": "#FF5E57",       # 부드러운 산호빛 코랄 레드
    "매출": "#00C56F",         # 이숲 에메랄드 그린
    "ROAS": "#06291B",         # 이숲 딥 포레스트
    "CPA": "#F59E0B",          # 웜 앰버 골드
    "배경": "#FFFFFF",
    "그리드": "#F0F4F2"
}


def format_month_label(m_str: str) -> str:
    """
    '2026-10' 같은 월 문자열을 Plotly의 날짜 자동 변환 버그를 방지하고
    가독성을 극대화하기 위해 '2026년 10월' 형태의 명확한 한글 레이블로 변환합니다.

    Args:
        m_str: 원본 월 문자열 (예: '2026-10')

    Returns:
        str: 한글 표기 월 문자열 (예: '2026년 10월')
    """
    clean_val = str(m_str).strip()
    if "-" in clean_val:
        parts = clean_val.split("-")
        if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
            return f"{parts[0]}년 {int(parts[1])}월"
    return f"{clean_val} "


def plot_monthly_revenue_and_cost(monthly_df: pd.DataFrame, metric_mode: str = "💰 광고비 vs 매출 & ROAS") -> go.Figure:
    """
    월별 성과 추이를 다양한 지표 필터링 모드로 시각화합니다.
    제목을 차트 외부로 분리하여 상단 범례와의 텍스트 겹침을 원천 차단합니다.

    Args:
        monthly_df: 월별 성과 요약 데이터프레임
        metric_mode: 분석할 지표 조합 (기본: '💰 광고비 vs 매출 & ROAS')

    Returns:
        go.Figure: Plotly 복합 차트 객체
    """
    fig = go.Figure()

    if monthly_df.empty:
        return fig

    # 날짜 자동 파싱 방지용 안전 범주형 레이블 생성
    x_labels: List[str] = [format_month_label(m) for m in monthly_df["월"]]
    raw_months: List[str] = [str(m).strip() for m in monthly_df["월"]]

    # 단일 월 데이터일 때 막대 두께 조정
    gap_val = 0.5 if len(monthly_df) <= 2 else 0.2

    clean_mode = metric_mode.replace("💰", "").replace("👆", "").replace("🎯", "").replace("💳", "").replace("🏷️", "").strip()

    if "클릭수 vs 노출수 & CTR" in clean_mode or "노출수" in clean_mode:
        # 1. 노출수 막대
        if "노출수" in monthly_df.columns:
            fig.add_trace(
                go.Bar(
                    x=x_labels,
                    y=monthly_df["노출수"],
                    name="노출수 (회)",
                    marker_color="#64748B",
                    opacity=0.88,
                    customdata=raw_months,
                    hovertemplate="%{x}<br>노출수: %{y:,.0f}회<extra></extra>"
                )
            )
        # 2. 클릭수 막대
        if "클릭수" in monthly_df.columns:
            fig.add_trace(
                go.Bar(
                    x=x_labels,
                    y=monthly_df["클릭수"],
                    name="클릭수 (회)",
                    marker_color="#3B82F6",
                    opacity=0.88,
                    customdata=raw_months,
                    hovertemplate="%{x}<br>클릭수: %{y:,.0f}회<extra></extra>"
                )
            )
        # 3. CTR 라인 (보조축)
        if "CTR" in monthly_df.columns:
            fig.add_trace(
                go.Scatter(
                    x=x_labels,
                    y=monthly_df["CTR"],
                    name="CTR (%)",
                    mode="lines+markers",
                    line=dict(color="#10B981", width=3),
                    marker=dict(size=8, symbol="diamond"),
                    yaxis="y2",
                    customdata=raw_months,
                    hovertemplate="%{x}<br>CTR: %{y:.2f}%<extra></extra>"
                )
            )
        y1_title, y2_title, y2_fmt = "수량 (회)", "CTR (%)", ".2f"

    elif "광고비 vs 전환수" in clean_mode or ("전환수" in clean_mode and "광고비" in clean_mode):
        # 1. 총 광고비 막대 (좌측 Y1축)
        fig.add_trace(
            go.Bar(
                x=x_labels,
                y=monthly_df["총 광고비"],
                name="총 광고비 (원)",
                marker_color="#FF5E57",
                opacity=0.88,
                customdata=raw_months,
                hovertemplate="%{x}<br>총 광고비: ₩%{y:,.0f}<extra></extra>"
            )
        )
        # 2. 전환수 꺾은선 (우측 Y2축 보조축, 이숲 에메랄드 그린 컬러)
        if "전환수" in monthly_df.columns:
            # 툴팁에 건당 CPA 정보 연동
            cpa_data = monthly_df["CPA"].tolist() if "CPA" in monthly_df.columns else [0] * len(monthly_df)
            fig.add_trace(
                go.Scatter(
                    x=x_labels,
                    y=monthly_df["전환수"],
                    name="전환수 (건)",
                    mode="lines+markers",
                    line=dict(color="#00C56F", width=3.5),
                    marker=dict(size=8, symbol="circle"),
                    yaxis="y2",
                    customdata=cpa_data,
                    hovertemplate="%{x}<br>전환수: %{y:,.0f}건<br>CPA: ₩%{customdata:,.0f}<extra></extra>"
                )
            )
        y1_title, y2_title, y2_fmt = "총 광고비 (원)", "전환수 (건)", ",.0f"

    elif "전환수 vs CPA" in clean_mode or "CPA" in clean_mode:
        # 1. 전환수 막대 (좌측 Y1축, 이숲 에메랄드 그린 컬러)
        if "전환수" in monthly_df.columns:
            fig.add_trace(
                go.Bar(
                    x=x_labels,
                    y=monthly_df["전환수"],
                    name="전환수 (건)",
                    marker_color="#00C56F",
                    opacity=0.88,
                    customdata=raw_months,
                    hovertemplate="%{x}<br>전환수: %{y:,.0f}건<extra></extra>"
                )
            )
        # 2. CPA 꺾은선 (우측 Y2축 보조축, 웜 앰버 골드 컬러)
        if "CPA" in monthly_df.columns:
            fig.add_trace(
                go.Scatter(
                    x=x_labels,
                    y=monthly_df["CPA"],
                    name="CPA (원)",
                    mode="lines+markers",
                    line=dict(color="#F59E0B", width=3),
                    marker=dict(size=8, symbol="diamond"),
                    yaxis="y2",
                    customdata=raw_months,
                    hovertemplate="%{x}<br>CPA: ₩%{y:,.0f}<extra></extra>"
                )
            )
        y1_title, y2_title, y2_fmt = "전환수 (건)", "CPA (원)", ",.0f"


    elif "객단가 vs 전환율" in clean_mode or "CVR" in clean_mode:
        # 1. 객단가 막대
        fig.add_trace(
            go.Bar(
                x=x_labels,
                y=monthly_df["객단가"],
                name="객단가 (원)",
                marker_color="#8B5CF6",
                opacity=0.88,
                customdata=raw_months,
                hovertemplate="%{x}<br>객단가: ₩%{y:,.0f}<extra></extra>"
            )
        )
        # 2. CVR 라인 (보조축)
        fig.add_trace(
            go.Scatter(
                x=x_labels,
                y=monthly_df["CVR"],
                name="CVR (%)",
                mode="lines+markers",
                line=dict(color="#0284C7", width=3),
                marker=dict(size=8, symbol="square"),
                yaxis="y2",
                customdata=raw_months,
                hovertemplate="%{x}<br>CVR: %{y:.2f}%<extra></extra>"
            )
        )
        y1_title, y2_title, y2_fmt = "객단가 (원)", "CVR (%)", ".2f"

    else:
        # 기본: 총 광고비 vs 총 매출 & ROAS
        fig.add_trace(
            go.Bar(
                x=x_labels,
                y=monthly_df["총 광고비"],
                name="총 광고비 (원)",
                marker_color=COLORS["광고비"],
                opacity=0.88,
                customdata=raw_months,
                hovertemplate="%{x}<br>광고비: ₩%{y:,.0f}<extra></extra>"
            )
        )
        fig.add_trace(
            go.Bar(
                x=x_labels,
                y=monthly_df["총 매출"],
                name="총 매출 (원)",
                marker_color=COLORS["매출"],
                opacity=0.88,
                customdata=raw_months,
                hovertemplate="%{x}<br>매출: ₩%{y:,.0f}<extra></extra>"
            )
        )
        fig.add_trace(
            go.Scatter(
                x=x_labels,
                y=monthly_df["ROAS"],
                name="ROAS (배)",
                mode="lines+markers",
                line=dict(color=COLORS["ROAS"], width=3),
                marker=dict(size=8, symbol="diamond"),
                yaxis="y2",
                customdata=raw_months,
                hovertemplate="%{x}<br>ROAS: %{y:.2f}x<extra></extra>"
            )
        )
        y1_title, y2_title, y2_fmt = "금액 (원)", "ROAS (배수)", ".2f"

    fig.update_layout(
        font=dict(family="Pretendard, -apple-system, sans-serif"),
        title=None,
        barmode="group",
        bargap=gap_val,
        hovermode="x unified",
        clickmode="event+select",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(
            title="",
            showgrid=False,
            type="category",
            categoryorder="array",
            categoryarray=x_labels
        ),
        yaxis=dict(
            title=y1_title,
            showgrid=True,
            gridcolor=COLORS["그리드"],
            tickformat=",.0f"
        ),
        yaxis2=dict(
            title=y2_title,
            overlaying="y",
            side="right",
            showgrid=False,
            tickformat=y2_fmt
        ),
        margin=dict(l=30, r=30, t=35, b=25),
        plot_bgcolor="white",
        paper_bgcolor="white"
    )

    return fig


def plot_channel_comparison_bar(channel_df: pd.DataFrame) -> go.Figure:
    """
    네이버, 쿠팡, 메타 채널별 광고비와 전환매출을 비교하는 막대 차트를 생성합니다.

    Args:
        channel_df: '채널', '광고비', '전환매출' 컬럼을 갖는 데이터프레임

    Returns:
        go.Figure: Plotly 막대 차트 객체
    """
    df_plot = channel_df[channel_df["채널"] != "매출(전체)"].copy()
    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=df_plot["채널"],
            y=df_plot["광고비"],
            name="광고비",
            marker_color=COLORS["광고비"],
            hovertemplate="%{x}<br>광고비: ₩%{y:,.0f}<extra></extra>"
        )
    )

    fig.add_trace(
        go.Bar(
            x=df_plot["채널"],
            y=df_plot["전환매출"],
            name="전환매출",
            marker_color=COLORS["매출"],
            hovertemplate="%{x}<br>전환매출: ₩%{y:,.0f}<extra></extra>"
        )
    )

    fig.update_layout(
        font=dict(family="Pretendard, -apple-system, sans-serif"),
        title=None,
        barmode="group",
        xaxis=dict(title="", type="category"),
        yaxis=dict(title="금액 (원)", gridcolor=COLORS["그리드"], tickformat=",.0f"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=30, r=30, t=35, b=25)
    )

    return fig


def plot_channel_cost_share(channel_df: pd.DataFrame, metric: str = "광고비") -> go.Figure:
    """
    채널별 광고 성과 비중을 도넛 차트로 시각화합니다 (광고비, 전환매출, 전환수, 클릭수, 노출수 등 동적 전환).

    Args:
        channel_df: '채널' 및 성과 지표 컬럼을 포함하는 데이터프레임
        metric: 비중을 계산할 기준 지표 (예: '광고비', '전환매출', '전환수', '클릭수', '노출수')

    Returns:
        go.Figure: Plotly 도넛 차트 객체
    """
    df_plot = channel_df[channel_df["채널"] != "매출(전체)"].copy()
    color_map = {
        "네이버": COLORS["네이버"],
        "쿠팡": COLORS["쿠팡"],
        "메타": COLORS["메타"]
    }

    clean_metric = metric.replace("💸", "").replace("💰", "").replace("🎯", "").replace("👆", "").replace("👁️", "").strip()
    target_col = clean_metric if clean_metric in df_plot.columns else "광고비"

    # 화폐 단위 여부 판별
    is_currency = target_col in ["광고비", "전환매출"]
    unit_str = "원" if is_currency else ("건" if "전환수" in target_col else "회")
    fmt_prefix = "₩" if is_currency else ""

    fig = go.Figure(
        data=[
            go.Pie(
                labels=df_plot["채널"],
                values=df_plot[target_col],
                customdata=df_plot["채널"],
                hole=0.45,
                marker=dict(colors=[color_map.get(ch, "#a4b0be") for ch in df_plot["채널"]]),
                textinfo="label+percent",
                textposition="inside",
                insidetextfont=dict(color="#ffffff", size=13, family="Pretendard, sans-serif"),
                outsidetextfont=dict(color="#06291B", size=12, family="Pretendard, sans-serif"),
                textfont=dict(color="#ffffff", size=13, family="Pretendard, sans-serif"),
                hovertemplate=f"%{{label}}<br>{target_col}: {fmt_prefix}%{{value:,.0f}}{unit_str} (%{{percent}})<extra></extra>"
            )
        ]
    )

    fig.update_layout(
        font=dict(family="Pretendard, -apple-system, sans-serif"),
        title=None,
        clickmode="event+select",
        legend=dict(orientation="h", yanchor="top", y=-0.08, xanchor="center", x=0.5),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=20, b=40)
    )

    return fig


def plot_sales_channel_share(sales_df: pd.DataFrame, metric: str = "매출액") -> go.Figure:
    """
    판매 채널별 실적 비중을 도넛 차트로 시각화합니다 (매출액, 판매수량, 배송건수, 배송비, 객단가 등 동적 전환).

    Args:
        sales_df: '판매 채널' 및 실적 컬럼을 포함하는 데이터프레임
        metric: 비중을 계산할 기준 지표 (예: '매출액', '판매수량', '배송건수', '배송비', '객단가')

    Returns:
        go.Figure: Plotly 도넛 차트 객체
    """
    if sales_df.empty or "판매 채널" not in sales_df.columns:
        return go.Figure()

    clean_metric = metric.replace("💰", "").replace("📦", "").replace("🚚", "").replace("💳", "").replace("🏷️", "").strip()
    target_col = clean_metric if clean_metric in sales_df.columns else "매출액"

    df_plot = sales_df[sales_df[target_col] > 0].copy()
    if df_plot.empty:
        return go.Figure()

    # 이숲 감성의 세련된 다채널 컬러 팔레트
    channel_colors = [
        "#00C56F",  # 이숲 에메랄드 그린
        "#06291B",  # 딥 포레스트
        "#10B981",  # 브라이트 민트
        "#0284C7",  # 스카이 블루
        "#8B5CF6",  # 바이올렛
        "#F59E0B",  # 앰버 골드
        "#EC4899",  # 로즈 핑크
        "#14B8A6",  # 틸
        "#64748B"   # 슬레이트
    ]

    is_currency = target_col in ["매출액", "배송비", "객단가"]
    unit_str = "원" if is_currency else ("개" if "수량" in target_col else "건")
    fmt_prefix = "₩" if is_currency else ""

    fig = go.Figure(
        data=[
            go.Pie(
                labels=df_plot["판매 채널"],
                values=df_plot[target_col],
                customdata=df_plot["판매 채널"],
                hole=0.45,
                marker=dict(colors=channel_colors[:len(df_plot)]),
                textinfo="label+percent",
                textposition="inside",
                insidetextfont=dict(color="#ffffff", size=12, family="Pretendard, sans-serif"),
                outsidetextfont=dict(color="#06291B", size=11, family="Pretendard, sans-serif"),
                textfont=dict(color="#ffffff", size=12, family="Pretendard, sans-serif"),
                hovertemplate=f"%{{label}}<br>{target_col}: {fmt_prefix}%{{value:,.0f}}{unit_str} (%{{percent}})<extra></extra>"
            )
        ]
    )

    fig.update_layout(
        font=dict(family="Pretendard, -apple-system, sans-serif"),
        title=None,
        clickmode="event+select",
        legend=dict(orientation="h", yanchor="top", y=-0.08, xanchor="center", x=0.5),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=20, b=40)
    )

    return fig


def plot_pie_category(df: pd.DataFrame, category_col: str, value_col: str) -> go.Figure:
    """
    특정 카테고리 컬럼을 기준으로 값의 분포를 도넛 차트로 나타냅니다.

    Args:
        df: 데이터프레임
        category_col: 분류 기준 컬럼 (예: '마케팅 분류' or '매출분류')
        value_col: 집계할 값 컬럼 (예: '광고비' or '전환매출')

    Returns:
        go.Figure: 도넛 차트 객체
    """
    if df.empty or category_col not in df.columns or value_col not in df.columns:
        return go.Figure()

    grouped = df.groupby(category_col)[value_col].sum().reset_index()
    grouped = grouped[grouped[value_col] > 0]

    fig = px.pie(
        grouped,
        names=category_col,
        values=value_col,
        hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Pastel
    )

    fig.update_traces(
        textinfo="label+percent",
        hovertemplate="%{label}<br>" + value_col + ": ₩%{value:,.0f} (%{percent})<extra></extra>"
    )

    fig.update_layout(
        font=dict(family="Pretendard, -apple-system, sans-serif"),
        title=None,
        legend=dict(orientation="h", yanchor="top", y=-0.10, xanchor="center", x=0.5),
        margin=dict(l=20, r=20, t=20, b=40),
        plot_bgcolor="white",
        paper_bgcolor="white"
    )

    return fig


def plot_channel_trend_line(df: pd.DataFrame, metric_cols: List[str]) -> go.Figure:
    """
    단일 채널의 월별 주요 메트릭 추이를 다중 라인 차트로 표현합니다.

    Args:
        df: 해당 채널의 데이터프레임
        metric_cols: 시각화할 수치 컬럼 목록 (예: ['광고비', '전환매출'])

    Returns:
        go.Figure: 다중 라인 차트 객체
    """
    if df.empty or "월" not in df.columns:
        return go.Figure()

    grouped = df.groupby("월")[metric_cols].sum().reset_index().sort_values("월")
    x_labels: List[str] = [format_month_label(m) for m in grouped["월"]]

    fig = go.Figure()
    palette = [COLORS["광고비"], COLORS["매출"], COLORS["ROAS"], COLORS["CPA"]]

    for i, col in enumerate(metric_cols):
        fig.add_trace(
            go.Scatter(
                x=x_labels,
                y=grouped[col],
                name=col,
                mode="lines+markers",
                line=dict(width=2.5, color=palette[i % len(palette)]),
                marker=dict(size=7),
                hovertemplate="%{x}<br>" + col + ": %{y:,.0f}<extra></extra>"
            )
        )

    fig.update_layout(
        font=dict(family="Pretendard, -apple-system, sans-serif"),
        title=None,
        xaxis=dict(
            title="",
            showgrid=False,
            type="category",
            categoryorder="array",
            categoryarray=x_labels
        ),
        yaxis=dict(title="수치", gridcolor=COLORS["그리드"], tickformat=",.0f"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=30, r=30, t=35, b=25)
    )

    return fig
