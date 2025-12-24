import io
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots


# -----------------------------
# App config & fonts
# -----------------------------
st.set_page_config(
    page_title="극지식물 최적 EC 농도 연구",
    page_icon="🌱",
    layout="wide",
)

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR&display=swap');
html, body, [class*="css"] {
    font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}
</style>
""",
    unsafe_allow_html=True,
)

PLOTLY_FONT = "Malgun Gothic, Apple SD Gothic Neo, Noto Sans KR, sans-serif"


# -----------------------------
# Constants (given by project)
# -----------------------------
SCHOOL_ORDER = ["송도고", "하늘고", "아라고", "동산고"]

TARGET_EC_BY_SCHOOL = {
    "송도고": 1.0,
    "하늘고": 2.0,  # (최적 표시)
    "아라고": 4.0,
    "동산고": 8.0,
}

SCHOOL_COLOR = {
    "송도고": "#1f77b4",
    "하늘고": "#2ca02c",
    "아라고": "#ff7f0e",
    "동산고": "#d62728",
}


# -----------------------------
# Unicode-safe helpers
# -----------------------------
def _norm_all(s: str) -> Tuple[str, str]:
    """Return (NFC, NFD) for robust comparisons."""
    return (unicodedata.normalize("NFC", s), unicodedata.normalize("NFD", s))


def _same_name(a: str, b: str) -> bool:
    """Unicode-safe, case-insensitive name equality."""
    a_nfc, a_nfd = _norm_all(a)
    b_nfc, b_nfd = _norm_all(b)
    return (a_nfc.lower() == b_nfc.lower()) or (a_nfd.lower() == b_nfd.lower())


def find_file_by_normalized_name(folder: Path, wanted_name: str) -> Optional[Path]:
    """
    반드시 Path.iterdir()로 탐색하고,
    NFC/NFD 양방향 비교로 파일을 찾는다.
    """
    if not folder.exists():
        return None

    wanted_stems = _norm_all(wanted_name)
    for p in folder.iterdir():
        if not p.is_file():
            continue

        # compare file name (full) with both NFC/NFD
        name_nfc, name_nfd = _norm_all(p.name)
        if (name_nfc == wanted_stems[0]) or (name_nfd == wanted_stems[1]):
            return p

        # also allow "same_name" to handle subtle differences
        if _same_name(p.name, wanted_name):
            return p

    return None


def detect_school_from_filename(filename: str) -> Optional[str]:
    """
    파일명에서 학교명을 Unicode-safe하게 추정.
    예: '송도고_환경데이터.csv'
    """
    for school in SCHOOL_ORDER:
        if school in filename:
            return school
        # extra safe check with normalized contains
        fn_nfc, fn_nfd = _norm_all(filename)
        sc_nfc, sc_nfd = _norm_all(school)
        if (sc_nfc in fn_nfc) or (sc_nfd in fn_nfd):
            return school
    return None


# -----------------------------
# Data loading
# -----------------------------
@st.cache_data(show_spinner=False)
def load_environment_data(data_dir: Path) -> pd.DataFrame:
    """
    환경 CSV 4개 로드:
    columns: time, temperature, humidity, ph, ec
    학교별 측정 주기 다름 -> time은 datetime 파싱
    """
    rows: List[pd.DataFrame] = []

    if not data_dir.exists():
        return pd.DataFrame()

    for p in data_dir.iterdir():  # 필수: iterdir()
        if not p.is_file():
            continue

        # CSV만
        if p.suffix.lower() != ".csv":
            continue

        school = detect_school_from_filename(p.name)
        if school is None:
            continue

        try:
            df = pd.read_csv(p)
        except Exception:
            continue

        # normalize column names
        df.columns = [str(c).strip().lower() for c in df.columns]

        needed = {"time", "temperature", "humidity", "ph", "ec"}
        if not needed.issubset(set(df.columns)):
            continue

        # time parsing
        df["time"] = pd.to_datetime(df["time"], errors="coerce")
        df = df.dropna(subset=["time"])

        # numeric parsing
        for col in ["temperature", "humidity", "ph", "ec"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        df["school"] = school
        rows.append(df)

    if not rows:
        return pd.DataFrame()

    out = pd.concat(rows, ignore_index=True)
    return out


@st.cache_data(show_spinner=False)
def load_growth_data(data_dir: Path) -> pd.DataFrame:
    """
    XLSX 1개, 4개 시트 자동 로드(시트명 하드코딩 금지)
    columns: 개체번호, 잎 수(장), 지상부 길이(mm), 지하부길이(mm), 생중량(g)
    """
    xlsx_path = find_file_by_normalized_name(data_dir, "4개교_생육결과데이터.xlsx")
    if xlsx_path is None:
        return pd.DataFrame()

    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        return pd.DataFrame()

    all_frames: List[pd.DataFrame] = []

    # 시트명 하드코딩 금지: xls.sheet_names 사용
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xlsx_path, sheet_name=sheet, engine="openpyxl")
        except Exception:
            continue

        # 시트 이름에서 학교 매칭(Unicode-safe)
        school = None
        for s in SCHOOL_ORDER:
            if _same_name(sheet, s) or (s in sheet):
                school = s
                break
        if school is None:
            # 혹시 파일명/시트명에 '동산고등학교' 같은 변형이 있을 수 있어 contains로 재시도
            for s in SCHOOL_ORDER:
                sh_nfc, sh_nfd = _norm_all(sheet)
                sc_nfc, sc_nfd = _norm_all(s)
                if (sc_nfc in sh_nfc) or (sc_nfd in sh_nfd):
                    school = s
                    break

        if school is None:
            continue

        # columns (Korean)
        expected_cols = ["개체번호", "잎 수(장)", "지상부 길이(mm)", "지하부길이(mm)", "생중량(g)"]
        # allow slight whitespace variants
        df.columns = [str(c).strip() for c in df.columns]

        if not set(expected_cols).issubset(set(df.columns)):
            continue

        for c in expected_cols:
            if c != "개체번호":
                df[c] = pd.to_numeric(df[c], errors="coerce")

        df["school"] = school
        df["target_ec"] = TARGET_EC_BY_SCHOOL.get(school)

        all_frames.append(df)

    if not all_frames:
        return pd.DataFrame()

    out = pd.concat(all_frames, ignore_index=True)
    return out


def dataframe_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "data") -> bytes:
    """XLSX 다운로드용 bytes (TypeError 방지: BytesIO 사용)"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# Derived metrics
# -----------------------------
def env_means(env: pd.DataFrame) -> pd.DataFrame:
    if env.empty:
        return pd.DataFrame()

    g = (
        env.groupby("school", as_index=False)
        .agg(
            avg_temp=("temperature", "mean"),
            avg_humid=("humidity", "mean"),
            avg_ph=("ph", "mean"),
            avg_ec=("ec", "mean"),
            n_points=("ec", "count"),
        )
    )

    # keep order
    g["school"] = pd.Categorical(g["school"], categories=SCHOOL_ORDER, ordered=True)
    g = g.sort_values("school")
    g["target_ec"] = g["school"].map(TARGET_EC_BY_SCHOOL)
    return g


def growth_means(growth: pd.DataFrame) -> pd.DataFrame:
    if growth.empty:
        return pd.DataFrame()

    g = (
        growth.groupby("school", as_index=False)
        .agg(
            n=("개체번호", "count"),
            avg_weight=("생중량(g)", "mean"),
            avg_leaf=("잎 수(장)", "mean"),
            avg_shoot=("지상부 길이(mm)", "mean"),
        )
    )
    g["school"] = pd.Categorical(g["school"], categories=SCHOOL_ORDER, ordered=True)
    g = g.sort_values("school")
    g["target_ec"] = g["school"].map(TARGET_EC_BY_SCHOOL)
    return g


# -----------------------------
# UI helpers
# -----------------------------
def metric_card_row(total_n: int, avg_temp: Optional[float], avg_humid: Optional[float], best_ec: Optional[float]):
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 개체수", f"{total_n:,}")
    c2.metric("평균 온도", "-" if avg_temp is None else f"{avg_temp:.2f} °C")
    c3.metric("평균 습도", "-" if avg_humid is None else f"{avg_humid:.2f} %")
    c4.metric("최적 EC", "-" if best_ec is None else f"{best_ec:.1f}")


def plotly_apply_font(fig: go.Figure) -> go.Figure:
    fig.update_layout(font=dict(family=PLOTLY_FONT))
    return fig


# -----------------------------
# Main
# -----------------------------
st.title("🌱 극지식물 최적 EC 농도 연구")

data_dir = Path(__file__).parent / "data"

with st.spinner("데이터를 불러오는 중..."):
    env_df = load_environment_data(data_dir)
    growth_df = load_growth_data(data_dir)

if env_df.empty:
    st.error("환경 데이터(CSV)를 찾지 못했습니다. data/ 폴더와 파일명을 확인해주세요.")
if growth_df.empty:
    st.error("생육 결과 데이터(XLSX)를 찾지 못했습니다. data/4개교_생육결과데이터.xlsx를 확인해주세요.")

# Sidebar
school_option = ["전체"] + SCHOOL_ORDER
selected_school = st.sidebar.selectbox("학교 선택", school_option, index=0)

# Filter
if selected_school != "전체":
    env_view = env_df[env_df["school"] == selected_school].copy()
    growth_view = growth_df[growth_df["school"] == selected_school].copy()
else:
    env_view = env_df.copy()
    growth_view = growth_df.copy()

env_summary = env_means(env_df)
growth_summary = growth_means(growth_df)

# Best EC (based on avg fresh weight by target EC)
best_ec = None
if not growth_summary.empty:
    best_row = growth_summary.sort_values("avg_weight", ascending=False).head(1)
    if not best_row.empty:
        best_ec = float(best_row["target_ec"].iloc[0])

# Global metrics (based on selected view)
total_n = int(growth_view["개체번호"].count()) if not growth_view.empty else 0
avg_temp = float(env_view["temperature"].mean()) if not env_view.empty else None
avg_humid = float(env_view["humidity"].mean()) if not env_view.empty else None


tab1, tab2, tab3 = st.tabs(["📖 실험 개요", "🌡️ 환경 데이터", "📊 생육 결과"])


# -----------------------------
# Tab 1: Overview
# -----------------------------
with tab1:
    st.subheader("연구 배경 및 목적")
    st.write(
        "극지식물은 야외 환경이 아닌 **극지연구소 스마트팜 내부**에서 재배되며, "
        "스마트팜에서는 **EC 농도·온도·습도** 같은 환경 요인을 정밀하게 제어할 수 있다. "
        "따라서 식물이 가장 잘 자라는 **최적 조건(EC, 온도, 습도)** 을 찾는 것이 중요하며, "
        "4개 학교의 실험 데이터를 기반으로 예측 모델을 구성하여 최적 환경을 추정한다."
    )

    st.markdown("#### 학교별 EC 조건")
    # counts from growth sheets
    counts = {}
    if not growth_summary.empty:
        counts = dict(zip(growth_summary["school"], growth_summary["n"]))
    ec_table = pd.DataFrame(
        {
            "학교명": SCHOOL_ORDER,
            "EC 목표": [TARGET_EC_BY_SCHOOL[s] for s in SCHOOL_ORDER],
            "개체수": [int(counts.get(s, 0)) for s in SCHOOL_ORDER],
            "색상(대시보드)": [SCHOOL_COLOR[s] for s in SCHOOL_ORDER],
        }
    )
    st.dataframe(ec_table, use_container_width=True)

    st.markdown("#### 주요 지표")
    metric_card_row(total_n=total_n, avg_temp=avg_temp, avg_humid=avg_humid, best_ec=best_ec)

    st.info("참고: ‘최적 EC’는 현재 데이터에서 **평균 생중량이 가장 큰 학교의 EC 목표값**으로 계산됩니다.")


# -----------------------------
# Tab 2: Environment
# -----------------------------
with tab2:
    st.subheader("학교별 환경 평균 비교")

    if env_summary.empty:
        st.error("환경 평균을 계산할 데이터가 없습니다.")
    else:
        # 2x2 subplots
        fig = make_subplots(
            rows=2,
            cols=2,
            subplot_titles=("평균 온도(°C)", "평균 습도(%)", "평균 pH", "목표 EC vs 실측 EC(평균)"),
        )

        # Avg temp bar
        fig.add_trace(
            go.Bar(
                x=env_summary["school"].astype(str),
                y=env_summary["avg_temp"],
                name="Avg Temp",
            ),
            row=1,
            col=1,
        )

        # Avg humid bar
        fig.add_trace(
            go.Bar(
                x=env_summary["school"].astype(str),
                y=env_summary["avg_humid"],
                name="Avg Humidity",
            ),
            row=1,
            col=2,
        )

        # Avg pH bar
        fig.add_trace(
            go.Bar(
                x=env_summary["school"].astype(str),
                y=env_summary["avg_ph"],
                name="Avg pH",
            ),
            row=2,
            col=1,
        )

        # Target vs actual EC (double bar)
        fig.add_trace(
            go.Bar(
                x=env_summary["school"].astype(str),
                y=env_summary["target_ec"],
                name="Target EC",
            ),
            row=2,
            col=2,
        )
        fig.add_trace(
            go.Bar(
                x=env_summary["school"].astype(str),
                y=env_summary["avg_ec"],
                name="Measured EC (mean)",
            ),
            row=2,
            col=2,
        )

        fig.update_layout(barmode="group", height=700, showlegend=True)
        fig = plotly_apply_font(fig)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("선택한 학교 시계열")

    if selected_school == "전체":
        st.caption("‘전체’를 선택하면 시계열이 복잡해질 수 있어, 학교를 하나 선택하는 것을 권장합니다.")
    if env_view.empty:
        st.error("선택한 조건에서 시계열을 표시할 환경 데이터가 없습니다.")
    else:
        # Time series line charts (Temp, Humid, EC with target line)
        env_view_sorted = env_view.sort_values("time")

        # Temperature
        fig_t = px.line(env_view_sorted, x="time", y="temperature", title="Temperature over Time")
        fig_t = plotly_apply_font(fig_t)
        st.plotly_chart(fig_t, use_container_width=True)

        # Humidity
        fig_h = px.line(env_view_sorted, x="time", y="humidity", title="Humidity over Time")
        fig_h = plotly_apply_font(fig_h)
        st.plotly_chart(fig_h, use_container_width=True)

        # EC with target horizontal line (if one school selected)
        fig_ec = px.line(env_view_sorted, x="time", y="ec", title="EC over Time")
        if selected_school != "전체":
            target = TARGET_EC_BY_SCHOOL.get(selected_school)
            if target is not None:
                fig_ec.add_hline(y=target, line_dash="dash", annotation_text="Target EC", annotation_position="top left")
        fig_ec = plotly_apply_font(fig_ec)
        st.plotly_chart(fig_ec, use_container_width=True)

    with st.expander("원본 환경 데이터 보기 / 다운로드"):
        if env_view.empty:
            st.error("표시할 환경 데이터가 없습니다.")
        else:
            st.dataframe(env_view, use_container_width=True)

            # CSV download
            csv_bytes = env_view.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label="CSV 다운로드",
                data=csv_bytes,
                file_name="환경데이터_선택.csv",
                mime="text/csv",
            )


# -----------------------------
# Tab 3: Growth results
# -----------------------------
with tab3:
    st.subheader("핵심 결과: EC별 평균 생중량")

    if growth_summary.empty:
        st.error("생육 결과 요약을 계산할 데이터가 없습니다.")
    else:
        # Card-like: show avg fresh weight per school (EC condition)
        # Highlight max
        g = growth_summary.copy()
        g["표시"] = g["school"].astype(str) + " (EC " + g["target_ec"].astype(float).map(lambda x: f"{x:.1f}") + ")"
        max_idx = g["avg_weight"].idxmax()

        cols = st.columns(len(g))
        for i, (_, row) in enumerate(g.iterrows()):
            label = f"{row['표시']}"
            value = f"{row['avg_weight']:.2f} g"
            if row.name == max_idx:
                cols[i].metric("🥇 최고 평균 생중량", value, help=label)
            else:
                cols[i].metric(label, value)

        st.info("참고: 프로젝트 정의상 ‘하늘고(EC 2.0)’를 최적 후보로 표시할 수 있으며, 실제 최댓값은 데이터에 따라 달라질 수 있습니다.")

    st.markdown("---")
    st.subheader("EC별 생육 비교")

    if growth_summary.empty:
        st.error("그래프를 그릴 데이터가 없습니다.")
    else:
        # 2x2 bar charts
        fig2 = make_subplots(
            rows=2,
            cols=2,
            subplot_titles=("평균 생중량(g) ⭐", "평균 잎 수(장)", "평균 지상부 길이(mm)", "개체수(n)"),
        )

        x = growth_summary["school"].astype(str)

        fig2.add_trace(go.Bar(x=x, y=growth_summary["avg_weight"], name="Avg Weight"), row=1, col=1)
        fig2.add_trace(go.Bar(x=x, y=growth_summary["avg_leaf"], name="Avg Leaves"), row=1, col=2)
        fig2.add_trace(go.Bar(x=x, y=growth_summary["avg_shoot"], name="Avg Shoot"), row=2, col=1)
        fig2.add_trace(go.Bar(x=x, y=growth_summary["n"], name="Count"), row=2, col=2)

        fig2.update_layout(height=700, showlegend=False)
        fig2 = plotly_apply_font(fig2)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    st.subheader("학교별 생중량 분포")

    if growth_view.empty:
        st.error("선택한 조건에서 분포를 표시할 생육 데이터가 없습니다.")
    else:
        # Box plot
        fig_box = px.box(
            growth_view,
            x="school",
            y="생중량(g)",
            points="all",
            title="Fresh Weight Distribution by School",
        )
        fig_box = plotly_apply_font(fig_box)
        st.plotly_chart(fig_box, use_container_width=True)

    st.markdown("---")
    st.subheader("상관관계 분석")

    if growth_view.empty:
        st.error("선택한 조건에서 상관관계를 표시할 생육 데이터가 없습니다.")
    else:
        c1, c2 = st.columns(2)

        fig_sc1 = px.scatter(
            growth_view,
            x="잎 수(장)",
            y="생중량(g)",
            color="school",
            title="Leaves vs Fresh Weight",
        )
        fig_sc1 = plotly_apply_font(fig_sc1)
        c1.plotly_chart(fig_sc1, use_container_width=True)

        fig_sc2 = px.scatter(
            growth_view,
            x="지상부 길이(mm)",
            y="생중량(g)",
            color="school",
            title="Shoot Length vs Fresh Weight",
        )
        fig_sc2 = plotly_apply_font(fig_sc2)
        c2.plotly_chart(fig_sc2, use_container_width=True)

    with st.expander("원본 생육 데이터 보기 / 다운로드"):
        if growth_view.empty:
            st.error("표시할 생육 데이터가 없습니다.")
        else:
            st.dataframe(growth_view, use_container_width=True)

            # XLSX download (selected)
            xlsx_bytes = dataframe_to_xlsx_bytes(growth_view, sheet_name="growth")
            st.download_button(
                label="XLSX 다운로드",
                data=xlsx_bytes,
                file_name="생육데이터_선택.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
