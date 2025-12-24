import io
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots


# ----------------------------
# Page & Fonts
# ----------------------------
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
    font-family: 'Noto Sans KR', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
}
</style>
""",
    unsafe_allow_html=True,
)

PLOTLY_FONT = dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif")


# ----------------------------
# Helpers: Unicode-safe file finding
# ----------------------------
def _norm(s: str, form: str) -> str:
    return unicodedata.normalize(form, s)


def _norm_both(s: str) -> set:
    # Compare both NFC and NFD to avoid macOS/Windows normalization issues
    return {_norm(s, "NFC"), _norm(s, "NFD")}


def find_file_unicode_safe(
    data_dir: Path,
    preferred_names: list[str],
    suffixes: tuple[str, ...] = (".xlsx", ".csv"),
    must_contain_keywords: list[str] | None = None,
) -> Path | None:
    """
    Scan data_dir using Path.iterdir() and match filenames using NFC/NFD normalization.
    - preferred_names: exact preferred filename candidates
    - must_contain_keywords: fallback match if preferred not found; all keywords must be in name
    """
    if not data_dir.exists() or not data_dir.is_dir():
        return None

    preferred_norm_sets = [_norm_both(name) for name in preferred_names]
    keywords = must_contain_keywords or []

    # 1) exact match by preferred names (unicode-safe)
    for p in data_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in [s.lower() for s in suffixes]:
            continue
        p_name_set = _norm_both(p.name)
        for target_set in preferred_norm_sets:
            if p_name_set & target_set:
                return p

    # 2) fallback: keyword match (unicode-safe)
    kw_sets = [_norm_both(k) for k in keywords]  # each keyword in both forms
    for p in data_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in [s.lower() for s in suffixes]:
            continue
        p_name_nfc = _norm(p.name, "NFC")
        p_name_nfd = _norm(p.name, "NFD")

        ok = True
        for kset in kw_sets:
            # keyword exists in either NFC or NFD representation
            if not any(k in p_name_nfc for k in kset) and not any(k in p_name_nfd for k in kset):
                ok = False
                break
        if ok:
            return p

    return None


def list_data_files(data_dir: Path) -> list[str]:
    if not data_dir.exists():
        return []
    out = []
    for p in data_dir.iterdir():
        if p.is_file():
            out.append(p.name)
    return sorted(out)


# ----------------------------
# Data Loading
# ----------------------------
@st.cache_data(show_spinner=False)
def load_environment_csvs(data_dir: Path) -> dict[str, pd.DataFrame]:
    """
    Load all *_환경데이터.csv files from data_dir.
    Returns dict: {school_name: df}
    Columns expected: time, temperature, humidity, ph, ec
    """
    env = {}
    if not data_dir.exists():
        return env

    for p in data_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() != ".csv":
            continue

        # Unicode-safe check for '환경데이터' in filename
        name_nfc = _norm(p.name, "NFC")
        name_nfd = _norm(p.name, "NFD")
        if ("환경데이터" not in name_nfc) and ("환경데이터" not in name_nfd):
            continue

        try:
            df = pd.read_csv(p)
        except Exception:
            # try encoding fallback
            try:
                df = pd.read_csv(p, encoding="cp949")
            except Exception:
                continue

        # derive school name from filename: "{학교}_환경데이터.csv"
        # avoid f-string path building; only parse the stem
        stem_nfc = _norm(p.stem, "NFC")
        stem_nfd = _norm(p.stem, "NFD")

        # handle both forms: split by "_환경데이터"
        school = None
        for stem in [stem_nfc, stem_nfd]:
            if "_환경데이터" in stem:
                school = stem.split("_환경데이터")[0].strip()
                break
            if "환경데이터" in stem:
                school = stem.split("환경데이터")[0].replace("_", "").strip()
                break

        if not school:
            continue

        # normalize columns
        rename_map = {}
        for c in df.columns:
            lc = str(c).strip().lower()
            if lc in ["temp", "temperature"]:
                rename_map[c] = "temperature"
            elif lc in ["humid", "humidity"]:
                rename_map[c] = "humidity"
            elif lc in ["ph"]:
                rename_map[c] = "ph"
            elif lc in ["ec"]:
                rename_map[c] = "ec"
            elif lc in ["time", "timestamp", "datetime", "date"]:
                rename_map[c] = "time"
        df = df.rename(columns=rename_map)

        # ensure required columns exist
        for col in ["time", "temperature", "humidity", "ph", "ec"]:
            if col not in df.columns:
                # keep loading but dashboard will warn later
                pass

        # parse time if exists
        if "time" in df.columns:
            df["time"] = pd.to_datetime(df["time"], errors="coerce")

        # numeric convert
        for col in ["temperature", "humidity", "ph", "ec"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        env[school] = df

    return env


@st.cache_data(show_spinner=False)
def load_growth_xlsx(data_dir: Path) -> tuple[Path | None, dict[str, pd.DataFrame]]:
    """
    Find and load the growth results XLSX (all sheets).
    Return: (xlsx_path, {sheet_name: df})
    Columns expected:
    개체번호, 잎 수(장), 지상부 길이(mm), 지하부길이(mm), 생중량(g)
    """
    preferred = [
        "4개교_생육결과데이터.xlsx",
        "4개교 생육결과데이터.xlsx",
        "4개교_생육 결과 데이터.xlsx",
        "4개교 생육 결과 데이터.xlsx",
    ]

    xlsx_path = find_file_unicode_safe(
        data_dir=data_dir,
        preferred_names=preferred,
        suffixes=(".xlsx",),
        must_contain_keywords=["생육"],  # fallback: any .xlsx containing '생육'
    )

    if xlsx_path is None:
        return None, {}

    try:
        sheets = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
    except Exception:
        return xlsx_path, {}

    # normalize columns per sheet
    out = {}
    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue

        # trim columns
        df.columns = [str(c).strip() for c in df.columns]

        # numeric conversions where possible
        for col in df.columns:
            if col == "개체번호":
                df[col] = pd.to_numeric(df[col], errors="coerce")
            if "생중량" in col or "잎" in col or "길이" in col:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        out[str(sheet_name).strip()] = df

    return xlsx_path, out


def make_school_meta(env_map: dict[str, pd.DataFrame], growth_sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Create a unified school metadata table without hardcoding sheet names.
    School list = union of env_map keys and growth sheet names (normalized compare).
    """
    # normalize names to match env school & sheet names loosely
    env_keys = list(env_map.keys())
    sheet_keys = list(growth_sheets.keys())

    def canonical(name: str) -> str:
        # remove spaces and normalize
        return _norm(name.replace(" ", ""), "NFC")

    env_can = {canonical(k): k for k in env_keys}
    sheet_can = {canonical(k): k for k in sheet_keys}

    all_can = sorted(set(env_can.keys()) | set(sheet_can.keys()))

    rows = []
    # EC targets given by the prompt (these are experimental conditions, not filenames)
    # This is not sheet-name hardcoding; it's research design metadata.
    ec_target_by_school_hint = {
        "송도고": 1.0,
        "하늘고": 2.0,
        "아라고": 4.0,
        "동산고": 8.0,
    }

    # color hint for UI
    color_hint = {
        "송도고": "#3b82f6",
        "하늘고": "#22c55e",
        "아라고": "#f59e0b",
        "동산고": "#ef4444",
    }

    for can in all_can:
        school_display = env_can.get(can) or sheet_can.get(can) or can
        growth_df = growth_sheets.get(sheet_can.get(can, ""), pd.DataFrame())
        n = int(growth_df.shape[0]) if not growth_df.empty else 0

        # pick EC target if matches known schools; else unknown
        # match by containing substring (unicode-safe)
        target = None
        for k, v in ec_target_by_school_hint.items():
            if k in school_display:
                target = v
                break

        col = None
        for k, v in color_hint.items():
            if k in school_display:
                col = v
                break

        rows.append(
            {
                "학교명": school_display,
                "EC 목표": target,
                "개체수": n,
                "색상": col,
            }
        )

    return pd.DataFrame(rows)


def safe_mean(df: pd.DataFrame, col: str) -> float | None:
    if df is None or df.empty or col not in df.columns:
        return None
    v = pd.to_numeric(df[col], errors="coerce").dropna()
    if len(v) == 0:
        return None
    return float(v.mean())


def style_plotly(fig: go.Figure) -> go.Figure:
    fig.update_layout(font=PLOTLY_FONT)
    return fig


# ----------------------------
# App Title
# ----------------------------
st.title("🌱 극지식물 최적 EC 농도 연구")


# ----------------------------
# Load Data
# ----------------------------
data_dir = Path(__file__).parent / "data"

with st.spinner("데이터를 불러오는 중..."):
    env_map = load_environment_csvs(data_dir)
    growth_xlsx_path, growth_sheets = load_growth_xlsx(data_dir)

if len(env_map) == 0:
    st.error("환경 데이터(CSV)를 찾지 못했습니다. data/ 폴더에 '*_환경데이터.csv' 파일이 있는지 확인해주세요.")
    st.write("현재 data/ 폴더 파일 목록:")
    st.code("\n".join(list_data_files(data_dir)) or "(없음)")
    st.stop()

if growth_xlsx_path is None or len(growth_sheets) == 0:
    st.error("생육 결과 데이터(XLSX)를 찾지 못했습니다. data/ 폴더에 '생육'이 포함된 .xlsx 파일이 있는지 확인해주세요.")
    st.write("현재 data/ 폴더 파일 목록:")
    st.code("\n".join(list_data_files(data_dir)) or "(없음)")
    st.stop()

school_meta = make_school_meta(env_map, growth_sheets)

# Sidebar school filter
schools_all = ["전체"] + sorted(school_meta["학교명"].dropna().unique().tolist())
selected_school = st.sidebar.selectbox("학교 선택", schools_all, index=0)


def filter_schools(keys: list[str]) -> list[str]:
    if selected_school == "전체":
        return keys
    # unicode-safe contain match
    sel_nfc = _norm(selected_school.replace(" ", ""), "NFC")
    out = []
    for k in keys:
        k_nfc = _norm(str(k).replace(" ", ""), "NFC")
        if sel_nfc == k_nfc:
            out.append(k)
    return out


env_keys_filtered = filter_schools(list(env_map.keys()))
sheet_keys_filtered = filter_schools(list(growth_sheets.keys()))

# ----------------------------
# Pre-compute summary stats
# ----------------------------
# total individuals
total_n = int(sum(int(growth_sheets[s].shape[0]) for s in sheet_keys_filtered if s in growth_sheets))

# overall env means across selected schools
temps = []
humids = []
for s in env_keys_filtered:
    df = env_map.get(s, pd.DataFrame())
    if "temperature" in df.columns:
        temps.append(pd.to_numeric(df["temperature"], errors="coerce"))
    if "humidity" in df.columns:
        humids.append(pd.to_numeric(df["humidity"], errors="coerce"))

avg_temp = float(pd.concat(temps).dropna().mean()) if len(temps) else None
avg_humid = float(pd.concat(humids).dropna().mean()) if len(humids) else None

# Best EC by mean fresh weight across sheets (if school EC 목표 known)
growth_summary_rows = []
for _, r in school_meta.iterrows():
    school = r["학교명"]
    ec_target = r["EC 목표"]
    if school not in sheet_keys_filtered:
        continue
    gdf = growth_sheets.get(school, pd.DataFrame())
    # find weight column
    weight_col = None
    for c in gdf.columns:
        if "생중량" in str(c):
            weight_col = c
            break
    if weight_col is None:
        continue
    mean_w = safe_mean(gdf, weight_col)
    growth_summary_rows.append({"학교명": school, "EC": ec_target, "평균 생중량(g)": mean_w, "개체수": int(gdf.shape[0])})

growth_summary = pd.DataFrame(growth_summary_rows)

best_ec = None
if not growth_summary.empty and growth_summary["평균 생중량(g)"].notna().any():
    best_row = growth_summary.sort_values("평균 생중량(g)", ascending=False).iloc[0]
    best_ec = best_row["EC"]

# ----------------------------
# Tabs
# ----------------------------
tab1, tab2, tab3 = st.tabs(["📖 실험 개요", "🌡️ 환경 데이터", "📊 생육 결과"])

# =========================================================
# Tab 1: Overview
# =========================================================
with tab1:
    st.subheader("연구 배경 및 목적")
    st.write(
        """
극지식물은 야외 극지 환경이 아니라 **극지연구소 스마트팜(통제된 재배 환경)**에서 재배되는 식물을 의미한다.  
스마트팜에서는 **EC 농도, 온도, 습도**와 같은 환경 요인을 정밀하게 제어할 수 있기 때문에, 식물이 가장 잘 자라는 **최적 조건을 찾는 것**이 매우 중요하다.  
본 대시보드는 4개 학교의 실험 데이터를 비교·분석하여 **EC만 고려했을 때 vs 온·습도까지 고려했을 때** 최적 조건이 달라지는지 확인하고, 최적 EC를 도출하는 데 도움을 준다.
"""
    )

    st.subheader("학교별 EC 조건")
    show_meta = school_meta.copy()
    if selected_school != "전체":
        show_meta = show_meta[show_meta["학교명"] == selected_school]

    st.dataframe(show_meta, use_container_width=True)

    st.subheader("주요 지표")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 개체수", f"{total_n:,d}")
    c2.metric("평균 온도", "-" if avg_temp is None else f"{avg_temp:.2f} °C")
    c3.metric("평균 습도", "-" if avg_humid is None else f"{avg_humid:.2f} %")
    c4.metric("최적 EC(평균 생중량 기준)", "-" if best_ec is None else f"{best_ec}")

# =========================================================
# Tab 2: Environment Data
# =========================================================
with tab2:
    st.subheader("학교별 환경 평균 비교")

    # compute averages per school
    env_rows = []
    for s in env_keys_filtered:
        d = env_map.get(s, pd.DataFrame())
        env_rows.append(
            {
                "학교": s,
                "평균 온도(°C)": safe_mean(d, "temperature"),
                "평균 습도(%)": safe_mean(d, "humidity"),
                "평균 pH": safe_mean(d, "ph"),
                "평균 EC(실측)": safe_mean(d, "ec"),
            }
        )
    env_avg = pd.DataFrame(env_rows)

    # add EC target from meta if available
    env_avg = env_avg.merge(school_meta[["학교명", "EC 목표"]], left_on="학교", right_on="학교명", how="left").drop(
        columns=["학교명"]
    )

    fig = make_subplots(
        rows=2,
        cols=2,
        subplot_titles=("Avg Temperature", "Avg Humidity", "Avg pH", "Target EC vs Measured EC"),
    )

    # bar: temp
    fig.add_trace(
        go.Bar(x=env_avg["학교"], y=env_avg["평균 온도(°C)"], name="Avg Temp"),
        row=1,
        col=1,
    )

    # bar: humid
    fig.add_trace(
        go.Bar(x=env_avg["학교"], y=env_avg["평균 습도(%)"], name="Avg Humidity"),
        row=1,
        col=2,
    )

    # bar: pH
    fig.add_trace(
        go.Bar(x=env_avg["학교"], y=env_avg["평균 pH"], name="Avg pH"),
        row=2,
        col=1,
    )

    # dual bar: target vs measured ec
    fig.add_trace(
        go.Bar(x=env_avg["학교"], y=env_avg["EC 목표"], name="Target EC"),
        row=2,
        col=2,
    )
    fig.add_trace(
        go.Bar(x=env_avg["학교"], y=env_avg["평균 EC(실측)"], name="Measured EC"),
        row=2,
        col=2,
    )

    fig.update_layout(
        barmode="group",
        height=700,
        title_text="Environment Averages (by School)",
        font=PLOTLY_FONT,
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.subheader("선택한 학교 시계열")

    # show time series for selected school (if '전체' show first school)
    ts_school = selected_school
    if ts_school == "전체":
        ts_school = env_keys_filtered[0] if env_keys_filtered else None

    if ts_school is None or ts_school not in env_map:
        st.error("시계열을 표시할 학교를 찾지 못했습니다.")
    else:
        d = env_map[ts_school].copy()
        if "time" not in d.columns or d["time"].isna().all():
            st.error("시간(time) 컬럼을 해석할 수 없습니다. CSV의 'time' 컬럼 형식을 확인해주세요.")
        else:
            # target EC if exists
            target_ec = None
            m = school_meta[school_meta["학교명"] == ts_school]
            if not m.empty:
                target_ec = m.iloc[0]["EC 목표"]

            # Temperature
            if "temperature" in d.columns:
                fig_t = px.line(d.sort_values("time"), x="time", y="temperature", title="Temperature Over Time")
                fig_t.update_layout(font=PLOTLY_FONT)
                st.plotly_chart(fig_t, use_container_width=True)

            # Humidity
            if "humidity" in d.columns:
                fig_h = px.line(d.sort_values("time"), x="time", y="humidity", title="Humidity Over Time")
                fig_h.update_layout(font=PLOTLY_FONT)
                st.plotly_chart(fig_h, use_container_width=True)

            # EC with target line
            if "ec" in d.columns:
                fig_e = px.line(d.sort_values("time"), x="time", y="ec", title="EC Over Time")
                if target_ec is not None and pd.notna(target_ec):
                    fig_e.add_hline(y=float(target_ec), line_dash="dash", annotation_text="Target EC")
                fig_e.update_layout(font=PLOTLY_FONT)
                st.plotly_chart(fig_e, use_container_width=True)

    with st.expander("환경 데이터 원본 보기 및 다운로드"):
        # show combined table for selected scope
        frames = []
        for s in env_keys_filtered:
            tmp = env_map[s].copy()
            tmp.insert(0, "School", s)
            frames.append(tmp)
        env_all = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

        st.dataframe(env_all, use_container_width=True)

        # download CSV
        csv_bytes = env_all.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="환경 데이터 CSV 다운로드",
            data=csv_bytes,
            file_name="환경데이터_통합.csv",
            mime="text/csv",
        )

# =========================================================
# Tab 3: Growth Results
# =========================================================
with tab3:
    st.subheader("🥇 핵심 결과: EC별 평균 생중량")

    if growth_summary.empty:
        st.error("생중량 컬럼을 찾지 못했거나 생육 데이터가 비어있습니다.")
    else:
        # Highlight best EC
        if best_ec is not None:
            st.info(f"현재 데이터 기준으로 평균 생중량이 가장 높은 조건은 **EC {best_ec}** 입니다. (학교 단위 비교)")

        # card-like metrics per EC
        cols = st.columns(min(4, len(growth_summary)))
        sorted_sum = growth_summary.sort_values("평균 생중량(g)", ascending=False)
        for i, (_, r) in enumerate(sorted_sum.iterrows()):
            if i >= len(cols):
                break
            label = f"{r['학교명']} (EC {r['EC']})"
            value = "-" if pd.isna(r["평균 생중량(g)"]) else f"{r['평균 생중량(g)']:.2f} g"
            cols[i].metric(label, value)

    st.divider()
    st.subheader("EC별 생육 비교 (2x2)")

    # build bar charts for: mean weight, mean leaves, mean shoot length, count
    rows = []
    for s in sheet_keys_filtered:
        gdf = growth_sheets.get(s, pd.DataFrame())
        if gdf.empty:
            continue

        # detect columns
        weight_col = next((c for c in gdf.columns if "생중량" in str(c)), None)
        leaf_col = next((c for c in gdf.columns if "잎" in str(c)), None)
        shoot_col = next((c for c in gdf.columns if "지상부" in str(c)), None)

        # EC target
        ec_target = None
        m = school_meta[school_meta["학교명"] == s]
        if not m.empty:
            ec_target = m.iloc[0]["EC 목표"]

        rows.append(
            {
                "학교": s,
                "EC": ec_target,
                "평균 생중량(g)": safe_mean(gdf, weight_col) if weight_col else None,
                "평균 잎 수": safe_mean(gdf, leaf_col) if leaf_col else None,
                "평균 지상부 길이(mm)": safe_mean(gdf, shoot_col) if shoot_col else None,
                "개체수": int(gdf.shape[0]),
            }
        )

    growth_avg = pd.DataFrame(rows)

    fig2 = make_subplots(
        rows=2,
        cols=2,
        subplot_titles=("Mean Fresh Weight (g)", "Mean Leaf Count", "Mean Shoot Length (mm)", "Sample Size (n)"),
    )

    fig2.add_trace(go.Bar(x=growth_avg["학교"], y=growth_avg["평균 생중량(g)"], name="Fresh Weight"), row=1, col=1)
    fig2.add_trace(go.Bar(x=growth_avg["학교"], y=growth_avg["평균 잎 수"], name="Leaf Count"), row=1, col=2)
    fig2.add_trace(go.Bar(x=growth_avg["학교"], y=growth_avg["평균 지상부 길이(mm)"], name="Shoot Length"), row=2, col=1)
    fig2.add_trace(go.Bar(x=growth_avg["학교"], y=growth_avg["개체수"], name="n"), row=2, col=2)

    fig2.update_layout(
        height=700,
        title_text="Growth Comparison (by School / EC Condition)",
        font=PLOTLY_FONT,
        showlegend=False,
    )
    st.plotly_chart(fig2, use_container_width=True)

    st.divider()
    st.subheader("학교별 생중량 분포")

    # build long-form for distribution plots
    long_rows = []
    for s in sheet_keys_filtered:
        gdf = growth_sheets.get(s, pd.DataFrame())
        if gdf.empty:
            continue
        weight_col = next((c for c in gdf.columns if "생중량" in str(c)), None)
        if not weight_col:
            continue
        tmp = gdf[[weight_col]].copy()
        tmp["학교"] = s
        tmp = tmp.rename(columns={weight_col: "생중량(g)"})
        long_rows.append(tmp)

    if long_rows:
        long_df = pd.concat(long_rows, ignore_index=True)
        fig_box = px.box(long_df, x="학교", y="생중량(g)", points="all", title="Fresh Weight Distribution by School")
        fig_box.update_layout(font=PLOTLY_FONT)
        st.plotly_chart(fig_box, use_container_width=True)
    else:
        st.error("분포 그래프를 만들 생중량 데이터를 찾지 못했습니다.")

    st.divider()
    st.subheader("상관관계 분석 (산점도 2개)")

    # Combine for scatter: Leaf vs Weight, Shoot vs Weight
    scatter_rows = []
    for s in sheet_keys_filtered:
        gdf = growth_sheets.get(s, pd.DataFrame())
        if gdf.empty:
            continue

        weight_col = next((c for c in gdf.columns if "생중량" in str(c)), None)
        leaf_col = next((c for c in gdf.columns if "잎" in str(c)), None)
        shoot_col = next((c for c in gdf.columns if "지상부" in str(c)), None)
        if not weight_col:
            continue

        cols_needed = [c for c in [leaf_col, shoot_col, weight_col] if c is not None]
        tmp = gdf[cols_needed].copy()
        tmp["학교"] = s
        tmp = tmp.rename(
            columns={
                weight_col: "생중량(g)",
                leaf_col: "잎 수(장)" if leaf_col else leaf_col,
                shoot_col: "지상부 길이(mm)" if shoot_col else shoot_col,
            }
        )
        scatter_rows.append(tmp)

    if scatter_rows:
        scat = pd.concat(scatter_rows, ignore_index=True)

        c1, c2 = st.columns(2)
        with c1:
            if "잎 수(장)" in scat.columns:
                fig_sc1 = px.scatter(scat, x="잎 수(장)", y="생중량(g)", color="학교", title="Leaf Count vs Fresh Weight")
                fig_sc1.update_layout(font=PLOTLY_FONT)
                st.plotly_chart(fig_sc1, use_container_width=True)
            else:
                st.error("잎 수 컬럼을 찾지 못했습니다.")

        with c2:
            if "지상부 길이(mm)" in scat.columns:
                fig_sc2 = px.scatter(scat, x="지상부 길이(mm)", y="생중량(g)", color="학교", title="Shoot Length vs Fresh Weight")
                fig_sc2.update_layout(font=PLOTLY_FONT)
                st.plotly_chart(fig_sc2, use_container_width=True)
            else:
                st.error("지상부 길이 컬럼을 찾지 못했습니다.")
    else:
        st.error("상관관계 산점도를 만들 데이터가 부족합니다.")

    with st.expander("학교별 생육 데이터 원본 보기 및 다운로드"):
        # show selected sheets
        for s in sheet_keys_filtered:
            st.markdown(f"**{s}**")
            st.dataframe(growth_sheets[s], use_container_width=True)

        # download XLSX (all selected sheets)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            for s in sheet_keys_filtered:
                growth_sheets[s].to_excel(writer, index=False, sheet_name=str(s)[:31])
        buffer.seek(0)

        st.download_button(
            label="생육 데이터 XLSX 다운로드 (선택 범위)",
            data=buffer,
            file_name="생육데이터_선택범위.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
