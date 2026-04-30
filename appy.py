import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ─────────────────────────────────────────
# MERCK COLOR PALETTE
# ─────────────────────────────────────────
MERCK_TEAL       = "#00857C"
MERCK_TEAL_LIGHT = "#6ECEB2"
MERCK_BLUE       = "#0C2340"
MERCK_BLUE_MID   = "#005587"
MERCK_GRAY       = "#C1C6C8"
MERCK_ORANGE     = "#E37222"
MERCK_RED        = "#BF3030"

st.set_page_config(page_title="Submission Dashboard", layout="wide", page_icon="📋")

# ─────────────────────────────────────────
# GLOBAL CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer     {visibility: hidden;}
    header     {visibility: hidden;}

    /* ── Sidebar ── */
    section[data-testid="stSidebar"] {
        background-color: #0C2340;
        border-right: 3px solid #00857C;
    }
    section[data-testid="stSidebar"] * { color: #FFFFFF !important; }

    /* ── Page header banner ── */
    .merck-header {
        background: linear-gradient(90deg, #0C2340 0%, #005587 55%, #00857C 100%);
        color: white;
        padding: 18px 28px;
        border-radius: 8px;
        margin-bottom: 20px;
        font-size: 1.35rem;
        font-weight: 700;
        letter-spacing: 0.03em;
        border-left: 5px solid #6ECEB2;
        box-shadow: 0 2px 8px rgba(12,35,64,0.18);
    }
    .merck-header small {
        font-size: 0.70rem;
        font-weight: 400;
        opacity: 0.80;
        display: block;
        margin-top: 4px;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }

    /* ── Section sub-label ── */
    .section-label {
        background: #f0f6f6;
        border-left: 4px solid #00857C;
        padding: 8px 14px;
        border-radius: 0 6px 6px 0;
        color: #0C2340;
        font-weight: 600;
        font-size: 0.93rem;
        margin: 18px 0 12px 0;
    }

    /* ── Metric cards ── */
    div[data-testid="metric-container"] {
        background: #f4f8f8;
        border-left: 4px solid #00857C;
        border-radius: 6px;
        padding: 14px 16px;
        box-shadow: 0 1px 4px rgba(12,35,64,0.07);
    }
    div[data-testid="metric-container"] label {
        color: #005587 !important;
        font-size: 0.73rem !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    div[data-testid="metric-container"] div[data-testid="metric-value"] {
        color: #0C2340 !important;
        font-size: 1.60rem !important;
        font-weight: 700 !important;
    }

    /* ── Buttons ── */
    .stButton>button, .stDownloadButton>button {
        background-color: #00857C !important;
        color: white !important;
        border: none !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #005587 !important;
    }

    /* ── Tabs styling ── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: #f0f6f6;
        padding: 6px 8px;
        border-radius: 8px;
        margin-bottom: 16px;
    }
    .stTabs [data-baseweb="tab"] {
        background: white;
        border-radius: 6px;
        padding: 8px 20px;
        font-weight: 600;
        color: #005587;
        border: 1px solid #C1C6C8;
    }
    .stTabs [aria-selected="true"] {
        background: #0C2340 !important;
        color: white !important;
        border-color: #0C2340 !important;
    }

    /* ── Nav labels in sidebar ── */
    .nav-section {
        font-size: 0.60rem;
        text-transform: uppercase;
        letter-spacing: 0.10em;
        color: #6ECEB2 !important;
        margin-top: 10px;
        margin-bottom: 2px;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────
for key, val in {
    "rolling_data":    None,
    "nonrolling_data": None,
    "anchor_dates":    pd.DataFrame(columns=["Anchor Date", "Date", "Status"]),
}.items():
    if key not in st.session_state:
        st.session_state[key] = val


# ─────────────────────────────────────────
# SHARED HELPERS
# ─────────────────────────────────────────
def page_header(title, subtitle=""):
    sub = f"<small>{subtitle}</small>" if subtitle else ""
    st.markdown(f'<div class="merck-header">{title}{sub}</div>', unsafe_allow_html=True)

def section_label(text):
    st.markdown(f'<div class="section-label">{text}</div>', unsafe_allow_html=True)

def std_chart(fig):
    """Apply standard Merck chart styling."""
    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(color=MERCK_BLUE, size=11),
        title_font=dict(color=MERCK_BLUE, size=13),
        legend=dict(bgcolor="rgba(255,255,255,0.9)",
                    bordercolor=MERCK_GRAY, borderwidth=1),
    )
    fig.update_xaxes(showgrid=False, linecolor=MERCK_GRAY)
    fig.update_yaxes(gridcolor="#E8ECEC", linecolor=MERCK_GRAY)
    return fig


# ─────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────
def read_submission_excel(uploaded_file, sheet_name):
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    header_row = None
    for i, row in raw.iterrows():
        row_text = row.astype(str).str.strip().str.lower().tolist()
        if "task name" in row_text and (
            "planned start" in row_text or "planned finish" in row_text
        ):
            header_row = i
            break

    if header_row is not None:
        headers = raw.iloc[header_row].fillna("").astype(str).str.strip()
        clean = []
        for idx, h in enumerate(headers):
            clean.append(h if h and h.lower() != "nan" else f"blank_col_{idx}")
        df = raw.iloc[header_row + 1:].copy()
        df.columns = clean
    else:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    df.columns = df.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    df = df.dropna(how="all")

    rename_map = {}
    for col in df.columns:
        c = col.strip().lower()
        if   c == "actual start":       rename_map[col] = "Actual Start"
        elif c == "actual finish":      rename_map[col] = "Actual Finish"
        elif c == "planned start":      rename_map[col] = "Planned Start"
        elif c == "planned finish":     rename_map[col] = "Planned Finish"
        elif c == "task name":          rename_map[col] = "Task Name"
        elif c == "component id":       rename_map[col] = "Component ID"
        elif c == "component source":   rename_map[col] = "Component Source"
        elif c == "filing status":      rename_map[col] = "Filing Status"
        elif c == "task index":         rename_map[col] = "Task Index"
        elif c == "wave":               rename_map[col] = "Wave"
        elif c == "module":             rename_map[col] = "Module"
    df = df.rename(columns=rename_map)

    if "Wave" not in df.columns:
        for col in df.columns:
            if df[col].astype(str).str.contains("Rolling Submission|Wave", case=False, na=False).any():
                df["Wave"] = df[col].ffill()
                break

    if "Module" not in df.columns:
        for col in df.columns:
            if df[col].astype(str).str.contains("Module", case=False, na=False).any():
                df["Module"] = df[col].ffill()
                break
    return df


def clean_submission_data(df):
    df.columns = df.columns.astype(str).str.strip()
    required = ["Task Name", "Planned Start", "Actual Start", "Planned Finish", "Actual Finish"]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {missing}")
        st.write("Found:", list(df.columns))
        st.stop()

    if "Component ID" in df.columns:
        df = df[df["Component ID"].notna() & df["Task Name"].notna()]
    else:
        df = df[df["Task Name"].notna()]

    for col in ["Planned Start", "Actual Start", "Planned Finish", "Actual Finish"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    if "Filing Status" in df.columns:
        df["Filing Status"] = df["Filing Status"].fillna("Incomplete")
    else:
        df["Filing Status"] = df["Actual Finish"].apply(
            lambda x: "Completed" if pd.notna(x) else "Incomplete"
        )

    for col_flag, fallback in [("Actually Completed", "Actual Finish"),
                                ("Planned Completed",  "Planned Finish")]:
        if col_flag in df.columns:
            df[col_flag] = (
                df[col_flag].astype(str).str.strip().str.lower()
                .map({"true": True, "false": False})
            )
            df[col_flag] = df[col_flag].fillna(df[fallback].notna())
        else:
            df[col_flag] = df[fallback].notna()

    df["StartVarianceDays"]  = (df["Actual Start"]  - df["Planned Start"]).dt.days
    df["FinishVarianceDays"] = (df["Actual Finish"] - df["Planned Finish"]).dt.days

    if "Wave"   not in df.columns: df["Wave"]   = "No Wave"
    if "Module" not in df.columns: df["Module"] = "No Module"
    return df


def calculate_metrics(df):
    total     = len(df)
    completed = df["Actually Completed"].eq(True).sum()
    remaining = df["Actually Completed"].eq(False).sum()
    planned   = df["Planned Completed"].eq(True).sum()
    rate      = completed / total if total > 0 else 0
    variance  = df["FinishVarianceDays"].sum(skipna=True) if "FinishVarianceDays" in df.columns else 0
    return total, completed, remaining, planned, rate, variance


# ─────────────────────────────────────────
# GANTT BUILDER  (definitive vline fix)
# ─────────────────────────────────────────
def build_gantt(df, group_col="Wave", max_rows=50):
    """
    Paired Planned/Actual bars on separate y-rows.

    add_vline FIX: Plotly's timeline x-axis is milliseconds since epoch.
    We must pass x as an integer (ms) — NOT a string, NOT a pd.Timestamp.
    Passing a string causes the 'int + str' TypeError seen in newer plotly builds.
    """
    gdf = df[df["Planned Start"].notna() & df["Planned Finish"].notna()].copy()
    if gdf.empty:
        return None
    gdf = gdf.head(max_rows).reset_index(drop=True)

    rows = []
    for i, (_, r) in enumerate(gdf.iterrows()):
        name   = str(r["Task Name"])[:55]
        status = str(r.get("Filing Status", ""))
        grp    = str(r.get(group_col, ""))

        rows.append({
            "y_key":  f"{i:04d}_P",
            "Label":  name,
            "Type":   "Planned",
            "Start":  r["Planned Start"],
            "Finish": r["Planned Finish"],
            "Hover":  (f"<b>{name}</b><br>"
                       f"Planned: {r['Planned Start'].strftime('%d %b %Y')} → "
                       f"{r['Planned Finish'].strftime('%d %b %Y')}<br>"
                       f"Status: {status} | {group_col}: {grp}"),
        })

        if pd.notna(r.get("Actual Start")) and pd.notna(r.get("Actual Finish")):
            rows.append({
                "y_key":  f"{i:04d}_A",
                "Label":  name,
                "Type":   "Actual",
                "Start":  r["Actual Start"],
                "Finish": r["Actual Finish"],
                "Hover":  (f"<b>{name}</b><br>"
                           f"Actual: {r['Actual Start'].strftime('%d %b %Y')} → "
                           f"{r['Actual Finish'].strftime('%d %b %Y')}<br>"
                           f"Status: {status} | {group_col}: {grp}"),
            })

    if not rows:
        return None

    plot_df = pd.DataFrame(rows).sort_values("y_key").reset_index(drop=True)
    fig = px.timeline(
        plot_df,
        x_start="Start", x_end="Finish",
        y="y_key", color="Type",
        color_discrete_map={"Planned": MERCK_BLUE_MID, "Actual": MERCK_TEAL_LIGHT},
        custom_data=["Hover"],
    )
    fig.update_traces(hovertemplate="%{customdata[0]}<extra></extra>")

    tick_vals  = plot_df["y_key"].tolist()
    tick_texts = [r["Label"] if r["Type"] == "Planned" else "  ↳ actual"
                  for _, r in plot_df.iterrows()]

    # ── DEFINITIVE FIX ──────────────────────────────────────────────────────
    # px.timeline sets x-axis type to "date", so add_vline must receive the
    # x value as an integer number of milliseconds since the Unix epoch.
    # Using a string triggers "unsupported operand type(s) for +: 'int' and 'str'".
    today_ms = int(pd.Timestamp(datetime.now().date()).timestamp() * 1000)
    fig.add_vline(
        x=today_ms,
        line=dict(color=MERCK_RED, width=1.5, dash="dash"),
        annotation_text="Today",
        annotation_font_color=MERCK_RED,
        annotation_position="top right",
    )
    # ────────────────────────────────────────────────────────────────────────

    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(color=MERCK_BLUE, size=10),
        legend=dict(title="Schedule Type", orientation="h",
                    yanchor="bottom", y=1.01, xanchor="right", x=1,
                    bgcolor="rgba(255,255,255,0.85)",
                    bordercolor=MERCK_GRAY, borderwidth=1),
        height=max(400, len(plot_df) * 22 + 160),
        xaxis=dict(showgrid=True, gridcolor="#E8ECEC", tickformat="%b %Y", title="Date"),
        yaxis=dict(autorange="reversed", showgrid=False,
                   tickmode="array", tickvals=tick_vals,
                   ticktext=tick_texts, tickfont=dict(size=10)),
        margin=dict(l=10, r=30, t=70, b=30),
        bargap=0.15,
        hoverlabel=dict(bgcolor="white", bordercolor=MERCK_GRAY, font_size=12),
    )
    return fig


# ─────────────────────────────────────────
# WAVE / MODULE SUMMARIES
# ─────────────────────────────────────────
def get_wave_summary(df):
    ws = df.groupby("Wave").agg(
        Total=("Task Name", "count"),
        Completed=("Actually Completed", lambda x: x.eq(True).sum()),
        Remaining=("Actually Completed", lambda x: x.eq(False).sum()),
        Planned=("Planned Completed",   lambda x: x.eq(True).sum()),
    ).reset_index()
    ws["Rate_%"]    = ws["Completed"] / ws["Total"] * 100
    ws["Done_%"]    = (ws["Completed"] / ws["Planned"] * 100).fillna(0).clip(upper=100)
    ws["Left_%"]    = 100 - ws["Done_%"]
    return ws


def compute_module_group(cid):
    s = str(cid).strip() if pd.notna(cid) else ""
    d = s.find(".")
    return s[:d] if d > 0 else s

def compute_module_sort(cid):
    mg = compute_module_group(cid)
    try:   return int(mg) if mg else None
    except ValueError: return None

def get_nonrolling_summary(df):
    work = df.copy()
    work["IsComplete"] = (
        work["Filing Status"].astype(str).str.strip().str.lower() == "completed"
    ).astype(int)
    if "Component ID" in work.columns:
        work["MG"] = work["Component ID"].apply(compute_module_group)
        work["MS"] = work["Component ID"].apply(compute_module_sort)
    else:
        work["MG"] = work.get("Module", pd.Series("", index=work.index))
        work["MS"] = None
    work = work[work["MG"].str.strip() != ""]
    s = work.groupby("MG").agg(
        Sort=("MS", "first"),
        Total=("Task Name", "count"),
        Completed=("IsComplete", "sum"),
    ).reset_index().rename(columns={"MG": "Module Group"})
    s["Remaining"]      = s["Total"] - s["Completed"]
    s["Pct_Complete"]   = (s["Completed"] / s["Total"]).fillna(0)
    s["Pct_Remaining"]  = 1 - s["Pct_Complete"]
    return s.sort_values("Sort", na_position="last").reset_index(drop=True)


# ─────────────────────────────────────────
# SHARED CHART BLOCKS
# ─────────────────────────────────────────
def render_gauge(rate, title="Overall Completion Rate"):
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=rate * 100,
        number={"suffix": "%", "valueformat": ".1f", "font": {"size": 38, "color": MERCK_BLUE}},
        delta={"reference": 80, "valueformat": ".1f", "suffix": "%"},
        title={"text": title, "font": {"color": MERCK_BLUE, "size": 13}},
        gauge={
            "axis": {"range": [0, 100], "tickcolor": MERCK_GRAY},
            "bar":  {"color": MERCK_TEAL},
            "steps": [
                {"range": [0,  50], "color": "#F5E8E8"},
                {"range": [50, 80], "color": "#EAF4F3"},
                {"range": [80, 100],"color": "#D0EDE9"},
            ],
            "threshold": {"line": {"color": MERCK_BLUE, "width": 2},
                          "thickness": 0.80, "value": 80},
        }
    ))
    fig.update_layout(paper_bgcolor="white", font=dict(color=MERCK_BLUE),
                      height=270, margin=dict(t=40, b=10, l=20, r=20))
    return fig


def render_status_donut(df):
    if "Filing Status" not in df.columns:
        return None
    sdf = df["Filing Status"].fillna("Unknown").value_counts().reset_index()
    sdf.columns = ["Status", "Count"]
    cmap = {"Completed": MERCK_TEAL, "Incomplete": MERCK_ORANGE,
            "Unknown": MERCK_GRAY, "In Progress": MERCK_BLUE_MID}
    fig = px.pie(sdf, names="Status", values="Count", hole=0.50,
                 title="Document Status Breakdown",
                 color="Status", color_discrete_map=cmap)
    fig.update_traces(textinfo="label+percent", textfont_size=11)
    fig.update_layout(paper_bgcolor="white", font=dict(color=MERCK_BLUE),
                      height=270, margin=dict(t=40, b=10))
    return fig


def render_variance_bar(df):
    if "FinishVarianceDays" not in df.columns:
        return None
    vdf = df[df["FinishVarianceDays"].notna()].copy()
    if vdf.empty:
        return None
    vdf["Cat"] = vdf["FinishVarianceDays"].apply(
        lambda v: "On Time / Early" if v <= 0 else ("1–7 Days Late" if v <= 7 else "7+ Days Late")
    )
    cc = vdf["Cat"].value_counts().reset_index()
    cc.columns = ["Category", "Count"]
    fig = px.bar(cc, x="Category", y="Count", color="Category", text="Count",
                 color_discrete_map={"On Time / Early": MERCK_TEAL,
                                     "1–7 Days Late":   MERCK_ORANGE,
                                     "7+ Days Late":    MERCK_RED})
    fig.update_traces(textposition="outside")
    fig.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                      font=dict(color=MERCK_BLUE), showlegend=False,
                      title="Task Finish Variance Breakdown")
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#E8ECEC")
    return fig


def render_gantt_section(df, group_col, tab_key=""):
    # tab_key namespaces widget keys so rolling/non-rolling tabs never collide (fixes DuplicateElementId)
    _k = f"{tab_key}_{group_col}" if tab_key else group_col
    fc1, fc2, fc3 = st.columns([2, 2, 1])
    with fc1:
        opts = sorted(df[group_col].dropna().unique().tolist()) if group_col in df.columns else []
        choices = ["All"] + opts
        sel = st.selectbox(f"Filter by {group_col}", choices, key=f"gantt_grp_{_k}")
        gdf = df if sel == "All" else df[df[group_col] == sel]

    with fc2:
        if "Filing Status" in df.columns:
            sts = ["All"] + sorted(df["Filing Status"].dropna().unique().tolist())
            sel_s = st.selectbox("Filter by Filing Status", sts, key=f"gantt_status_{_k}")
            if sel_s != "All":
                gdf = gdf[gdf["Filing Status"] == sel_s]

    with fc3:
        max_rows = st.slider("Max tasks", 10, 200, 50, step=10, key=f"gantt_rows_{_k}")

    st.caption("🟦 Blue = Planned  |  🟩 Teal = Actual  |  🔴 Dashed = Today")

    fig = build_gantt(gdf, group_col=group_col, max_rows=max_rows)
    if fig is None:
        st.warning("No tasks with valid Planned Start and Finish dates found.")
    else:
        st.plotly_chart(fig, use_container_width=True, key=f"gantt_main_{_k}")

    fig_v = render_variance_bar(gdf)
    if fig_v:
        st.markdown("---")
        section_label("Schedule Variance Summary")
        st.plotly_chart(fig_v, use_container_width=True, key=f"gantt_var_{_k}")


# ─────────────────────────────────────────
# SIDEBAR  (minimal — just anchor dates nav)
# ─────────────────────────────────────────
with st.sidebar:
    # ── Merck wordmark (text-based, no broken image) ──────────────────────
    st.markdown("""
    <div style="padding:20px 0 4px 0;">
        <div style="font-size:1.55rem;font-weight:900;color:#6ECEB2;
                    letter-spacing:0.10em;text-transform:uppercase;
                    line-height:1;">MERCK</div>
        <div style="font-size:0.58rem;color:rgba(255,255,255,0.45);
                    letter-spacing:0.18em;text-transform:uppercase;
                    margin-top:2px;">Regulatory Affairs</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr style='border-color:rgba(110,206,178,0.3);margin:10px 0 16px 0;'>",
                unsafe_allow_html=True)

    # ── Navigation ────────────────────────────────────────────────────────
    st.markdown("<div style='font-size:0.60rem;letter-spacing:0.12em;color:#6ECEB2;"
                "text-transform:uppercase;margin-bottom:6px;'>Navigation</div>",
                unsafe_allow_html=True)
    page = st.radio(
        "",
        ["📋 Submission Dashboard", "📌 Anchor Dates"],
        label_visibility="collapsed",
    )

    st.markdown("<hr style='border-color:rgba(110,206,178,0.3);margin:16px 0 14px 0;'>",
                unsafe_allow_html=True)

    # ── Data status ───────────────────────────────────────────────────────
    r_loaded  = st.session_state.rolling_data is not None
    nr_loaded = st.session_state.nonrolling_data is not None

    st.markdown("<div style='font-size:0.60rem;letter-spacing:0.12em;color:#6ECEB2;"
                "text-transform:uppercase;margin-bottom:8px;'>Data Status</div>",
                unsafe_allow_html=True)

    def _status_pill(label, loaded, df):
        dot   = "#00857C" if loaded else "rgba(255,255,255,0.2)"
        count = f"<span style='color:rgba(255,255,255,0.45);font-size:0.70rem;'>&nbsp;·&nbsp;{len(df):,} rows</span>" if loaded else ""
        return (f"<div style='display:flex;align-items:center;gap:8px;"
                f"margin-bottom:8px;'>"
                f"<div style='width:8px;height:8px;border-radius:50%;background:{dot};flex-shrink:0;'></div>"
                f"<span style='font-size:0.78rem;font-weight:600;'>{label}</span>{count}</div>")

    st.markdown(
        _status_pill("Rolling", r_loaded, st.session_state.rolling_data or []) +
        _status_pill("Non-Rolling", nr_loaded, st.session_state.nonrolling_data or []),
        unsafe_allow_html=True,
    )

    st.markdown("<hr style='border-color:rgba(110,206,178,0.3);margin:14px 0 12px 0;'>",
                unsafe_allow_html=True)

    st.markdown(
        f"<div style='font-size:0.62rem;color:rgba(255,255,255,0.35);'>"
        f"{datetime.now().strftime('%d %b %Y')}</div>",
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════════════════
# MAIN PAGE – SUBMISSION DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
if page == "📋 Submission Dashboard":

    r_loaded  = st.session_state.rolling_data is not None
    nr_loaded = st.session_state.nonrolling_data is not None
    neither_loaded = not r_loaded and not nr_loaded

    # ══════════════════════════════════════════════════════════════════════
    # LANDING PAGE — shown when no data is loaded yet
    # ══════════════════════════════════════════════════════════════════════
    if neither_loaded:
        st.markdown("""
        <style>
        .lp-hero {
            background: linear-gradient(135deg, #0C2340 0%, #005587 50%, #00857C 100%);
            border-radius: 16px;
            padding: 56px 48px 48px 48px;
            margin-bottom: 32px;
            position: relative;
            overflow: hidden;
        }
        .lp-hero::before {
            content: "";
            position: absolute;
            top: -60px; right: -60px;
            width: 320px; height: 320px;
            border-radius: 50%;
            background: rgba(110,206,178,0.12);
        }
        .lp-hero::after {
            content: "";
            position: absolute;
            bottom: -80px; left: -40px;
            width: 240px; height: 240px;
            border-radius: 50%;
            background: rgba(0,133,124,0.15);
        }
        .lp-title {
            font-size: 2.6rem;
            font-weight: 800;
            color: #FFFFFF;
            letter-spacing: -0.01em;
            line-height: 1.15;
            margin-bottom: 12px;
        }
        .lp-sub {
            font-size: 1.05rem;
            color: #6ECEB2;
            font-weight: 500;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            margin-bottom: 24px;
        }
        .lp-desc {
            font-size: 1rem;
            color: rgba(255,255,255,0.82);
            max-width: 620px;
            line-height: 1.65;
        }
        .lp-badge {
            display: inline-block;
            background: rgba(110,206,178,0.18);
            border: 1px solid rgba(110,206,178,0.45);
            color: #6ECEB2;
            border-radius: 20px;
            padding: 4px 14px;
            font-size: 0.75rem;
            font-weight: 600;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            margin-right: 8px;
            margin-bottom: 20px;
        }
        .lp-feature-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 18px;
            margin-bottom: 28px;
        }
        .lp-feature-card {
            background: white;
            border-radius: 12px;
            padding: 24px 22px;
            border-top: 4px solid #00857C;
            box-shadow: 0 2px 12px rgba(12,35,64,0.08);
        }
        .lp-feature-card h4 {
            color: #0C2340;
            font-size: 0.95rem;
            font-weight: 700;
            margin: 10px 0 8px 0;
        }
        .lp-feature-card p {
            color: #5a6a7a;
            font-size: 0.82rem;
            line-height: 1.55;
            margin: 0;
        }
        .lp-feature-icon {
            font-size: 1.8rem;
        }
        .lp-stat-row {
            display: flex;
            gap: 18px;
            margin-bottom: 28px;
        }
        .lp-stat {
            flex: 1;
            background: linear-gradient(135deg, #0C2340, #005587);
            border-radius: 10px;
            padding: 20px 18px;
            text-align: center;
            color: white;
        }
        .lp-stat-num {
            font-size: 2rem;
            font-weight: 800;
            color: #6ECEB2;
        }
        .lp-stat-label {
            font-size: 0.72rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: rgba(255,255,255,0.7);
            margin-top: 4px;
        }
        .lp-step {
            display: flex;
            align-items: flex-start;
            gap: 16px;
            margin-bottom: 16px;
        }
        .lp-step-num {
            min-width: 36px;
            height: 36px;
            border-radius: 50%;
            background: linear-gradient(135deg, #00857C, #6ECEB2);
            color: white;
            font-weight: 800;
            font-size: 0.95rem;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .lp-step-text strong { color: #0C2340; font-size: 0.9rem; }
        .lp-step-text p { color: #5a6a7a; font-size: 0.82rem; margin: 2px 0 0 0; }
        .lp-divider {
            border: none;
            border-top: 1px solid #e0e8ea;
            margin: 28px 0;
        }
        </style>

        <!-- HERO BANNER -->
        <div class="lp-hero">
            <div class="lp-sub">Merck · Regulatory Affairs</div>
            <div class="lp-title">Submission<br>Intelligence Dashboard</div>
            <div>
                <span class="lp-badge">Rolling Submissions</span>
                <span class="lp-badge">Non-Rolling Submissions</span>
                <span class="lp-badge">Anchor Dates</span>
            </div>
            <div class="lp-desc">
                A unified regulatory filing tracker for upper management — providing real-time visibility
                into wave progress, module completion, schedule variance, and key milestone dates
                across your NDA / BLA portfolio.
            </div>
        </div>

        <!-- CAPABILITY PILLS -->
        <div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:28px;">
            <div style="background:linear-gradient(135deg,#0C2340,#005587);border-radius:10px;padding:16px 22px;color:white;flex:1;min-width:160px;">
                <div style="font-size:0.68rem;text-transform:uppercase;letter-spacing:0.08em;color:#6ECEB2;margin-bottom:6px;">Rolling</div>
                <div style="font-size:0.9rem;font-weight:700;">Wave-by-Wave Tracking</div>
            </div>
            <div style="background:linear-gradient(135deg,#005587,#00857C);border-radius:10px;padding:16px 22px;color:white;flex:1;min-width:160px;">
                <div style="font-size:0.68rem;text-transform:uppercase;letter-spacing:0.08em;color:#6ECEB2;margin-bottom:6px;">Non-Rolling</div>
                <div style="font-size:0.9rem;font-weight:700;">Module Group Analysis</div>
            </div>
            <div style="background:linear-gradient(135deg,#00857C,#6ECEB2);border-radius:10px;padding:16px 22px;color:#0C2340;flex:1;min-width:160px;">
                <div style="font-size:0.68rem;text-transform:uppercase;letter-spacing:0.08em;color:#0C2340;opacity:0.7;margin-bottom:6px;">Filter</div>
                <div style="font-size:0.9rem;font-weight:700;">Regional · Central · Local</div>
            </div>
            <div style="background:linear-gradient(135deg,#0C2340,#00857C);border-radius:10px;padding:16px 22px;color:white;flex:1;min-width:160px;">
                <div style="font-size:0.68rem;text-transform:uppercase;letter-spacing:0.08em;color:#6ECEB2;margin-bottom:6px;">Milestones</div>
                <div style="font-size:0.9rem;font-weight:700;">Anchor Date Manager</div>
            </div>
        </div>

        <!-- FEATURE CARDS -->
        <div class="lp-feature-grid">
            <div class="lp-feature-card">
                <div class="lp-feature-icon">🌊</div>
                <h4>Rolling Submission Tracker</h4>
                <p>Monitor wave-by-wave filing progress with completion rates, schedule variance,
                and Gantt visualizations across all submission waves.</p>
            </div>
            <div class="lp-feature-card">
                <div class="lp-feature-icon">📦</div>
                <h4>Non-Rolling Module Analysis</h4>
                <p>Drill into CTD module groups filtered by Regional or Central components.
                Instantly see document status breakdowns and outstanding items.</p>
            </div>
            <div class="lp-feature-card">
                <div class="lp-feature-icon">📅</div>
                <h4>Gantt & Schedule Variance</h4>
                <p>Side-by-side Planned vs Actual timeline bars with a live "Today" indicator
                and automated variance categorization.</p>
            </div>
            <div class="lp-feature-card">
                <div class="lp-feature-icon">📌</div>
                <h4>Anchor Date Management</h4>
                <p>Manually add milestones or bulk-import from CSV / Excel.
                Track Plan Baseline, Agency Submission, Decision dates and more.</p>
            </div>
            <div class="lp-feature-card">
                <div class="lp-feature-icon">🗂️</div>
                <h4>Regional vs Central Filter</h4>
                <p>Instantly segment non-rolling submissions by Component Source —
                Local, Central, or Regional — across all analysis tabs.</p>
            </div>
            <div class="lp-feature-card">
                <div class="lp-feature-icon">⬇️</div>
                <h4>Export & Drill-Through</h4>
                <p>Download filtered datasets as CSV for any wave or module group.
                Built for regulatory operations and upper management reporting.</p>
            </div>
        </div>

        <hr class="lp-divider">

        <!-- HOW TO GET STARTED -->
        <div style="font-size:1rem;font-weight:700;color:#0C2340;margin-bottom:16px;">
            🚀 &nbsp;Get Started in 3 Steps
        </div>
        <div class="lp-step">
            <div class="lp-step-num">1</div>
            <div class="lp-step-text">
                <strong>Upload your Rolling Submission Excel file</strong>
                <p>Open the Rolling Submission tab → click the upload area → select your PSPM Planner Report sheet.</p>
            </div>
        </div>
        <div class="lp-step">
            <div class="lp-step-num">2</div>
            <div class="lp-step-text">
                <strong>Upload your Non-Rolling Submission Excel file</strong>
                <p>Switch to the Non-Rolling tab and repeat — then use the Component Source filter to toggle Regional vs Central.</p>
            </div>
        </div>
        <div class="lp-step">
            <div class="lp-step-num">3</div>
            <div class="lp-step-text">
                <strong>Add Anchor Dates</strong>
                <p>Navigate to 📌 Anchor Dates in the sidebar to manually enter milestones or bulk-import a CSV / Excel file.</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Always show the page header & tabs (tabs contain uploads) ─────────
    page_header(
        "📋 Submission Dashboard",
        subtitle="Regulatory Filing Progress · Upper Management View",
    )

    # ── TOP-LEVEL TABS: Rolling  |  Non-Rolling ───────────────────────────
    tab_rolling, tab_nonrolling = st.tabs(
        ["🌊  Rolling Submission", "📦  Non-Rolling Submission"]
    )

    # ══════════════════════════════════════
    # TAB 1 – ROLLING SUBMISSION
    # ══════════════════════════════════════
    with tab_rolling:

        # ── Upload ────────────────────────────────────────────────────────
        with st.expander("📂  Upload Rolling Submission File", expanded=st.session_state.rolling_data is None):
            col_up, col_hint = st.columns([2, 1])
            with col_up:
                f = st.file_uploader("Rolling Excel file", type=["xlsx", "xls"],
                                     key="rolling_uploader", label_visibility="collapsed")
                if f is not None:
                    try:
                        xf    = pd.ExcelFile(f)
                        sheet = st.selectbox("Select sheet", xf.sheet_names, key="r_sheet")
                        df_r  = read_submission_excel(f, sheet)
                        df_r  = clean_submission_data(df_r)
                        st.session_state.rolling_data = df_r
                        st.success(f"✅ Loaded **{f.name}** — {len(df_r):,} records")
                    except Exception as e:
                        st.error(f"Error: {e}")
            with col_hint:
                st.markdown("""
                <div style='background:#f0f6f6;border-left:4px solid #00857C;
                            border-radius:6px;padding:12px;font-size:0.82rem;color:#0C2340;'>
                <b>Expected columns</b><br>
                Task Name · Planned Start · Planned Finish<br>
                Actual Start · Actual Finish · Filing Status<br>
                Wave · Component ID
                </div>""", unsafe_allow_html=True)

        if st.session_state.rolling_data is not None:
            df = st.session_state.rolling_data
            total, completed, remaining, planned, rate, variance = calculate_metrics(df)

            # ── Inner sub-tabs ────────────────────────────────────────────
            st_r1, st_r2, st_r3, st_r4 = st.tabs([
                "📊 Executive Summary",
                "🌊 Wave Analysis",
                "🔍 Component Drill-Through",
                "📅 Gantt Chart",
            ])

            # ── R1: Executive Summary ─────────────────────────────────────
            with st_r1:
                section_label("Key Performance Indicators")
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("Total Documents",   f"{total:,}")
                k2.metric("Completed",         f"{completed:,}")
                k3.metric("Remaining",         f"{remaining:,}")
                k4.metric("Completion Rate",   f"{rate:.1%}")
                k5.metric("Σ Finish Variance", f"{variance:,.0f} days")

                st.markdown("---")
                cg, cd = st.columns(2)
                with cg:
                    st.plotly_chart(render_gauge(rate, "Overall Completion Rate"),
                                    use_container_width=True)
                with cd:
                    fig_d = render_status_donut(df)
                    if fig_d:
                        st.plotly_chart(fig_d, use_container_width=True)

                if "Wave" in df.columns:
                    st.markdown("---")
                    section_label("Progress by Wave")
                    ws  = get_wave_summary(df)
                    pct = ws.melt(id_vars="Wave", value_vars=["Done_%", "Left_%"],
                                  var_name="Type", value_name="Pct")
                    pct["Type"] = pct["Type"].map({"Done_%": "Completed", "Left_%": "Remaining"})
                    fig_w = px.bar(pct, x="Wave", y="Pct", color="Type",
                                   barmode="stack", text="Pct",
                                   color_discrete_map={"Completed": MERCK_TEAL,
                                                       "Remaining": MERCK_GRAY})
                    fig_w.update_traces(texttemplate="%{text:.0f}%",
                                        textposition="inside", textfont_color="white")
                    fig_w.update_layout(yaxis_title="Percent (%)", xaxis_title="Wave",
                                        yaxis_range=[0, 115],
                                        plot_bgcolor="white", paper_bgcolor="white",
                                        font=dict(color=MERCK_BLUE), height=360)
                    fig_w.update_xaxes(showgrid=False)
                    fig_w.update_yaxes(gridcolor="#E8ECEC")
                    st.plotly_chart(fig_w, use_container_width=True)

            # ── R2: Wave Analysis ─────────────────────────────────────────
            with st_r2:
                if "Wave" not in df.columns:
                    st.error("No Wave column found.")
                else:
                    ws = get_wave_summary(df)
                    section_label("Wave Summary Table")
                    disp = ws[["Wave", "Total", "Completed", "Remaining", "Planned", "Rate_%"]].copy()
                    disp["Rate_%"] = disp["Rate_%"].map("{:.1f}%".format)
                    disp.columns   = ["Wave", "Total", "Completed", "Remaining",
                                      "Planned Completed", "Completion Rate"]
                    st.dataframe(disp, use_container_width=True, hide_index=True)

                    st.markdown("---")
                    c1, c2 = st.columns(2)
                    with c1:
                        fig_r = px.bar(ws, x="Wave", y="Rate_%",
                                       title="Completion Rate by Wave (%)",
                                       text="Rate_%",
                                       color_discrete_sequence=[MERCK_TEAL])
                        fig_r.update_traces(texttemplate="%{text:.1f}%",
                                            textposition="outside")
                        fig_r.update_layout(yaxis_title="Completion Rate (%)",
                                            yaxis_range=[0, 120],
                                            plot_bgcolor="white", paper_bgcolor="white",
                                            font=dict(color=MERCK_BLUE))
                        fig_r.update_xaxes(showgrid=False)
                        fig_r.update_yaxes(gridcolor="#E8ECEC")
                        st.plotly_chart(fig_r, use_container_width=True)
                    with c2:
                        fig_i = px.bar(ws, x="Wave", y="Remaining",
                                       title="Remaining Documents by Wave",
                                       text="Remaining",
                                       color_discrete_sequence=[MERCK_ORANGE])
                        fig_i.update_traces(texttemplate="%{text:,}",
                                            textposition="outside")
                        fig_i.update_layout(yaxis_title="Remaining",
                                            plot_bgcolor="white", paper_bgcolor="white",
                                            font=dict(color=MERCK_BLUE))
                        fig_i.update_xaxes(showgrid=False)
                        fig_i.update_yaxes(gridcolor="#E8ECEC")
                        st.plotly_chart(fig_i, use_container_width=True)

                    st.markdown("---")
                    pct = ws.melt(id_vars="Wave", value_vars=["Done_%", "Left_%"],
                                  var_name="Type", value_name="Pct")
                    pct["Type"] = pct["Type"].map({"Done_%": "Completed", "Left_%": "Remaining"})
                    fig_h = px.bar(pct, y="Wave", x="Pct", color="Type",
                                   orientation="h", barmode="stack",
                                   title="Completed vs Remaining by Wave (Horizontal)",
                                   color_discrete_map={"Completed": MERCK_TEAL,
                                                       "Remaining": MERCK_GRAY})
                    fig_h.update_layout(xaxis_title="Percent (%)", yaxis_title="Wave",
                                        xaxis_range=[0, 115],
                                        plot_bgcolor="white", paper_bgcolor="white",
                                        font=dict(color=MERCK_BLUE), height=360)
                    fig_h.update_xaxes(showgrid=True, gridcolor="#E8ECEC")
                    fig_h.update_yaxes(showgrid=False)
                    st.plotly_chart(fig_h, use_container_width=True)

            # ── R3: Component Drill-Through ───────────────────────────────
            with st_r3:
                cf1, cf2 = st.columns([2, 2])
                with cf1:
                    if "Wave" in df.columns:
                        sel_w = st.selectbox("Select Wave",
                                             sorted(df["Wave"].dropna().unique()),
                                             key="r_drill_wave")
                        ddf = df[df["Wave"] == sel_w]
                    else:
                        ddf = df
                        st.info("No Wave column — showing all records.")
                with cf2:
                    search = st.text_input("🔎 Search Task Name", key="r_search")
                    if search:
                        ddf = ddf[ddf["Task Name"].astype(str)
                                    .str.contains(search, case=False, na=False)]

                if not ddf.empty:
                    section_label("Wave Metrics")
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Total",     f"{len(ddf):,}")
                    m2.metric("Completed", f"{ddf['Actually Completed'].eq(True).sum():,}")
                    m3.metric("Remaining", f"{ddf['Actually Completed'].eq(False).sum():,}")

                st.markdown("---")
                dcols = ["Task Name", "Wave", "Filing Status",
                         "Planned Start", "Planned Finish",
                         "Actual Start",  "Actual Finish"]
                dcols = [c for c in dcols if c in ddf.columns]
                st.dataframe(ddf[dcols], use_container_width=True, hide_index=True)

                csv = ddf[dcols].to_csv(index=False).encode()
                st.download_button("⬇️ Download Detail",
                                   data=csv, file_name="rolling_drill_through.csv",
                                   mime="text/csv")

            # ── R4: Gantt ─────────────────────────────────────────────────
            with st_r4:
                gcol = "Wave" if "Wave" in df.columns else "Module"
                render_gantt_section(df, gcol, tab_key="rolling")

    # ══════════════════════════════════════
    # TAB 2 – NON-ROLLING SUBMISSION
    # ══════════════════════════════════════
    with tab_nonrolling:

        # ── Upload ────────────────────────────────────────────────────────
        with st.expander("📂  Upload Non-Rolling Submission File",
                         expanded=st.session_state.nonrolling_data is None):
            col_up2, col_hint2 = st.columns([2, 1])
            with col_up2:
                f2 = st.file_uploader("Non-Rolling Excel file", type=["xlsx", "xls"],
                                      key="nonrolling_uploader", label_visibility="collapsed")
                if f2 is not None:
                    try:
                        xf2    = pd.ExcelFile(f2)
                        sheet2 = st.selectbox("Select sheet", xf2.sheet_names, key="nr_sheet")
                        df_nr  = read_submission_excel(f2, sheet2)
                        df_nr  = clean_submission_data(df_nr)
                        st.session_state.nonrolling_data = df_nr
                        st.success(f"✅ Loaded **{f2.name}** — {len(df_nr):,} records")
                    except Exception as e:
                        st.error(f"Error: {e}")
            with col_hint2:
                st.markdown("""
                <div style='background:#f0f6f6;border-left:4px solid #00857C;
                            border-radius:6px;padding:12px;font-size:0.82rem;color:#0C2340;'>
                <b>Expected columns</b><br>
                Task Name · Planned Start · Planned Finish<br>
                Actual Start · Actual Finish · Filing Status<br>
                Module · Component ID (e.g. 1.3.4)
                </div>""", unsafe_allow_html=True)

        if st.session_state.nonrolling_data is not None:
            df2 = st.session_state.nonrolling_data
            total2, completed2, remaining2, planned2, rate2, variance2 = calculate_metrics(df2)

            # ── Regional Component Source filter ──────────────────────────
            if "Component Source" in df2.columns:
                src_options = ["All"] + sorted(
                    df2["Component Source"].dropna().astype(str).str.strip().unique().tolist()
                )
                sel_src = st.selectbox(
                    "🗂️ Filter by Component Source (Regional / Central / Local)",
                    src_options, key="nr_comp_source",
                )
            else:
                sel_src = "All"

            df2_view = df2.copy()
            if sel_src != "All" and "Component Source" in df2_view.columns:
                df2_view = df2_view[
                    df2_view["Component Source"].astype(str).str.strip().str.lower()
                    == sel_src.strip().lower()
                ]

            mod_sum = get_nonrolling_summary(df2_view)

            st_nr1, st_nr2, st_nr3, st_nr4 = st.tabs([
                "📊 Executive Summary",
                "📦 Module Analysis",
                "🔍 Module Drill-Down",
                "📅 Gantt Chart",
            ])

            # ── NR1: Executive Summary ────────────────────────────────────
            with st_nr1:
                section_label("Key Performance Indicators")
                pct_c  = completed2 / total2 if total2 > 0 else 0
                pct_nc = 1 - pct_c
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("Total Documents",  f"{total2:,}")
                k2.metric("Completed",        f"{completed2:,}")
                k3.metric("Remaining",        f"{remaining2:,}")
                k4.metric("% Complete",       f"{pct_c:.1%}")
                k5.metric("% Not Complete",   f"{pct_nc:.1%}")

                st.markdown("---")
                cg2, cs2 = st.columns(2)
                with cg2:
                    st.plotly_chart(render_gauge(rate2, "Overall Module Completion"),
                                    use_container_width=True)
                with cs2:
                    if not mod_sum.empty:
                        stk = mod_sum.melt(
                            id_vars="Module Group",
                            value_vars=["Completed", "Remaining"],
                            var_name="Status", value_name="Count"
                        )
                        fig_stk = px.bar(
                            stk, x="Module Group", y="Count", color="Status",
                            title="Completed vs Remaining by Module Group",
                            barmode="stack",
                            color_discrete_map={"Completed": MERCK_TEAL,
                                                "Remaining": MERCK_ORANGE}
                        )
                        fig_stk.update_layout(
                            plot_bgcolor="white", paper_bgcolor="white",
                            font=dict(color=MERCK_BLUE), height=270,
                            margin=dict(t=40, b=10)
                        )
                        fig_stk.update_xaxes(showgrid=False)
                        fig_stk.update_yaxes(gridcolor="#E8ECEC")
                        st.plotly_chart(fig_stk, use_container_width=True)

            # ── NR2: Module Analysis ──────────────────────────────────────
            with st_nr2:
                if mod_sum.empty:
                    st.error("Could not derive Module Groups. Check Component ID column.")
                else:
                    section_label("Module Group Summary Table")
                    disp2 = mod_sum[["Module Group", "Total", "Completed",
                                     "Remaining", "Pct_Complete", "Pct_Remaining"]].copy()
                    disp2["Pct_Complete"]  = disp2["Pct_Complete"].map("{:.1%}".format)
                    disp2["Pct_Remaining"] = disp2["Pct_Remaining"].map("{:.1%}".format)
                    disp2.columns = ["Module Group", "Total", "Completed",
                                     "Remaining", "% Complete", "% Not Complete"]
                    st.dataframe(disp2, use_container_width=True, hide_index=True)

                    st.markdown("---")
                    c1n, c2n = st.columns(2)
                    with c1n:
                        fig_mp = px.bar(mod_sum, x="Module Group", y="Pct_Complete",
                                        title="% Complete by Module Group",
                                        text="Pct_Complete",
                                        color_discrete_sequence=[MERCK_TEAL])
                        fig_mp.update_traces(texttemplate="%{text:.1%}",
                                             textposition="outside")
                        fig_mp.update_layout(
                            yaxis=dict(tickformat=".0%", range=[0, 1.25],
                                       title="% Complete"),
                            xaxis_title="Module Group",
                            plot_bgcolor="white", paper_bgcolor="white",
                            font=dict(color=MERCK_BLUE)
                        )
                        fig_mp.update_xaxes(showgrid=False)
                        fig_mp.update_yaxes(gridcolor="#E8ECEC")
                        st.plotly_chart(fig_mp, use_container_width=True)

                    with c2n:
                        pct2 = mod_sum.melt(
                            id_vars="Module Group",
                            value_vars=["Pct_Complete", "Pct_Remaining"],
                            var_name="Metric", value_name="Value"
                        )
                        pct2["Metric"] = pct2["Metric"].map(
                            {"Pct_Complete": "% Complete",
                             "Pct_Remaining": "% Not Complete"}
                        )
                        fig_mg = px.bar(
                            pct2, x="Module Group", y="Value", color="Metric",
                            title="% Complete vs % Not Complete",
                            barmode="group",
                            color_discrete_map={"% Complete":     MERCK_TEAL,
                                                "% Not Complete": MERCK_GRAY}
                        )
                        fig_mg.update_layout(
                            yaxis=dict(tickformat=".0%", range=[0, 1.25],
                                       title="Percent"),
                            xaxis_title="Module Group",
                            plot_bgcolor="white", paper_bgcolor="white",
                            font=dict(color=MERCK_BLUE)
                        )
                        fig_mg.update_xaxes(showgrid=False)
                        fig_mg.update_yaxes(gridcolor="#E8ECEC")
                        st.plotly_chart(fig_mg, use_container_width=True)

            # ── NR3: Module Drill-Down ────────────────────────────────────
            with st_nr3:
                if mod_sum.empty:
                    st.error("No module data available.")
                else:
                    sel_mg = st.selectbox("Select Module Group",
                                          mod_sum["Module Group"].tolist(),
                                          key="nr_mg_select")

                    det = df2_view.copy()
                    if "Component ID" in det.columns:
                        det["_mg"] = det["Component ID"].apply(compute_module_group)
                        det = det[det["_mg"] == sel_mg]
                    elif "Module" in det.columns:
                        det = det[det["Module"].astype(str).str.startswith(sel_mg)]

                    det["IsComplete"] = (
                        det["Filing Status"].astype(str).str.strip().str.lower() == "completed"
                    ).map({True: "✅ Yes", False: "❌ No"})

                    section_label(f"Module Group {sel_mg} — Detail")
                    d1, d2, d3 = st.columns(3)
                    d1.metric("Total",     len(det))
                    d2.metric("Completed", (det["IsComplete"] == "✅ Yes").sum())
                    d3.metric("Remaining", (det["IsComplete"] == "❌ No").sum())

                    dcols2 = ["Component ID", "Task Name", "Filing Status", "IsComplete",
                              "Planned Start", "Planned Finish", "Actual Start", "Actual Finish"]
                    dcols2 = [c for c in dcols2 if c in det.columns]
                    st.dataframe(det[dcols2], use_container_width=True, hide_index=True)

                    csv2 = det[dcols2].to_csv(index=False).encode()
                    st.download_button("⬇️ Download Detail", data=csv2,
                                       file_name=f"module_{sel_mg}_detail.csv",
                                       mime="text/csv")

            # ── NR4: Gantt ────────────────────────────────────────────────
            with st_nr4:
                gcol2 = "Module" if "Module" in df2_view.columns else "Wave"
                render_gantt_section(df2_view, gcol2, tab_key="nonrolling")


# ═══════════════════════════════════════════════════════════════════════════
# ANCHOR DATES PAGE
# ═══════════════════════════════════════════════════════════════════════════
elif page == "📌 Anchor Dates":
    page_header("📌 Anchor Dates", subtitle="Manual milestone entry & tracking")

    col_form, col_tbl = st.columns([1, 2])

    with col_form:
        section_label("Add New Milestone")
        with st.form("anchor_form"):
            aname  = st.text_input("Milestone Name")
            adate  = st.date_input("Target Date")
            astatus = st.selectbox("Status", ["Complete", "In Progress", "Not Started"])
            sub    = st.form_submit_button("➕ Add")
            if sub:
                if len(st.session_state.anchor_dates) >= 18:
                    st.warning("Maximum 18 anchor dates reached.")
                elif not aname.strip():
                    st.warning("Name cannot be empty.")
                else:
                    st.session_state.anchor_dates = pd.concat([
                        st.session_state.anchor_dates,
                        pd.DataFrame([[aname, adate, astatus]],
                                     columns=["Anchor Date", "Date", "Status"])
                    ], ignore_index=True)
                    st.success("✅ Added.")

    with col_tbl:
        section_label("Current Anchor Dates")
        if not st.session_state.anchor_dates.empty:
            cmap_s = {"Complete":    MERCK_TEAL,
                      "In Progress": MERCK_ORANGE,
                      "Not Started": MERCK_GRAY}
            styled = st.session_state.anchor_dates.style.applymap(
                lambda v: f"color:{cmap_s.get(v,'black')};font-weight:bold;",
                subset=["Status"]
            )
            st.dataframe(styled, use_container_width=True)

            row_del = st.selectbox(
                "Row to remove",
                st.session_state.anchor_dates.index,
                format_func=lambda x:
                    f"Row {x+1}: {st.session_state.anchor_dates.loc[x,'Anchor Date']}"
            )
            c1, c2 = st.columns(2)
            with c1:
                if st.button("🗑️ Remove Selected"):
                    st.session_state.anchor_dates = (
                        st.session_state.anchor_dates
                        .drop(row_del).reset_index(drop=True)
                    )
                    st.success("Removed.")
            with c2:
                csv_a = st.session_state.anchor_dates.to_csv(index=False).encode()
                st.download_button("⬇️ Export CSV", data=csv_a,
                                   file_name="anchor_dates.csv", mime="text/csv")
        else:
            st.info("No anchor dates added yet.")