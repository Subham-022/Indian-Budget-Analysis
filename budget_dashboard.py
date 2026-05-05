import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Page configuration

st.set_page_config(
    page_title="India Budget Dashboard | 2019-2024",
    page_icon="IN",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600&display=swap');
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400&display=swap');
    html, body, [class*="css"]  { font-family: 'Poppins', sans-serif; }

    .stApp { background-color: #f6f7fb; }

    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1f2d3d;
        margin-bottom: 0;
    }
    .main-sub {
        font-size: 0.95rem;
        color: #6b7280;
        margin-top: 4px;
        margin-bottom: 18px;
    }
    .tricolor-bar {
        height: 4px;
        width: 90px;
        background: linear-gradient(90deg, #ff9933 0%, #ffffff 50%, #138808 100%);
        border: 1px solid #e5e7eb;
        border-radius: 3px;
        margin: 6px 0 14px 0;
    }

    .kpi-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 16px 18px;
        box-shadow: 0 1px 2px rgba(17, 24, 39, 0.04);
        height: 100%;
    }
    .kpi-label {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #6b7280;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-family: 'JetBrains Mono', monospace;
        font-size: 1.5rem;
        font-weight: 600;
        color: #1f2d3d;
    }
    .kpi-sub {
        font-size: 0.78rem;
        color: #16a34a;
        margin-top: 4px;
    }
    .kpi-sub.neg { color: #dc2626; }

    .section-title {
        font-size: 1.15rem;
        font-weight: 600;
        color: #1f2d3d;
        margin: 22px 0 6px 0;
    }
    .section-desc {
        font-size: 0.85rem;
        color: #6b7280;
        margin-bottom: 14px;
    }

    .footer {
        text-align: center;
        color: #9ca3af;
        font-size: 0.78rem;
        padding: 16px 0;
    }

    /* Upload zone label tweak */
    .upload-note {
        font-size: 0.78rem;
        color: #9ca3af;
        margin-top: 4px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# Default file paths 

DEFAULT_MASTER   = Path(r"C:\Users\subha\Downloads\Indian_Budget_Analysis_Master.xlsx")

SHEET_LOCAL_PATHS = {
    "States": Path(r"C:\Users\subha\Downloads\states_data.xlsx"),
    "Ministries": Path(r"C:\Users\subha\Downloads\ministries_data.xlsx"),
    "Schemes": Path(r"C:\Users\subha\Downloads\schemes_data.xlsx"),
    "Union Territories": Path(r"C:\Users\subha\Downloads\union_territories_data.xlsx"),
    "Macro": Path(r"C:\Users\subha\Downloads\macro_data.xlsx"),
}

# Column name in row 0 / header detection for each  sheet
SHEET_ENTITY_COL = {
    "Ministries":        "Ministry",
    "States":            "State",
    "Schemes":           "Scheme",
    "Union Territories": "Union Territory",
    "Macro":             "Fiscal Year",
}

YEAR_COLS = ["2019-20", "2020-21", "2021-22", "2022-23", "2023-24"]


# Data cleaning

def _clean_sheet(df_raw: pd.DataFrame, entity_col_name: str) -> pd.DataFrame:
    if df_raw.columns[0].startswith("Unnamed"):
        df_raw = df_raw.iloc[:, 1:]

    header_row_idx = None
    for i in range(min(6, len(df_raw))):
        first_cell = str(df_raw.iloc[i, 0]).strip()
        if first_cell.lower() == entity_col_name.lower():
            header_row_idx = i
            break
    if header_row_idx is None:
        header_row_idx = 3

    df = df_raw.iloc[header_row_idx + 1 :].copy()
    df.columns = df_raw.iloc[header_row_idx].tolist()
    df = df.reset_index(drop=True)

    first_col = df.columns[0]
    df = df[~df[first_col].astype(str).str.lower().str.contains("total", na=False)]
    df = df.dropna(how="all").reset_index(drop=True)

    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    rename_map = {}
    for c in df.columns:
        s = str(c).strip()
        if s in YEAR_COLS:
            rename_map[c] = s
    df = df.rename(columns=rename_map)

    return df


@st.cache_data(show_spinner=False)
def load_workbook(file_source) -> dict:
    ministries = _clean_sheet(
        pd.read_excel(file_source, sheet_name="1. Ministries"), "Ministry"
    )
    states = _clean_sheet(
        pd.read_excel(file_source, sheet_name="2. States"), "State"
    )
    schemes = _clean_sheet(
        pd.read_excel(file_source, sheet_name="3. Schemes"), "Scheme"
    )
    uts = _clean_sheet(
        pd.read_excel(file_source, sheet_name="4. Union Territories"),
        "Union Territory",
    )
    macro = _clean_sheet(
        pd.read_excel(file_source, sheet_name="5. Macroeconomics"), "Fiscal Year"
    )

    return {
        "Ministries": ministries,
        "States": states,
        "Schemes": schemes,
        "Union Territories": uts,
        "Macro": macro,
    }


@st.cache_data(show_spinner=False)
def load_standalone(file_source, entity_col: str) -> pd.DataFrame:
    """Load a single-sheet Excel file (no multi-sheet workbook assumed)."""
    df_raw = pd.read_excel(file_source)
    return _clean_sheet(df_raw, entity_col)


def resolve_sheet_source(sheet_name: str, uploaded_file):
    """
    Priority order for a given sheet:
      1. File uploaded via the per-page uploader in the sidebar
      2. Local path in SHEET_LOCAL_PATHS (if the file actually exists on disk)
      3. Master workbook (already loaded in `data`)

    Returns (df, source_label) where df is the cleaned DataFrame.
    """
    entity_col = SHEET_ENTITY_COL.get(sheet_name, sheet_name)

    if uploaded_file is not None:
        df = load_standalone(uploaded_file, entity_col)
        return df, f"uploaded file — {uploaded_file.name}"

    local_path = SHEET_LOCAL_PATHS.get(sheet_name)
    if local_path and local_path.exists():
        df = load_standalone(str(local_path), entity_col)
        return df, f"local file — {local_path.name}"

    # Fall back to master workbook
    return data[sheet_name], "master workbook"

# Sidebar — navigation + master file

with st.sidebar:
    st.markdown(
        "<div style='font-size:1.25rem;font-weight:700;color:#1f2d3d;"
        "margin-bottom:2px'>India Budget</div>"
        "<div style='font-size:0.8rem;color:#6b7280;margin-bottom:14px;"
        "font-family:JetBrains Mono,monospace'>FY 2019-20 to 2023-24</div>",
        unsafe_allow_html=True,
    )

    uploaded_master = st.file_uploader(
        "Replace master workbook (optional)",
        type=["xlsx"],
        help="Must follow the same sheet structure as the provided master file.",
    )
    master_source = uploaded_master if uploaded_master is not None else DEFAULT_MASTER

    if not uploaded_master and not DEFAULT_MASTER.exists():
        st.error(
            "Master file not found.\n\n"
            "Put `Indian_Budget_Analysis_Master.xlsx` next to `app.py`, "
            "or upload it above."
        )
        st.stop()

    st.markdown("---")
    page = st.radio(
        "Navigate",
        [
            "Overview",
            "Ministries",
            "States",
            "Schemes",
            "Union Territories",
            "Compare Entities",
            "Data Explorer",
        ],
        index=0,
    )

    st.markdown("---")

    # Per-sheet uploader 
    sheet_pages = ["Ministries", "States", "Schemes", "Union Territories"]
    per_sheet_upload = None
    if page in sheet_pages:
        st.markdown(
            f"<div style='font-size:0.8rem;font-weight:600;color:#1f2d3d;"
            f"margin-bottom:4px'>Load a different {page} file</div>",
            unsafe_allow_html=True,
        )
        per_sheet_upload = st.file_uploader(
            f"Upload {page} Excel",
            type=["xlsx"],
            key=f"upload_{page}",
            help=(
                f"Single-sheet Excel file. First column should be the "
                f"{SHEET_ENTITY_COL.get(page, page)} name, "
                f"remaining columns are year-wise allocations in Rs Crore."
            ),
        )
        local_path = SHEET_LOCAL_PATHS.get(page)
        if local_path and local_path.exists():
            st.markdown(
                f"<div class='upload-note'>Default: <code>{local_path.name}</code> "
                f"(found on disk)</div>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                "<div class='upload-note'>No local file found for this sheet — "
                "using master workbook.</div>",
                unsafe_allow_html=True,
            )

    st.markdown("---")
    st.caption(
        "Source: indiabudget.gov.in, RBI Annual Report, Ministry of Finance.\n\n"
        "Values in Rupees Crore."
    )

# Load master workbook
try:
    data = load_workbook(master_source)
except Exception as e:
    st.error(f"Could not read the workbook: {e}")
    st.stop()


# Shared helpers

def fmt_cr(n: float) -> str:
    if pd.isna(n):
        return "-"
    if abs(n) >= 1_00_000:
        return f"Rs {n/1e5:.2f} L Cr"
    if abs(n) >= 1000:
        return f"Rs {n/1000:.1f}K Cr"
    return f"Rs {n:,.0f} Cr"


def growth(old: float, new: float) -> float:
    if pd.isna(old) or old == 0 or pd.isna(new):
        return 0.0
    return (new - old) / old * 100


def kpi_card(col, label: str, value: str, sub: str = "", neg: bool = False):
    sub_cls = "kpi-sub neg" if neg else "kpi-sub"
    col.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="{sub_cls}">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def year_totals(df: pd.DataFrame) -> pd.Series:
    existing = [c for c in YEAR_COLS if c in df.columns]
    return df[existing].sum(numeric_only=True)


def top_n(df: pd.DataFrame, entity_col: str, year: str, n: int = 10) -> pd.DataFrame:
    return df[[entity_col, year]].dropna().sort_values(year, ascending=False).head(n)


def entity_trend_fig(df: pd.DataFrame, entity_col: str, entities: list) -> go.Figure:
    fig = go.Figure()
    colors = px.colors.qualitative.Set2
    for i, name in enumerate(entities):
        row = df[df[entity_col] == name]
        if row.empty:
            continue
        y_vals = [row.iloc[0].get(y, np.nan) for y in YEAR_COLS]
        fig.add_trace(
            go.Scatter(
                x=YEAR_COLS,
                y=y_vals,
                mode="lines+markers",
                name=str(name),
                line=dict(color=colors[i % len(colors)], width=2.5),
                marker=dict(size=8),
                hovertemplate="<b>%{x}</b><br>Rs %{y:,.0f} Cr<extra></extra>",
            )
        )
    fig.update_layout(
        height=420,
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        margin=dict(l=40, r=30, t=40, b=40),
        xaxis=dict(showgrid=False, linecolor="#e5e7eb"),
        yaxis=dict(gridcolor="#f1f5f9", title="Rs Crore"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    return fig


def heatmap_fig(df: pd.DataFrame, entity_col: str, title: str) -> go.Figure:
    years = [c for c in YEAR_COLS if c in df.columns]
    pivot = df.set_index(entity_col)[years].fillna(0)
    norm = pivot.div(pivot.max(axis=1).replace(0, np.nan), axis=0).fillna(0)

    fig = go.Figure(
        data=go.Heatmap(
            z=norm.values,
            x=years,
            y=norm.index.astype(str),
            colorscale=[[0, "#f1f5f9"], [0.5, "#fbbf24"], [1, "#b45309"]],
            hovertemplate=(
                "<b>%{y}</b><br>Year: %{x}<br>"
                "Relative intensity: %{z:.2f}<extra></extra>"
            ),
            colorbar=dict(title="Row max = 1"),
        )
    )
    fig.update_layout(
        title=title,
        height=max(360, 28 * len(pivot)),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        margin=dict(l=140, r=30, t=50, b=40),
        xaxis=dict(side="top"),
    )
    return fig

# Page: Overview

def render_overview():
    st.markdown('<div class="main-title">India Union Budget Dashboard</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="tricolor-bar"></div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="main-sub">Five fiscal years of Union Budget allocations — '
        'ministries, states, schemes and union territories in one place.</div>',
        unsafe_allow_html=True,
    )

    ministries = data["Ministries"]
    total_24 = year_totals(ministries).get("2023-24", np.nan)
    total_23 = year_totals(ministries).get("2022-23", np.nan)
    total_20 = year_totals(ministries).get("2019-20", np.nan)

    yoy = growth(total_23, total_24)
    five_yr = growth(total_20, total_24)

    top_m = ministries.sort_values("2023-24", ascending=False).iloc[0]
    top_m_name = top_m["Ministry"]
    top_m_val = top_m["2023-24"]

    c1, c2, c3, c4 = st.columns(4)
    kpi_card(c1, "Ministries outlay — FY 2023-24", fmt_cr(total_24),
             "Sum of tracked ministries")
    kpi_card(c2, "Year-on-year change", f"{yoy:+.1f} %",
             f"vs FY 2022-23 ({fmt_cr(total_23)})", neg=yoy < 0)
    kpi_card(c3, "5-year growth", f"{five_yr:+.1f} %",
             f"Since FY 2019-20 ({fmt_cr(total_20)})", neg=five_yr < 0)
    kpi_card(c4, "Largest ministry — 2023-24",
             (top_m_name[:22] + "...") if len(str(top_m_name)) > 25 else str(top_m_name),
             fmt_cr(top_m_val))

    st.markdown('<div class="section-title">Allocation totals by dataset</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<div class="section-desc">Total budgeted across ministries, states, '
        'schemes, and union territories for each year.</div>',
        unsafe_allow_html=True,
    )

    dataset_totals = pd.DataFrame(
        {
            "Ministries": year_totals(data["Ministries"]),
            "States": year_totals(data["States"]),
            "Schemes": year_totals(data["Schemes"]),
            "Union Territories": year_totals(data["Union Territories"]),
        }
    ).reindex(YEAR_COLS).reset_index().rename(columns={"index": "Year"})

    long_df = dataset_totals.melt(id_vars="Year", var_name="Dataset",
                                  value_name="Amount")

    fig_totals = px.bar(
        long_df, x="Year", y="Amount", color="Dataset",
        barmode="group",
        color_discrete_sequence=px.colors.qualitative.Set2,
        labels={"Amount": "Rs Crore"},
    )
    fig_totals.update_layout(
        height=420, plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
        margin=dict(l=40, r=30, t=20, b=40),
        xaxis=dict(showgrid=False, linecolor="#e5e7eb"),
        yaxis=dict(gridcolor="#f1f5f9"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    st.plotly_chart(fig_totals, use_container_width=True)

    macro = data["Macro"].copy()
    if not macro.empty:
        st.markdown('<div class="section-title">Macroeconomic indicators</div>',
                    unsafe_allow_html=True)
        st.markdown(
            '<div class="section-desc">Key fiscal numbers for the same period '
            '(RBI Annual Report, Ministry of Finance).</div>',
            unsafe_allow_html=True,
        )

        macro_cols = [c for c in macro.columns if c != "Fiscal Year"]
        macro_long = macro.melt(id_vars="Fiscal Year", value_vars=macro_cols,
                                var_name="Indicator", value_name="Value")

        fig_macro = px.line(
            macro_long, x="Fiscal Year", y="Value", color="Indicator",
            markers=True,
            color_discrete_sequence=px.colors.qualitative.Prism,
        )
        fig_macro.update_traces(line=dict(width=2.5), marker=dict(size=9))
        fig_macro.update_layout(
            height=400, plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
            margin=dict(l=40, r=30, t=20, b=40),
            xaxis=dict(showgrid=False, linecolor="#e5e7eb"),
            yaxis=dict(gridcolor="#f1f5f9", title="Value (%)"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        )
        st.plotly_chart(fig_macro, use_container_width=True)

        with st.expander("Show macro data table"):
            st.dataframe(macro, use_container_width=True, hide_index=True)


# Generic page (Ministries / States / Schemes / UTs)

def render_dataset(name: str, entity_col: str):
    # Resolve data source — uploaded file > local path > master workbook
    df, source_label = resolve_sheet_source(name, per_sheet_upload)

    st.markdown(f'<div class="main-title">{name}</div>', unsafe_allow_html=True)
    st.markdown('<div class="tricolor-bar"></div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="main-sub">{name} allocations — FY 2019-20 to FY 2023-24. '
        f'<span style="color:#9ca3af;font-size:0.8rem">Data: {source_label}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if df.empty:
        st.warning("The loaded file has no usable rows. Check the format and try again.")
        return

    # ----- KPI row -----
    totals = year_totals(df)
    total_24 = totals.get("2023-24", np.nan)
    total_23 = totals.get("2022-23", np.nan)
    total_20 = totals.get("2019-20", np.nan)

    yoy = growth(total_23, total_24)
    five_yr = growth(total_20, total_24)
    top = df.sort_values("2023-24", ascending=False).iloc[0]
    top_share = top["2023-24"] / total_24 * 100 if total_24 else 0

    c1, c2, c3, c4 = st.columns(4)
    kpi_card(c1, f"Total — FY 2023-24", fmt_cr(total_24),
             f"Across {len(df)} {name.lower()}")
    kpi_card(c2, "YoY change", f"{yoy:+.1f} %",
             f"vs FY 2022-23 ({fmt_cr(total_23)})", neg=yoy < 0)
    kpi_card(c3, "5-year growth", f"{five_yr:+.1f} %",
             f"Since FY 2019-20 ({fmt_cr(total_20)})", neg=five_yr < 0)
    kpi_card(c4, f"Top {entity_col.lower()} — 2023-24",
             (str(top[entity_col])[:22] + "...")
             if len(str(top[entity_col])) > 25 else str(top[entity_col]),
             f"{fmt_cr(top['2023-24'])} ({top_share:.1f}% of total)")

    # ----- Controls -----
    st.markdown('<div class="section-title">Explore</div>', unsafe_allow_html=True)
    col_a, col_b, col_c = st.columns([2, 2, 1])
    search = col_a.text_input(f"Search {entity_col.lower()}", "")
    year = col_b.selectbox("Year for ranking / donut", YEAR_COLS,
                           index=len(YEAR_COLS) - 1)
    top_k = col_c.slider("Top N", 3, max(3, len(df)), min(10, len(df)))

    filtered = df.copy()
    if search:
        filtered = filtered[
            filtered[entity_col].astype(str).str.contains(search, case=False, na=False)
        ]
    if filtered.empty:
        st.info("Nothing matched your search.")
        return

    tab_rank, tab_trend, tab_donut, tab_heat, tab_table = st.tabs(
        ["Ranking", "Trend", "Share (donut)", "Heatmap", "Data table"]
    )

    with tab_rank:
        rank_df = top_n(filtered, entity_col, year, top_k).iloc[::-1]
        fig_rank = px.bar(
            rank_df, x=year, y=entity_col, orientation="h",
            color=year, color_continuous_scale="Oranges",
            labels={year: "Rs Crore"},
        )
        fig_rank.update_layout(
            height=max(360, 32 * len(rank_df)),
            plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
            margin=dict(l=40, r=30, t=20, b=40),
            xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
            yaxis=dict(showgrid=False),
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig_rank, use_container_width=True)

    with tab_trend:
        default_sel = top_n(filtered, entity_col, "2023-24",
                            min(5, len(filtered)))[entity_col].tolist()
        picked = st.multiselect(
            f"Select {entity_col.lower()}(s) to plot",
            filtered[entity_col].dropna().tolist(),
            default=default_sel,
        )
        if picked:
            st.plotly_chart(entity_trend_fig(filtered, entity_col, picked),
                            use_container_width=True)
        else:
            st.info(f"Pick at least one {entity_col.lower()} to plot.")

    with tab_donut:
        donut_df = top_n(filtered, entity_col, year, top_k)
        others_val = (
            filtered[year].sum() - donut_df[year].sum()
            if len(filtered) > top_k else 0
        )
        if others_val > 0:
            donut_df = pd.concat(
                [donut_df, pd.DataFrame({entity_col: ["Others"], year: [others_val]})],
                ignore_index=True,
            )

        fig_donut = go.Figure(
            data=go.Pie(
                labels=donut_df[entity_col],
                values=donut_df[year],
                hole=0.55,
                textinfo="label+percent",
                marker=dict(
                    colors=px.colors.qualitative.Set2,
                    line=dict(color="#ffffff", width=2),
                ),
                hovertemplate="<b>%{label}</b><br>Rs %{value:,.0f} Cr"
                              "<br>%{percent}<extra></extra>",
            )
        )
        total_val = donut_df[year].sum()
        fig_donut.add_annotation(
            text=f"<b>{fmt_cr(total_val)}</b><br>"
                 f"<span style='font-size:11px;color:#6b7280'>Total {year}</span>",
            x=0.5, y=0.5, showarrow=False,
        )
        fig_donut.update_layout(
            height=460,
            paper_bgcolor="#ffffff",
            margin=dict(l=30, r=30, t=20, b=30),
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    with tab_heat:
        st.caption(
            "Each row is normalised against its own peak year. "
            "Colour shows relative spend intensity — not absolute size."
        )
        st.plotly_chart(
            heatmap_fig(filtered, entity_col, f"{name} — year-on-year intensity"),
            use_container_width=True,
        )

    with tab_table:
        show_cols = [entity_col] + [c for c in filtered.columns if c != entity_col]
        tbl = filtered[show_cols].reset_index(drop=True)
        st.dataframe(tbl, use_container_width=True, hide_index=True)
        st.download_button(
            "Download filtered data (CSV)",
            data=tbl.to_csv(index=False).encode("utf-8"),
            file_name=f"{name.lower().replace(' ', '_')}_filtered.csv",
            mime="text/csv",
        )


# Page: Compare Entities

def render_compare():
    st.markdown('<div class="main-title">Compare entities</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="tricolor-bar"></div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="main-sub">Pick any two entities from any dataset and '
        'put their allocation numbers side by side.</div>',
        unsafe_allow_html=True,
    )

    datasets = {
        "Ministries": ("Ministry", data["Ministries"]),
        "States": ("State", data["States"]),
        "Schemes": ("Scheme", data["Schemes"]),
        "Union Territories": ("Union Territory", data["Union Territories"]),
    }

    col1, col2 = st.columns(2)
    with col1:
        ds1 = st.selectbox("Dataset A", list(datasets.keys()), index=0)
        ent1 = st.selectbox(
            f"{datasets[ds1][0]} A",
            datasets[ds1][1][datasets[ds1][0]].dropna().tolist(),
        )
    with col2:
        ds2 = st.selectbox("Dataset B", list(datasets.keys()), index=1)
        ent2 = st.selectbox(
            f"{datasets[ds2][0]} B",
            datasets[ds2][1][datasets[ds2][0]].dropna().tolist(),
        )

    row1 = datasets[ds1][1][datasets[ds1][1][datasets[ds1][0]] == ent1].iloc[0]
    row2 = datasets[ds2][1][datasets[ds2][1][datasets[ds2][0]] == ent2].iloc[0]

    compare_df = pd.DataFrame(
        {
            "Year": YEAR_COLS,
            f"{ent1} ({ds1})": [row1.get(y, np.nan) for y in YEAR_COLS],
            f"{ent2} ({ds2})": [row2.get(y, np.nan) for y in YEAR_COLS],
        }
    )

    long_df = compare_df.melt(id_vars="Year", var_name="Entity", value_name="Amount")
    fig = px.bar(
        long_df, x="Year", y="Amount", color="Entity",
        barmode="group",
        color_discrete_sequence=["#ff9933", "#138808"],
        labels={"Amount": "Rs Crore"},
    )
    fig.update_layout(
        height=440, plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
        margin=dict(l=40, r=30, t=20, b=40),
        xaxis=dict(showgrid=False, linecolor="#e5e7eb"),
        yaxis=dict(gridcolor="#f1f5f9"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    st.plotly_chart(fig, use_container_width=True)

    total_a = compare_df[f"{ent1} ({ds1})"].sum()
    total_b = compare_df[f"{ent2} ({ds2})"].sum()
    g_a = growth(row1.get("2019-20"), row1.get("2023-24"))
    g_b = growth(row2.get("2019-20"), row2.get("2023-24"))

    c1, c2, c3, c4 = st.columns(4)
    kpi_card(c1, f"{ent1} — 5 yr total", fmt_cr(total_a))
    kpi_card(c2, f"{ent2} — 5 yr total", fmt_cr(total_b))
    kpi_card(c3, f"{ent1} — 5 yr growth", f"{g_a:+.1f} %", neg=g_a < 0)
    kpi_card(c4, f"{ent2} — 5 yr growth", f"{g_b:+.1f} %", neg=g_b < 0)

    with st.expander("Show data table"):
        st.dataframe(compare_df, use_container_width=True, hide_index=True)


# Page: Data Explorer

def render_explorer():
    st.markdown('<div class="main-title">Data explorer</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="tricolor-bar"></div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="main-sub">Raw cleaned sheets from the workbook. '
        'Download any of them as CSV.</div>',
        unsafe_allow_html=True,
    )

    sheet = st.selectbox(
        "Pick a sheet",
        ["Ministries", "States", "Schemes", "Union Territories", "Macro"],
    )
    df = data[sheet]
    st.write(f"**Rows:** {len(df)}  |  **Columns:** {len(df.columns)}")
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button(
        "Download CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name=f"{sheet.lower().replace(' ', '_')}.csv",
        mime="text/csv",
    )


# Router

if page == "Overview":
    render_overview()
elif page == "Ministries":
    render_dataset("Ministries", "Ministry")
elif page == "States":
    render_dataset("States", "State")
elif page == "Schemes":
    render_dataset("Schemes", "Scheme")
elif page == "Union Territories":
    render_dataset("Union Territories", "Union Territory")
elif page == "Compare Entities":
    render_compare()
elif page == "Data Explorer":
    render_explorer()


# Footer

st.markdown("---")
st.markdown(
    '<div class="footer">India Union Budget Dashboard &middot; '
    'Built with Streamlit and Plotly &middot; Data: Ministry of Finance, '
    'indiabudget.gov.in</div>',
    unsafe_allow_html=True,
)
