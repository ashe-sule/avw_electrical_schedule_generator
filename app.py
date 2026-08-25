from __future__ import annotations

from datetime import datetime
from pathlib import Path
import base64
import traceback

import pandas as pd
import streamlit as st
from slugify import slugify
from st_aggrid import AgGrid, DataReturnMode, GridOptionsBuilder, JsCode

from avw_schedule.processor import (
    QuickBooksCatalog,
    SCHEDULE_COLUMNS,
    apply_row_format,
    apply_rows_format,
    delete_row_subtree,
    delete_row_subtrees,
    ensure_schedule_hierarchy,
    insert_schedule_rows,
    make_custom_schedule_row,
    move_row_subtree,
    move_row_subtrees,
    next_group_id_from_df,
    outdent_row,
    process_files,
    refresh_requirement_status,
    refresh_schedule_calculations,
    reparent_row,
    replace_row_from_master,
    review_item_count_snapshot,
    reset_row_requirements_from_master,
    renumber_schedule_orders,
    schedule_row_for_master_excel_row,
    schedule_row_for_quickbooks_excel_row,
    schedule_rows_for_master_subtree_excel_row,
    schedule_rows_for_review_items,
    selected_subtree_ids,
    subtree_descendant_counts,
    write_output_workbook,
)


APP_DIR = Path(__file__).resolve().parent
LOGO_PATH = APP_DIR / "assets" / "avw_logo.png"
BANNER_PATH = APP_DIR / "assets" / "avw_brand_banner.jpg"
QUICKBOOKS_PATH = APP_DIR / "data" / "QuickBooks_Items.xlsx"

st.set_page_config(
    page_title="AVW Demo Web App",
    page_icon="assets/avw_logo.png" if LOGO_PATH.exists() else "⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def _asset_data_uri(path: Path, mime: str) -> str:
    if not path.exists():
        return ""
    encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:{mime};base64,{encoded}"


logo_uri = _asset_data_uri(LOGO_PATH, "image/png")
banner_uri = _asset_data_uri(BANNER_PATH, "image/jpeg")

st.markdown(
    """
    <style>
    :root {
        --avw-blue: #113A8F;
        --avw-navy: #071D49;
        --avw-light-blue: #EAF1FF;
        --avw-silver: #F5F7FB;
        --avw-border: #D8E0EF;
        --avw-red: #C5222E;
        --avw-text: #152238;
    }

    .stApp {
        background: linear-gradient(180deg, #F7FAFF 0%, #FFFFFF 42%, #F8FAFE 100%);
        color: var(--avw-text);
    }

    /* Streamlit can inherit white text from a viewer's dark-mode settings
       while this app forces a light background. Establish a complete light
       contrast baseline for native Streamlit controls and generated text. */
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
    .stApp p, .stApp label, .stApp li,
    .stApp [data-testid="stMarkdownContainer"],
    .stApp [data-testid="stWidgetLabel"] p,
    .stApp [data-testid="stCaptionContainer"] p,
    .stApp [data-testid="stFileUploaderDropzone"] span,
    .stApp [data-testid="stFileUploaderDropzone"] small {
        color: var(--avw-text) !important;
    }

    .stApp input,
    .stApp textarea,
    .stApp [data-baseweb="input"] input,
    .stApp [data-baseweb="textarea"] textarea,
    .stApp [data-baseweb="select"] > div,
    .stApp [data-baseweb="select"] span {
        color: var(--avw-text) !important;
        -webkit-text-fill-color: var(--avw-text) !important;
    }

    .stApp input::placeholder,
    .stApp textarea::placeholder {
        color: #66758A !important;
        opacity: 1 !important;
    }

    .stApp [data-baseweb="select"] > div,
    .stApp [data-baseweb="input"] > div,
    .stApp [data-baseweb="textarea"] > div,
    .stApp [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF !important;
        border-color: var(--avw-border) !important;
    }

    .stApp button:not([kind="primary"]),
    .stApp button:not([kind="primary"]) p,
    .stApp button:not([kind="primary"]) span {
        color: var(--avw-navy) !important;
    }

    .stApp button[kind="primary"],
    .stApp button[kind="primary"] p,
    .stApp button[kind="primary"] span,
    .stApp .stDownloadButton button,
    .stApp .stDownloadButton button p,
    .stApp .stDownloadButton button span {
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
    }

    .stApp [data-baseweb="tab-list"] button,
    .stApp [data-baseweb="tab-list"] button p {
        color: #40516A !important;
    }

    .stApp [data-baseweb="tab-list"] button[aria-selected="true"],
    .stApp [data-baseweb="tab-list"] button[aria-selected="true"] p {
        color: var(--avw-blue) !important;
        font-weight: 800 !important;
    }

    .stApp [data-testid="stAlert"] p,
    .stApp [data-testid="stAlert"] div,
    .stApp [data-testid="stNotification"] p {
        color: #152238 !important;
    }

    .stApp [data-testid="stDataFrame"] {
        color: var(--avw-text) !important;
        background: #FFFFFF !important;
    }

    .block-container {
        padding-top: 1.3rem;
        padding-bottom: 3rem;
        max-width: 1760px;
    }

    [data-testid="stHeader"] {
        background: rgba(255, 255, 255, 0);
    }

    .avw-brand-strip {
        width: 40%;
        height: 110px;
        border-radius: 22px;
        overflow: hidden;
        border: 1px solid rgba(17, 58, 143, 0.18);
        box-shadow: 0 14px 34px rgba(7, 29, 73, 0.16);
        margin-bottom: 16px;
        background: linear-gradient(135deg, #071D49, #113A8F);
    }

    .avw-brand-strip img {
        width: 100%;
        height: 100%;
        object-fit: cover;
        object-position: center;
        display: block;
    }

    .avw-hero {
        display: grid;
        grid-template-columns: minmax(0, 1fr) 310px;
        align-items: stretch;
        gap: 20px;
        padding: 24px;
        border-radius: 24px;
        background:
            radial-gradient(circle at 92% 20%, rgba(197, 34, 46, 0.20), transparent 26%),
            linear-gradient(135deg, #071D49 0%, #113A8F 58%, #1B55B8 100%);
        color: white;
        box-shadow: 0 18px 45px rgba(7, 29, 73, 0.24);
        border: 1px solid rgba(255, 255, 255, 0.18);
        margin-bottom: 18px;
    }

    .avw-hero,
    .avw-hero h1,
    .avw-hero p,
    .avw-hero span,
    .avw-hero div {
        color: #FFFFFF !important;
    }


    .avw-hero-copy {
        display: flex;
        align-items: center;
        gap: 24px;
    }

    .avw-hero-card {
        border-radius: 20px;
        overflow: hidden;
        background: rgba(255, 255, 255, 0.10);
        border: 1px solid rgba(255, 255, 255, 0.20);
        min-height: 170px;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: inset 0 0 0 1px rgba(255,255,255,0.08);
    }

    .avw-hero-card {
    border-radius: 0;
    overflow: hidden;
    background: transparent;
    border: none;
    min-height: 170px;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: none;
}

    .avw-hero-card img {
        width: 0%;
        height: 0%;
        object-fit: cover;
        object-position: center;
        display: block;
    }

    

    @media (max-width: 980px) {
        .avw-hero {
            grid-template-columns: 1fr;
        }
        .avw-hero-card {
            min-height: 120px;
        }
        .avw-title {
            font-size: 32px;
        }
    }

    .avw-logo-shell {
        width: 98px;
        min-width: 98px;
        height: 98px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        background: rgba(255, 255, 255, 0.98);
        box-shadow: inset 0 0 0 1px rgba(17, 58, 143, 0.12), 0 10px 25px rgba(0, 0, 0, 0.22);
        overflow: hidden;
    }

    .avw-logo-shell img {
        width: 90px;
        height: 90px;
        object-fit: contain;
    }

    .avw-logo-fallback {
        color: var(--avw-blue) !important;
        font-weight: 900;
        font-size: 28px;
        letter-spacing: -1px;
    }

    .avw-hero .avw-logo-fallback {
        color: var(--avw-blue) !important;
    }

    .avw-kicker {
        text-transform: uppercase;
        letter-spacing: 1.6px;
        font-size: 12px;
        font-weight: 800;
        opacity: 0.92;
        margin-bottom: 8px;
    }

    .avw-title {
        font-size: 40px;
        line-height: 1.08;
        font-weight: 900;
        margin: 0;
        letter-spacing: -0.8px;
    }

    .avw-subtitle {
        font-size: 16px;
        line-height: 1.45;
        color: rgba(255,255,255,0.86);
        margin-top: 10px;
        max-width: 830px;
    }

    .avw-chip-row {
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        margin-top: 16px;
    }

    .avw-chip {
        padding: 7px 12px;
        border-radius: 999px;
        background: rgba(255,255,255,0.14);
        border: 1px solid rgba(255,255,255,0.24);
        color: white;
        font-size: 13px;
        font-weight: 700;
    }

    .avw-panel {
        background: #FFFFFF;
        border: 1px solid var(--avw-border);
        border-radius: 18px;
        padding: 20px;
        box-shadow: 0 10px 28px rgba(7, 29, 73, 0.07);
    }

    .avw-section-title {
        color: var(--avw-navy);
        font-weight: 900;
        font-size: 24px;
        margin: 4px 0 4px 0;
    }

    .avw-section-caption {
        color: #536274;
        margin: 0 0 12px 0;
        font-size: 14px;
    }

    .avw-mini-card {
        background: #FFFFFF;
        border: 1px solid var(--avw-border);
        border-radius: 16px;
        padding: 15px 16px;
        min-height: 112px;
        box-shadow: 0 8px 18px rgba(7, 29, 73, 0.05);
    }

    .avw-mini-card b {
        color: var(--avw-navy) !important;
        font-size: 15px;
    }

    .avw-mini-card p {
        color: #46566C !important;
        font-size: 13px;
        margin: 6px 0 0 0;
        line-height: 1.35;
    }

    .avw-step {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: var(--avw-blue);
        color: white;
        font-size: 14px;
        font-weight: 900;
        margin-right: 8px;
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 12px !important;
        font-weight: 800 !important;
        min-height: 42px;
    }

    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--avw-blue), #0B2E6E) !important;
        border: 0 !important;
        box-shadow: 0 8px 20px rgba(17, 58, 143, 0.22);
    }

    div[data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid var(--avw-border);
        border-radius: 16px;
        padding: 14px 16px;
        box-shadow: 0 8px 18px rgba(7, 29, 73, 0.05);
    }

    div[data-testid="stMetricLabel"] p {
        color: #113A8F;
        font-weight: 700;
    }

    div[data-testid="stMetricValue"] {
        color: var(--avw-navy);
        font-weight: 900;
    }

    div[data-testid="stVerticalBlockBorderWrapper"]:has(.avw-toolbar-marker) {
        position: sticky;
        top: 2.7rem;
        z-index: 50;
        background: rgba(255, 255, 255, 0.98);
        border: 1px solid #B8C8E2;
        box-shadow: 0 12px 28px rgba(7, 29, 73, 0.14);
        backdrop-filter: blur(10px);
    }

    .avw-toolbar-marker {
        display: block;
        height: 0;
    }

    [data-testid="stDataEditor"] {
        background: #FFFFFF !important;
        border: 1px solid var(--avw-border);
        border-radius: 12px;
    }

    .avw-footer-note {
        color: #46566C !important;
        font-size: 13px;
        text-align: center;
        padding: 18px 0 0 0;
    }

    </style>
    """,
    unsafe_allow_html=True,
)

logo_html = (
    f'<img src="{logo_uri}" alt="A.V.W. logo" />'
    if logo_uri
    else '<div class="avw-logo-fallback">AVW</div>'
)

brand_strip_html = f'<div class="avw-brand-strip"><img src="{banner_uri}" alt="A.V.W. built on what works banner" /></div>' if banner_uri else ""
hero_card_html = f'<div class="avw-hero-card"><img src="{banner_uri}" alt="A.V.W. built on what works" /></div>' if banner_uri else '<div class="avw-hero-card"><div class="avw-logo-fallback">Built on what works</div></div>'

st.markdown(
    f"""
    {brand_strip_html}
    <div class="avw-hero">
        <div class="avw-hero-copy">
            <div class="avw-logo-shell">{logo_html}</div>
            <div>
                <div class="avw-kicker">A.V.W. Equipment Co., Inc. • Demo Web App</div>
                <h1 class="avw-title">Electrical Schedule Generator</h1>
                <div class="avw-subtitle">
                    Built for AVW car wash equipment projects: upload a QuickBooks quote PDF and the Master List,
                    then use strict Master-first matching with the bundled QuickBooks fallback before review and export.
                </div>
                <div class="avw-chip-row">
                    <div class="avw-chip">Quote PDF → Schedule</div>
                    <div class="avw-chip">Editable Order</div>
                    <div class="avw-chip">Excel Export</div>
                </div>
            </div>
        </div>
        {hero_card_html}
    </div>
    """,
    unsafe_allow_html=True,
)

info_cols = st.columns(3)
with info_cols[0]:
    st.markdown(
        """
        <div class="avw-mini-card">
            <b>1. Upload</b>
            <p>Add the customer quote PDF and the Master List workbook. The app keeps the quote order as the starting point.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with info_cols[1]:
    st.markdown(
        """
        <div class="avw-mini-card">
            <b>2. Review & Edit</b>
            <p>Move, delete, or insert parent groups while keeping all nested child rows together.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with info_cols[2]:
    st.markdown(
        """
        <div class="avw-mini-card">
            <b>3. Export</b>
            <p>Download a clean Excel schedule with formulas, totals, review items, and quote extract sheets.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)


def reset_draft() -> None:
    for key in [
        "meta",
        "quote_df",
        "schedule_df",
        "review_df",
        "master",
        "parent_catalog_df",
        "component_catalog_df",
        "quickbooks",
        "quickbooks_catalog_df",
        "rules",
        "output_xlsx",
        "schedule_history",
        "schedule_grid_revision",
        "schedule_grid_selected_ids",
        "review_grid_selected_indices",
        "review_grid_revision",
        "workspace_master_selected_excel_row",
        "workspace_master_grid_revision",
        "show_master_add_panel",
        "workspace_quickbooks_selected_excel_row",
        "show_quickbooks_add_panel",
        "review_insert_notice",
    ]:
        st.session_state.pop(key, None)


def store_result(result: dict) -> None:
    for key in [
        "meta",
        "quote_df",
        "schedule_df",
        "review_df",
        "master",
        "parent_catalog_df",
        "component_catalog_df",
        "quickbooks",
        "quickbooks_catalog_df",
        "rules",
        "output_xlsx",
    ]:
        st.session_state[key] = result.get(key)


def group_options(schedule_df: pd.DataFrame) -> dict[str, str]:
    """Return {label: group_id} for parent/group controls."""
    if schedule_df is None or schedule_df.empty or "Group ID" not in schedule_df.columns:
        return {}
    labels: dict[str, str] = {}
    seen: set[str] = set()
    for _, row in schedule_df.iterrows():
        gid = str(row.get("Group ID", ""))
        if not gid or gid in seen:
            continue
        seen.add(gid)
        labels[f"{gid} | Order {row.get('Order', '')} | {row.get('Item', '')} | {str(row.get('Description', ''))[:70]}"] = gid
    return labels


def row_options(schedule_df: pd.DataFrame) -> dict[str, str]:
    """Return readable outline labels mapped to stable Row IDs."""
    if schedule_df is None or schedule_df.empty:
        return {}
    labels: dict[str, str] = {}
    for position, (_, row) in enumerate(schedule_df.iterrows(), start=1):
        code = str(row.get("Nested Code", ""))
        outline = str(row.get("Order", "")) if code.lower() == "parent" else f"{row.get('Order', '')}{code}"
        depth = int(row.get("Depth", 0) or 0)
        indent = "› " * depth
        label = f"{position:03d} | {outline:>5} | {indent}{row.get('Item', '')} | {str(row.get('Description', ''))[:64]}"
        labels[label] = str(row.get("Row ID", ""))
    return labels


def merge_grid_edits(
    schedule_df: pd.DataFrame,
    visible_grid_df: pd.DataFrame,
    editable_columns: list[str],
) -> pd.DataFrame:
    """Merge edited visible rows back into the complete, possibly filtered tree."""
    if visible_grid_df is None or visible_grid_df.empty or "Row ID" not in visible_grid_df.columns:
        return schedule_df
    updated = schedule_df.copy()
    for column in editable_columns:
        if column in updated.columns:
            updated[column] = updated[column].astype("object")
    row_index = {
        str(row_id): index
        for index, row_id in updated["Row ID"].astype(str).items()
    }
    for _, visible_row in visible_grid_df.iterrows():
        target_index = row_index.get(str(visible_row.get("Row ID", "")))
        if target_index is None:
            continue
        for column in editable_columns:
            if column in visible_grid_df.columns and column in updated.columns:
                value = visible_row.get(column, "")
                updated.at[target_index, column] = "" if pd.isna(value) else value
    return updated


def schedules_have_same_values(
    left: pd.DataFrame,
    right: pd.DataFrame,
    columns: list[str],
) -> bool:
    if left is None or right is None or len(left) != len(right):
        return False
    usable = [column for column in columns if column in left.columns and column in right.columns]
    if not usable:
        return True
    left_values = left[usable].fillna("").astype(str).reset_index(drop=True)
    right_values = right[usable].fillna("").astype(str).reset_index(drop=True)
    return left_values.equals(right_values)


def clear_grid_selection() -> None:
    st.session_state["schedule_grid_selected_ids"] = []
    st.session_state["schedule_grid_revision"] = int(
        st.session_state.get("schedule_grid_revision", 0)
    ) + 1


def preserve_grid_selection(row_ids: list[str]) -> None:
    """Refresh structural grid data while reselecting the same stable rows."""
    st.session_state["schedule_grid_selected_ids"] = [
        str(row_id) for row_id in row_ids if str(row_id).strip()
    ]
    st.session_state["schedule_grid_revision"] = int(
        st.session_state.get("schedule_grid_revision", 0)
    ) + 1


@st.cache_data(show_spinner=False, max_entries=12)
def cached_output_workbook_bytes(
    meta: dict,
    quote_df: pd.DataFrame,
    schedule_df: pd.DataFrame,
    review_df: pd.DataFrame,
    rules: dict,
) -> bytes:
    """Regenerate Excel only when draft content changes, not on row selection."""
    output = write_output_workbook(meta, quote_df, schedule_df, review_df, rules=rules)
    return output.getvalue()


@st.cache_resource(show_spinner=False)
def cached_quickbooks_catalog(path_text: str, modified_ns: int) -> QuickBooksCatalog:
    """Parse and index the bundled QuickBooks workbook once per app version."""
    del modified_ns
    return QuickBooksCatalog.from_excel(path_text)


def commit_schedule(new_df: pd.DataFrame) -> None:
    """Save one reversible draft edit and keep a short in-session undo history."""
    current = st.session_state.get("schedule_df")
    history = st.session_state.setdefault("schedule_history", [])
    if isinstance(current, pd.DataFrame):
        history.append(current.copy(deep=True))
        del history[:-20]
    st.session_state["schedule_df"] = new_df.reset_index(drop=True)


def undo_schedule() -> bool:
    history = st.session_state.get("schedule_history", [])
    if not history:
        return False
    st.session_state["schedule_df"] = history.pop()
    return True


st.markdown('<div class="avw-panel">', unsafe_allow_html=True)
st.markdown('<div class="avw-section-title"><span class="avw-step">1</span>Upload project files</div>', unsafe_allow_html=True)
st.markdown('<p class="avw-section-caption">Only two uploads are needed: the quote PDF and Master List. The QuickBooks fallback catalog is bundled with the app.</p>', unsafe_allow_html=True)

left, middle, right = st.columns([1, 1, 1])
with left:
    quote_pdf = st.file_uploader("Quote PDF / QuickBooks PDF", type=["pdf"], help="Upload the customer quote/invoice PDF.")
with middle:
    master_excel = st.file_uploader("Master List Excel", type=["xlsx", "xlsm"], help="Upload the AVW engineering Master List workbook.")
with right:
    non_schedule_rules = st.file_uploader(
        "Known non-schedule list (optional CSV)",
        type=["csv"],
        help="Columns: item, reason, enabled. These items create no schedule rows and always remain visible in Review.",
    )

button_cols = st.columns([1.4, 1, 4.6])
with button_cols[0]:
    run = st.button("Generate Draft", type="primary", disabled=not (quote_pdf and master_excel), use_container_width=True)
with button_cols[1]:
    if st.button("Clear", use_container_width=True):
        reset_draft()
        st.rerun()
with button_cols[2]:
    st.caption("Master-first matching, bundled QuickBooks fallback, then review, edit, and export.")

st.markdown('</div>', unsafe_allow_html=True)

if run:
    reset_draft()
    with st.spinner("Building the draft with Master-first and QuickBooks fallback matching..."):
        try:
            if not QUICKBOOKS_PATH.exists():
                raise FileNotFoundError(
                    "The bundled QuickBooks item catalog is missing from the app package."
                )
            quickbooks_catalog = cached_quickbooks_catalog(
                str(QUICKBOOKS_PATH), QUICKBOOKS_PATH.stat().st_mtime_ns
            )
            result = process_files(
                quote_pdf=quote_pdf,
                master_excel=master_excel,
                quickbooks_catalog=quickbooks_catalog,
                alias_rules=None,
                replacement_rules=None,
                special_rules=None,
                ignore_rules=non_schedule_rules,
            )
            store_result(result)
        except Exception:
            st.error("The demo app hit an error while processing the files.")
            st.code(traceback.format_exc())
            st.stop()

if "schedule_df" not in st.session_state:
    st.info("Upload the Quote PDF and Master List Excel, then click **Generate Draft**.")
    st.markdown(
        '<div class="avw-footer-note">Made with ❤️ by Aroon Kumar</div>',
        unsafe_allow_html=True,
    )
    st.stop()

meta = st.session_state["meta"] or {}
quote_df = st.session_state["quote_df"]
schedule_df = st.session_state["schedule_df"]
review_df = st.session_state["review_df"]
master = st.session_state["master"]
parent_catalog_df = st.session_state["parent_catalog_df"]
component_catalog_df = st.session_state.get("component_catalog_df")
quickbooks = st.session_state.get("quickbooks")
quickbooks_catalog_df = st.session_state.get("quickbooks_catalog_df")
rules = st.session_state.get("rules") or {}

st.markdown("<br>", unsafe_allow_html=True)
st.success("Draft schedule is ready. Review, edit, then export the Excel workbook.")

kpis = st.columns(5)
kpis[0].metric("Customer", meta.get("customer") or "Unknown")
kpis[1].metric("Invoice", meta.get("invoice_number") or "Unknown")
kpis[2].metric("Country", meta.get("country") or "Unknown")
kpis[3].metric("Schedule Rows", len(schedule_df))
kpis[4].metric("Review Items", len(review_df))

st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<div class="avw-section-title"><span class="avw-step">2</span>Edit draft schedule</div>', unsafe_allow_html=True)
st.markdown('<p class="avw-section-caption">Parent groups move with their nested child rows so the engineering order stays clean.</p>', unsafe_allow_html=True)

edit_tab, preview_tab, review_tab, quote_tab = st.tabs([
    "Edit Schedule Tree",
    "Electrical Schedule Preview",
    "Review Items",
    "Quote Extract",
])

with edit_tab:
    schedule_df = refresh_requirement_status(ensure_schedule_hierarchy(schedule_df), master)
    st.session_state["schedule_df"] = schedule_df
    if schedule_df.empty:
        st.warning("No schedule rows were generated. Use Review Items for unresolved quote lines.")
    else:
        st.markdown("#### Engineering schedule workspace")
        st.caption(
            "Select one or more rows directly in the grid. Parent rows are blue and bold; "
            "manual highlights are yellow. Structural actions keep complete subtrees together, "
            "and project numbers plus A / AA / AAA codes regenerate automatically."
        )

        filter_cols = st.columns([1.25, 2.4, 2.1])
        view_mode = filter_cols[0].radio(
            "Rows shown",
            ["All rows", "Parents only"],
            horizontal=True,
            key="schedule_view_mode",
        )
        schedule_search = filter_cols[1].text_input(
            "Find in schedule",
            placeholder="Part #, description, project number, or nested code",
            key="schedule_workspace_search",
        )
        filter_cols[2].markdown(
            "<div style='padding-top:28px'><b style='color:#071D49'>Blue</b> = parent/standalone &nbsp; "
            "<b style='background:#FFF2B2;color:#071D49;padding:3px 7px;border-radius:5px'>Yellow</b> = manual highlight &nbsp; "
            "<b style='background:#F4CCCC;color:#071D49;padding:3px 7px;border-radius:5px'>Red</b> = added from Review</div>",
            unsafe_allow_html=True,
        )

        visible_schedule = schedule_df.copy()
        if view_mode == "Parents only":
            visible_schedule = visible_schedule[visible_schedule["Depth"].fillna(0).astype(int).eq(0)]
        if schedule_search.strip():
            query = schedule_search.strip()
            search_columns = ["Order", "Nested Code", "Item", "Description"]
            search_mask = pd.Series(False, index=visible_schedule.index)
            for column in search_columns:
                if column in visible_schedule.columns:
                    search_mask |= visible_schedule[column].astype(str).str.contains(
                        query, case=False, na=False, regex=False
                    )
            visible_schedule = visible_schedule[search_mask]

        child_counts = subtree_descendant_counts(schedule_df)

        grid_df = visible_schedule.copy()
        grid_df.insert(
            0,
            "Row",
            [
                f"PARENT (+{child_counts.get(str(row.get('Row ID')), 0)})"
                if int(row.get("Depth", 0) or 0) == 0
                else f"↳ {row.get('Nested Code', '')}"
                for _, row in grid_df.iterrows()
            ],
        )
        grid_df["_Parent Display"] = grid_df["Depth"].fillna(0).astype(int).eq(0)

        grid_columns = [
            "Row", "Order", "Nested Code", "Item", "Description", "Count Category", "#", "Qty",
            "HP", "Phase", "Volts", "Amps", "C.B.", "Air Port", "Cold Water",
            "Hot Water", "Reclaim Water", "Gas BTUH", "Requirements Status", "Requirements Source",
            "Row ID", "Parent Row ID", "Depth", "Master Excel Row", "Bold", "Italic", "Underline",
            "Highlight", "Review Added", "_Parent Display",
        ]
        grid_columns = [column for column in grid_columns if column in grid_df.columns]
        grid_df = grid_df[grid_columns].reset_index(drop=True)
        boolean_grid_columns = [
            column for column in [
                "Bold", "Italic", "Underline", "Highlight", "Review Added", "_Parent Display"
            ]
            if column in grid_df.columns
        ]
        for column in boolean_grid_columns:
            grid_df[column] = grid_df[column].map(
                lambda value: value is True or str(value).strip().lower() in {"true", "1", "yes"}
            )
        for column in [
            column for column in grid_df.columns
            if column not in boolean_grid_columns + ["Depth"]
        ]:
            grid_df[column] = grid_df[column].fillna("").astype(str)
        editable_value_columns = [
            "Item", "Description", "Count Category", "Qty", "HP", "Phase", "Volts", "Amps",
            "C.B.", "Air Port", "Cold Water", "Hot Water", "Reclaim Water", "Gas BTUH",
        ]
        editable_value_columns = [column for column in editable_value_columns if column in grid_df.columns]

        row_style = JsCode(
            """
            function(params) {
                const truthy = (value) => value === true || String(value).toLowerCase() === 'true';
                const style = { color: '#152238' };
                if (truthy(params.data['_Parent Display'])) {
                    style.backgroundColor = '#D6E3F7';
                    style.fontWeight = '700';
                }
                if (truthy(params.data['Highlight'])) {
                    style.backgroundColor = '#FFF2B2';
                }
                if (truthy(params.data['Review Added'])) {
                    style.backgroundColor = '#F4CCCC';
                }
                if (truthy(params.data['Bold'])) style.fontWeight = '700';
                if (truthy(params.data['Italic'])) style.fontStyle = 'italic';
                if (truthy(params.data['Underline'])) style.textDecoration = 'underline';
                return style;
            }
            """
        )
        schedule_tooltip = JsCode(
            r"""
            function(params) {
                const row = params.data || {};
                const description = String(row['Description'] || '').trim();
                const fields = [
                    ['HP', row['HP']], ['PHASE', row['Phase']], ['VOLTS', row['Volts']],
                    ['AMPS', row['Amps']], ['CB', row['C.B.']], ['AIR PORT', row['Air Port']],
                    ['COLD WATER', row['Cold Water']], ['HOT WATER', row['Hot Water']],
                    ['RECLAIM WATER', row['Reclaim Water']], ['GAS BTUH', row['Gas BTUH']]
                ].filter(pair => {
                    const value = String(pair[1] == null ? '' : pair[1]).trim();
                    return value && value !== '-';
                }).map(pair => pair[0] + ': ' + String(pair[1]));
                return description + (fields.length ? '\n\n' + fields.join(' | ') : '\n\nNo electrical or utility requirements listed');
            }
            """
        )
        grid_return = JsCode(
            """
            function({streamlitRerunEventTriggerName, eventData}) {
                const api = eventData.api;
                const selectedRowIds = api.getSelectedRows().map(
                    row => String(row['Row ID'] || '')
                ).filter(Boolean);
                const editedRows = [];
                if (streamlitRerunEventTriggerName === 'cellValueChanged') {
                    api.forEachNode(node => editedRows.push(node.data));
                }
                return {
                    trigger: streamlitRerunEventTriggerName,
                    selectedRowIds: selectedRowIds,
                    editedRows: editedRows
                };
            }
            """
        )
        grid_builder = GridOptionsBuilder.from_dataframe(grid_df)
        grid_builder.configure_default_column(
            editable=False,
            sortable=False,
            filter=False,
            resizable=True,
            wrapText=True,
        )
        remembered_grid_ids = set(st.session_state.get("schedule_grid_selected_ids", []))
        preselected_grid_rows = [
            position for position, row_id in enumerate(grid_df.get("Row ID", pd.Series(dtype=str)).astype(str))
            if row_id in remembered_grid_ids
        ]
        grid_builder.configure_selection(
            selection_mode="multiple",
            use_checkbox=True,
            header_checkbox=True,
            header_checkbox_filtered_only=True,
            pre_selected_rows=preselected_grid_rows,
            rowMultiSelectWithClick=True,
            suppressRowDeselection=False,
        )
        grid_builder.configure_column("Row", pinned="left", width=138, editable=False)
        grid_builder.configure_column(
            "Item", pinned="left", width=185, editable=True,
            tooltipValueGetter=schedule_tooltip,
        )
        grid_builder.configure_column(
            "Description", width=370, editable=True, wrapText=True,
            tooltipValueGetter=schedule_tooltip,
        )
        grid_builder.configure_column("Order", width=82, editable=False)
        grid_builder.configure_column("Nested Code", width=105, editable=False)
        grid_builder.configure_column("#", width=65, editable=False)
        grid_builder.configure_column("Requirements Status", width=185, editable=False)
        grid_builder.configure_column("Requirements Source", width=210, editable=False)
        for column in editable_value_columns:
            if column not in {"Item", "Description"}:
                grid_builder.configure_column(column, editable=True, width=105)
        for hidden_column in [
            "Row ID", "Parent Row ID", "Depth", "Master Excel Row", "Bold", "Italic",
            "Underline", "Highlight", "Review Added", "_Parent Display",
        ]:
            if hidden_column in grid_df.columns:
                grid_builder.configure_column(hidden_column, hide=True)
        grid_builder.configure_grid_options(
            getRowStyle=row_style,
            rowHeight=38,
            headerHeight=40,
            tooltipShowDelay=2000,
            tooltipHideDelay=12000,
            suppressDragLeaveHidesColumns=True,
        )

        toolbar_placeholder = st.empty()
        grid_revision = int(st.session_state.get("schedule_grid_revision", 0))
        grid_response = AgGrid(
            grid_df,
            gridOptions=grid_builder.build(),
            data_return_mode=DataReturnMode.CUSTOM,
            custom_jscode_for_grid_return=grid_return,
            update_on=[("cellValueChanged", 250), ("selectionChanged", 450)],
            allow_unsafe_jscode=True,
            fit_columns_on_grid_load=False,
            height=760,
            theme="light",
            key=f"engineering_schedule_grid_{grid_revision}",
            custom_css={
                ".ag-header": {"background-color": "#EAF1FF", "color": "#071D49", "font-weight": "700"},
                ".ag-root-wrapper": {"border": "1px solid #B8C8E2", "border-radius": "12px"},
                ".ag-row-selected": {"outline": "2px solid #2A62C9 !important", "outline-offset": "-2px"},
            },
        )

        returned_grid = grid_response.get("editedRows", [])
        if not isinstance(returned_grid, pd.DataFrame):
            returned_grid = pd.DataFrame(returned_grid)
        merged_schedule = merge_grid_edits(schedule_df, returned_grid, editable_value_columns)
        if not schedules_have_same_values(schedule_df, merged_schedule, editable_value_columns):
            merged_schedule = refresh_requirement_status(
                refresh_schedule_calculations(merged_schedule), master
            )
            commit_schedule(merged_schedule)
            schedule_df = merged_schedule
            st.session_state["schedule_df"] = merged_schedule

        response_selected_ids = grid_response.get("selectedRowIds", None)
        if response_selected_ids is not None:
            st.session_state["schedule_grid_selected_ids"] = [
                str(row_id) for row_id in response_selected_ids if str(row_id).strip()
            ]
        selected_ids = list(st.session_state.get("schedule_grid_selected_ids", []))
        valid_selected_ids = [
            row_id for row_id in selected_ids
            if row_id in set(schedule_df["Row ID"].astype(str))
        ]
        affected_ids = selected_subtree_ids(schedule_df, valid_selected_ids)
        row_map = row_options(schedule_df)
        row_labels = {row_id: label for label, row_id in row_map.items()}

        with toolbar_placeholder.container(border=True):
            st.markdown('<span class="avw-toolbar-marker"></span>', unsafe_allow_html=True)
            action_cols = st.columns([1.25, 1, 1, 1.1, 1.1, 1, 1, 1, 1, 0.85])
            action_cols[0].markdown(
                f"**{len(valid_selected_ids)} selected**  \n"
                f"{len(affected_ids)} row{'s' if len(affected_ids) != 1 else ''} with descendants"
            )

            with action_cols[1]:
                step_cols = st.columns(2)
                step_up = step_cols[0].button(
                    "↑", help="Move the selected row/group one step up",
                    disabled=len(valid_selected_ids) != 1,
                    use_container_width=True,
                )
                step_down = step_cols[1].button(
                    "↓", help="Move the selected row/group one step down",
                    disabled=len(valid_selected_ids) != 1,
                    use_container_width=True,
                )
                if step_up or step_down:
                    commit_schedule(move_row_subtree(
                        schedule_df, valid_selected_ids[0], -1 if step_up else 1
                    ))
                    preserve_grid_selection(valid_selected_ids)
                    st.rerun()

            with action_cols[1].popover("Move / place", use_container_width=True):
                if not valid_selected_ids:
                    st.info("Select one or more rows in the table first.")
                else:
                    moving_set = set(affected_ids)
                    destination_df = schedule_df[~schedule_df["Row ID"].astype(str).isin(moving_set)]
                    if view_mode == "Parents only":
                        destination_df = destination_df[destination_df["Depth"].fillna(0).astype(int).eq(0)]
                    destination_options = row_options(destination_df)
                    destination_labels = list(destination_options.keys())
                    placement = st.selectbox(
                        "Place selected rows",
                        ["Before destination", "After destination", "As last child", "At bottom"],
                        key="batch_move_mode",
                    )
                    destination_label = ""
                    if placement != "At bottom":
                        destination_label = st.selectbox(
                            "Destination row",
                            destination_labels,
                            key="batch_move_destination",
                        ) if destination_labels else ""
                    if st.button(
                        "Move selected rows",
                        type="primary",
                        use_container_width=True,
                        disabled=placement != "At bottom" and not destination_label,
                    ):
                        mode_map = {
                            "Before destination": "before",
                            "After destination": "after",
                            "As last child": "child",
                            "At bottom": "bottom",
                        }
                        target_id = destination_options.get(destination_label, "")
                        commit_schedule(move_row_subtrees(
                            schedule_df, valid_selected_ids, target_id, mode_map[placement]
                        ))
                        preserve_grid_selection(valid_selected_ids)
                        st.rerun()

            with action_cols[2].popover("Hierarchy", use_container_width=True):
                if len(valid_selected_ids) != 1:
                    st.info("Select exactly one row to nest or outdent it.")
                else:
                    active_row_id = valid_selected_ids[0]
                    invalid_targets = set(selected_subtree_ids(schedule_df, [active_row_id]))
                    parent_df = schedule_df[~schedule_df["Row ID"].astype(str).isin(invalid_targets)]
                    parent_options = row_options(parent_df)
                    parent_label = st.selectbox(
                        "New parent row",
                        list(parent_options.keys()),
                        key="workspace_nest_target",
                    ) if parent_options else ""
                    if st.button("Make last child", type="primary", disabled=not parent_label):
                        commit_schedule(reparent_row(
                            schedule_df, active_row_id, parent_options.get(parent_label, "")
                        ))
                        preserve_grid_selection(valid_selected_ids)
                        st.rerun()
                    if st.button("Outdent one level", use_container_width=True):
                        commit_schedule(outdent_row(schedule_df, active_row_id))
                        preserve_grid_selection(valid_selected_ids)
                        st.rerun()

            master_panel_open = action_cols[3].toggle(
                "Add from Master",
                key="show_master_add_panel",
                help="Keeps the Master search open while you search, select, and insert.",
            )

            quickbooks_panel_open = action_cols[4].toggle(
                "Add from QuickBooks",
                key="show_quickbooks_add_panel",
                help="Keeps the QuickBooks search open while you search, select, and insert.",
            )

            with action_cols[5].popover("Add custom", use_container_width=True):
                custom_item = st.text_input("Custom item / part #", key="workspace_custom_item")
                custom_description = st.text_input("Description", key="workspace_custom_description")
                custom_mode_label = st.selectbox(
                    "Place",
                    ["After selected", "Before selected", "As child", "Bottom"],
                    key="workspace_custom_insert_mode",
                )
                custom_anchor_required = custom_mode_label != "Bottom"
                custom_anchor_ok = len(valid_selected_ids) == 1 or not custom_anchor_required
                if custom_anchor_required and not custom_anchor_ok:
                    st.caption("Select exactly one schedule row as the insertion anchor.")
                if st.button(
                    "Add custom row",
                    type="primary",
                    disabled=not custom_anchor_ok,
                    use_container_width=True,
                ):
                    custom_df = pd.DataFrame(
                        [make_custom_schedule_row(custom_item, custom_description)],
                        columns=SCHEDULE_COLUMNS,
                    )
                    mode_map = {
                        "After selected": "after", "Before selected": "before",
                        "As child": "child", "Bottom": "bottom",
                    }
                    anchor_id = valid_selected_ids[0] if len(valid_selected_ids) == 1 else ""
                    commit_schedule(insert_schedule_rows(
                        schedule_df, custom_df,
                        mode=mode_map[custom_mode_label], target_row_id=anchor_id,
                    ))
                    clear_grid_selection()
                    st.rerun()

            with action_cols[6].popover("Format", use_container_width=True):
                st.caption("Formatting applies to the selected rows only and is exported to Excel.")
                format_actions = st.multiselect(
                    "Changes to apply",
                    [
                        "Bold", "Remove bold", "Italic", "Remove italic",
                        "Underline", "Remove underline", "Highlight", "Remove highlight",
                    ],
                    key="workspace_format_actions",
                )
                contradictory = any(
                    positive in format_actions and negative in format_actions
                    for positive, negative in [
                        ("Bold", "Remove bold"),
                        ("Italic", "Remove italic"),
                        ("Underline", "Remove underline"),
                        ("Highlight", "Remove highlight"),
                    ]
                )
                if contradictory:
                    st.error("Choose either apply or remove for each style, not both.")
                if st.button(
                    "Apply formatting",
                    type="primary",
                    use_container_width=True,
                    disabled=not valid_selected_ids or not format_actions or contradictory,
                ):
                    def requested(positive: str, negative: str):
                        if positive in format_actions:
                            return True
                        if negative in format_actions:
                            return False
                        return None

                    commit_schedule(apply_rows_format(
                        schedule_df,
                        valid_selected_ids,
                        bold=requested("Bold", "Remove bold"),
                        italic=requested("Italic", "Remove italic"),
                        underline=requested("Underline", "Remove underline"),
                        highlight=requested("Highlight", "Remove highlight"),
                    ))
                    preserve_grid_selection(valid_selected_ids)
                    st.rerun()

            with action_cols[7].popover("Master values", use_container_width=True):
                if len(valid_selected_ids) != 1:
                    st.info("Select exactly one row to inspect or reset its requirements.")
                else:
                    active_record = schedule_df[
                        schedule_df["Row ID"].astype(str).eq(valid_selected_ids[0])
                    ].iloc[0]
                    st.caption(str(active_record.get("Requirements Source", "No Master source")))
                    st.write(f"Status: **{active_record.get('Requirements Status', '')}**")
                    if st.button(
                        "Reset from Master",
                        disabled=not active_record.get("Master Excel Row"),
                        use_container_width=True,
                    ):
                        commit_schedule(reset_row_requirements_from_master(
                            schedule_df, valid_selected_ids[0], master
                        ))
                        preserve_grid_selection(valid_selected_ids)
                        st.rerun()

            with action_cols[8].popover("Delete", use_container_width=True):
                if not valid_selected_ids:
                    st.info("Select one or more rows first.")
                else:
                    st.warning(
                        f"You selected {len(valid_selected_ids)} row(s). Including nested children, "
                        f"{len(affected_ids)} row(s) will be removed."
                    )
                    confirm_delete = st.checkbox(
                        "I understand these rows will be removed",
                        key="workspace_confirm_delete",
                    )
                    if st.button(
                        "Confirm delete",
                        type="primary",
                        use_container_width=True,
                        disabled=not confirm_delete,
                    ):
                        commit_schedule(delete_row_subtrees(schedule_df, valid_selected_ids))
                        clear_grid_selection()
                        st.rerun()

            if action_cols[9].button(
                "Undo",
                use_container_width=True,
                disabled=not st.session_state.get("schedule_history"),
            ):
                if undo_schedule():
                    clear_grid_selection()
                    st.rerun()

            if master_panel_open:
                st.markdown("##### Add from Master")
                st.caption(
                    "This panel stays open while you search. Hover over an item or description for two seconds "
                    "to see the full description and its available requirements."
                )
                if component_catalog_df is None or component_catalog_df.empty:
                    st.warning("No searchable Master rows were found.")
                else:
                    master_search = st.text_input(
                        "Find Master item",
                        placeholder="Part # or description",
                        key="workspace_master_search",
                    )
                    filtered_master = component_catalog_df.copy()
                    if master_search.strip():
                        master_mask = (
                            filtered_master["Item"].astype(str).str.contains(
                                master_search, case=False, na=False, regex=False
                            )
                            | filtered_master["Description"].astype(str).str.contains(
                                master_search, case=False, na=False, regex=False
                            )
                        )
                        filtered_master = filtered_master[master_mask]
                    filtered_master = filtered_master.head(100).reset_index(drop=True)
                    if filtered_master.empty:
                        st.info("No Master rows match that search.")
                    else:
                        master_tooltip = JsCode(
                            r"""
                            function(params) {
                                const row = params.data || {};
                                return String(row['Description'] || '') + '\n\n' +
                                    String(row['Requirements'] || 'No electrical or utility requirements listed');
                            }
                            """
                        )
                        master_return = JsCode(
                            """
                            function({eventData}) {
                                const api = eventData.api;
                                return {
                                    selectedExcelRows: api.getSelectedRows().map(
                                        row => String(row['Excel Row'] || '')
                                    ).filter(Boolean)
                                };
                            }
                            """
                        )
                        remembered_excel_row = str(
                            st.session_state.get("workspace_master_selected_excel_row", "")
                        )
                        preselected_master_rows = [
                            position for position, value in enumerate(
                                filtered_master["Excel Row"].astype(str).tolist()
                            ) if value == remembered_excel_row
                        ]
                        master_builder = GridOptionsBuilder.from_dataframe(
                            filtered_master[[
                                "Item", "Description", "Requirements", "Is Master Assembly", "Excel Row"
                            ]]
                        )
                        master_builder.configure_default_column(
                            editable=False, sortable=False, filter=False, resizable=True,
                            wrapText=True,
                        )
                        master_builder.configure_selection(
                            selection_mode="single",
                            use_checkbox=True,
                            pre_selected_rows=preselected_master_rows,
                            suppressRowClickSelection=False,
                        )
                        master_builder.configure_column(
                            "Item", pinned="left", width=170,
                            tooltipValueGetter=master_tooltip,
                        )
                        master_builder.configure_column(
                            "Description", width=470, tooltipValueGetter=master_tooltip,
                        )
                        master_builder.configure_column("Requirements", width=420)
                        master_builder.configure_column("Is Master Assembly", width=145)
                        master_builder.configure_column("Excel Row", width=100)
                        master_builder.configure_grid_options(
                            rowHeight=42,
                            headerHeight=38,
                            tooltipShowDelay=2000,
                            tooltipHideDelay=12000,
                        )
                        master_grid_response = AgGrid(
                            filtered_master[[
                                "Item", "Description", "Requirements", "Is Master Assembly", "Excel Row"
                            ]],
                            gridOptions=master_builder.build(),
                            data_return_mode=DataReturnMode.CUSTOM,
                            custom_jscode_for_grid_return=master_return,
                            update_on=[("selectionChanged", 250)],
                            allow_unsafe_jscode=True,
                            fit_columns_on_grid_load=False,
                            height=275,
                            theme="light",
                            key="workspace_master_results_grid",
                            server_sync_strategy="server_wins",
                            custom_css={
                                ".ag-header": {"background-color": "#EAF1FF", "color": "#071D49", "font-weight": "700"},
                                ".ag-root-wrapper": {"border": "1px solid #B8C8E2", "border-radius": "10px"},
                            },
                        )
                        selected_excel_rows = master_grid_response.get("selectedExcelRows", None)
                        if selected_excel_rows:
                            st.session_state["workspace_master_selected_excel_row"] = str(
                                selected_excel_rows[0]
                            )
                        selected_excel_row = str(
                            st.session_state.get("workspace_master_selected_excel_row", "")
                        )
                        selected_matches = filtered_master[
                            filtered_master["Excel Row"].astype(str).eq(selected_excel_row)
                        ]
                        if selected_matches.empty:
                            st.info("Select one exact Master row from the results.")
                        else:
                            selected_master = selected_matches.iloc[0]
                            control_cols = st.columns([1.2, 1, 1, 1])
                            insert_mode_label = control_cols[0].selectbox(
                                "Place",
                                ["After selected", "Before selected", "As child", "Top", "Bottom"],
                                key="workspace_master_insert_mode",
                            )
                            is_assembly = bool(selected_master["Is Master Assembly"])
                            if is_assembly:
                                complete_group = control_cols[1].checkbox(
                                    "Insert stored subtree",
                                    value=True,
                                    key="workspace_complete_group",
                                )
                            else:
                                complete_group = False
                                control_cols[1].caption("Component row")
                            anchor_required = insert_mode_label not in {"Top", "Bottom"}
                            anchor_ok = len(valid_selected_ids) == 1 or not anchor_required
                            if anchor_required and not anchor_ok:
                                control_cols[2].caption("Select exactly one schedule row as the anchor.")
                            if control_cols[2].button(
                                "Insert",
                                type="primary",
                                disabled=not anchor_ok,
                                use_container_width=True,
                            ):
                                excel_row = int(selected_master["Excel Row"])
                                if complete_group:
                                    rows = schedule_rows_for_master_subtree_excel_row(
                                        master, excel_row, qty=1,
                                        group_id=next_group_id_from_df(schedule_df), order="",
                                    )
                                else:
                                    rows = [schedule_row_for_master_excel_row(
                                        master, excel_row,
                                        group_id=next_group_id_from_df(schedule_df), order="",
                                    )]
                                mode_map = {
                                    "After selected": "after", "Before selected": "before",
                                    "As child": "child", "Top": "top", "Bottom": "bottom",
                                }
                                anchor_id = valid_selected_ids[0] if len(valid_selected_ids) == 1 else ""
                                commit_schedule(insert_schedule_rows(
                                    schedule_df,
                                    pd.DataFrame(rows, columns=SCHEDULE_COLUMNS),
                                    mode=mode_map[insert_mode_label],
                                    target_row_id=anchor_id,
                                ))
                                preserve_grid_selection(valid_selected_ids)
                                st.rerun()
                            if control_cols[3].button(
                                "Replace selected",
                                disabled=len(valid_selected_ids) != 1,
                                use_container_width=True,
                            ):
                                commit_schedule(replace_row_from_master(
                                    schedule_df, valid_selected_ids[0], master, excel_row=int(selected_master["Excel Row"])
                                ))
                                preserve_grid_selection(valid_selected_ids)
                                st.rerun()

            if quickbooks_panel_open:
                st.markdown("##### Add from QuickBooks")
                st.caption(
                    "QuickBooks supplies the item number and description only. Hover for two seconds to see "
                    "the full description; engineering requirements remain visibly marked as required."
                )
                if quickbooks is None or quickbooks_catalog_df is None or quickbooks_catalog_df.empty:
                    st.warning("The bundled QuickBooks item catalog is unavailable.")
                else:
                    quickbooks_search = st.text_input(
                        "Find QuickBooks item",
                        placeholder="Item number or description",
                        key="workspace_quickbooks_search",
                    )
                    filtered_quickbooks = quickbooks_catalog_df.copy()
                    if quickbooks_search.strip():
                        quickbooks_mask = (
                            filtered_quickbooks["Item"].astype(str).str.contains(
                                quickbooks_search, case=False, na=False, regex=False
                            )
                            | filtered_quickbooks["Description"].astype(str).str.contains(
                                quickbooks_search, case=False, na=False, regex=False
                            )
                        )
                        filtered_quickbooks = filtered_quickbooks[quickbooks_mask]
                    filtered_quickbooks = filtered_quickbooks.head(100).reset_index(drop=True)
                    if filtered_quickbooks.empty:
                        st.info("No QuickBooks items match that search.")
                    else:
                        quickbooks_tooltip = JsCode(
                            r"""
                            function(params) {
                                const row = params.data || {};
                                return String(row['Description'] || '') + '\n\n' +
                                    'Engineering requirements: not defined in QuickBooks';
                            }
                            """
                        )
                        quickbooks_return = JsCode(
                            """
                            function({eventData}) {
                                const api = eventData.api;
                                return {
                                    selectedQuickBooksRows: api.getSelectedRows().map(
                                        row => String(row['QuickBooks Row'] || '')
                                    ).filter(Boolean)
                                };
                            }
                            """
                        )
                        remembered_quickbooks_row = str(
                            st.session_state.get("workspace_quickbooks_selected_excel_row", "")
                        )
                        preselected_quickbooks_rows = [
                            position for position, value in enumerate(
                                filtered_quickbooks["QuickBooks Row"].astype(str).tolist()
                            ) if value == remembered_quickbooks_row
                        ]
                        quickbooks_grid_columns = [
                            "Item", "Description", "Primary Category", "Schedule Status",
                            "Active Status", "QuickBooks Category", "Requirements", "QuickBooks Row",
                        ]
                        quickbooks_builder = GridOptionsBuilder.from_dataframe(
                            filtered_quickbooks[quickbooks_grid_columns]
                        )
                        quickbooks_builder.configure_default_column(
                            editable=False, sortable=False, filter=False, resizable=True, wrapText=True,
                        )
                        quickbooks_builder.configure_selection(
                            selection_mode="single",
                            use_checkbox=True,
                            pre_selected_rows=preselected_quickbooks_rows,
                            suppressRowClickSelection=False,
                        )
                        quickbooks_builder.configure_column(
                            "Item", pinned="left", width=175,
                            tooltipValueGetter=quickbooks_tooltip,
                        )
                        quickbooks_builder.configure_column(
                            "Description", width=500, tooltipValueGetter=quickbooks_tooltip,
                        )
                        quickbooks_builder.configure_column("Primary Category", width=165)
                        quickbooks_builder.configure_column("Schedule Status", width=130)
                        quickbooks_builder.configure_column("Active Status", width=130)
                        quickbooks_builder.configure_column("QuickBooks Category", width=165)
                        quickbooks_builder.configure_column("Requirements", width=290)
                        quickbooks_builder.configure_column("QuickBooks Row", width=120)
                        quickbooks_builder.configure_grid_options(
                            rowHeight=42,
                            headerHeight=38,
                            tooltipShowDelay=2000,
                            tooltipHideDelay=12000,
                        )
                        quickbooks_grid_response = AgGrid(
                            filtered_quickbooks[quickbooks_grid_columns],
                            gridOptions=quickbooks_builder.build(),
                            data_return_mode=DataReturnMode.CUSTOM,
                            custom_jscode_for_grid_return=quickbooks_return,
                            update_on=[("selectionChanged", 250)],
                            allow_unsafe_jscode=True,
                            fit_columns_on_grid_load=False,
                            height=275,
                            theme="light",
                            key="workspace_quickbooks_results_grid",
                            server_sync_strategy="server_wins",
                            custom_css={
                                ".ag-header": {"background-color": "#E8F5EE", "color": "#143D2B", "font-weight": "700"},
                                ".ag-root-wrapper": {"border": "1px solid #A9CFB9", "border-radius": "10px"},
                            },
                        )
                        selected_quickbooks_rows = quickbooks_grid_response.get(
                            "selectedQuickBooksRows", None
                        )
                        if selected_quickbooks_rows:
                            st.session_state["workspace_quickbooks_selected_excel_row"] = str(
                                selected_quickbooks_rows[0]
                            )
                        selected_quickbooks_row = str(
                            st.session_state.get("workspace_quickbooks_selected_excel_row", "")
                        )
                        selected_quickbooks_matches = filtered_quickbooks[
                            filtered_quickbooks["QuickBooks Row"].astype(str).eq(
                                selected_quickbooks_row
                            )
                        ]
                        if selected_quickbooks_matches.empty:
                            st.info("Select one exact QuickBooks item from the results.")
                        else:
                            selected_quickbooks = selected_quickbooks_matches.iloc[0]
                            quickbooks_controls = st.columns([1.3, 2, 1])
                            quickbooks_insert_mode = quickbooks_controls[0].selectbox(
                                "Place",
                                ["After selected", "Before selected", "As child", "Top", "Bottom"],
                                key="workspace_quickbooks_insert_mode",
                            )
                            quickbooks_controls[1].caption(
                                "Item # and description come from QuickBooks; requirements remain blank for engineering."
                            )
                            quickbooks_anchor_required = quickbooks_insert_mode not in {"Top", "Bottom"}
                            quickbooks_anchor_ok = (
                                len(valid_selected_ids) == 1 or not quickbooks_anchor_required
                            )
                            if quickbooks_anchor_required and not quickbooks_anchor_ok:
                                quickbooks_controls[1].caption(
                                    "Select exactly one schedule row as the insertion anchor."
                                )
                            if quickbooks_controls[2].button(
                                "Insert",
                                type="primary",
                                disabled=not quickbooks_anchor_ok,
                                use_container_width=True,
                                key="workspace_quickbooks_insert",
                            ):
                                quickbooks_row = schedule_row_for_quickbooks_excel_row(
                                    quickbooks,
                                    int(selected_quickbooks["QuickBooks Row"]),
                                    group_id=next_group_id_from_df(schedule_df),
                                    order="",
                                )
                                mode_map = {
                                    "After selected": "after", "Before selected": "before",
                                    "As child": "child", "Top": "top", "Bottom": "bottom",
                                }
                                anchor_id = valid_selected_ids[0] if len(valid_selected_ids) == 1 else ""
                                commit_schedule(insert_schedule_rows(
                                    schedule_df,
                                    pd.DataFrame([quickbooks_row], columns=SCHEDULE_COLUMNS),
                                    mode=mode_map[quickbooks_insert_mode],
                                    target_row_id=anchor_id,
                                ))
                                preserve_grid_selection(valid_selected_ids)
                                st.rerun()

        st.caption(
            f"Showing {len(grid_df)} of {len(schedule_df)} schedule rows. "
            "Cell changes save automatically to the draft; structural actions are immediately undoable."
        )

with preview_tab:
    st.subheader("Electrical Schedule Preview")
    preview_cols = [c for c in ["Order", "Nested Code", "Item", "#", "Description", "HP", "Phase", "Volts", "Amps", "C.B.", "Air Port", "Cold Water", "Hot Water", "Reclaim Water", "Gas BTUH"] if c in schedule_df.columns]
    preview_df = schedule_df[preview_cols].copy()
    for column in preview_df.columns:
        preview_df[column] = preview_df[column].fillna("").astype(str)
    st.dataframe(preview_df, use_container_width=True, hide_index=True, height=720)

with review_tab:
    st.subheader("Review Items")
    if review_df is None or review_df.empty:
        st.info("No review items found.")
    else:
        review_edit = review_df.copy()
        if "Decision" not in review_edit.columns:
            review_edit["Decision"] = "Needs Review"
        if "Engineer Note" not in review_edit.columns:
            review_edit["Engineer Note"] = ""
        review_snapshot = review_item_count_snapshot(schedule_df, review_edit)
        review_snapshot["_Review Index"] = [str(index) for index in review_edit.index]
        review_notice = st.session_state.pop("review_insert_notice", "")
        if review_notice:
            st.success(review_notice)

        st.caption(
            "Select one or more unresolved quote rows. The count check compares exact item codes and exact "
            "descriptions against the current draft before anything is inserted."
        )
        review_toolbar_placeholder = st.empty()

        review_tooltip = JsCode(
            r"""
            function(params) {
                const row = params.data || {};
                return String(row['Description'] || '') + '\n\nIssue: ' +
                    String(row['Issue Type'] || '') + '\n' + String(row['Details'] || '') +
                    '\n\nDraft check: ' + String(row['Draft Count Check'] || '');
            }
            """
        )
        review_return = JsCode(
            """
            function({streamlitRerunEventTriggerName, eventData}) {
                const api = eventData.api;
                const selectedReviewIndices = api.getSelectedRows().map(
                    row => String(row['_Review Index'] || '')
                ).filter(Boolean);
                const editedRows = [];
                if (streamlitRerunEventTriggerName === 'cellValueChanged') {
                    api.forEachNode(node => editedRows.push(node.data));
                }
                return {
                    selectedReviewIndices: selectedReviewIndices,
                    editedRows: editedRows
                };
            }
            """
        )
        review_columns = [
            "Order", "Item", "Description", "Qty", "Issue Type", "Details",
            "Exact Item Rows in Draft", "Exact Description Rows in Draft", "Quote Demand",
            "Suggested Rows To Add", "Draft Count Check", "Decision", "Engineer Note", "_Review Index",
        ]
        review_columns = [column for column in review_columns if column in review_snapshot.columns]
        review_grid_df = review_snapshot[review_columns].copy().reset_index(drop=True)
        for column in review_grid_df.columns:
            review_grid_df[column] = review_grid_df[column].fillna("")
        remembered_review_indices = set(
            st.session_state.get("review_grid_selected_indices", [])
        )
        preselected_review_rows = [
            position for position, value in enumerate(review_grid_df["_Review Index"].astype(str))
            if value in remembered_review_indices
        ]
        review_builder = GridOptionsBuilder.from_dataframe(review_grid_df)
        review_builder.configure_default_column(
            editable=False, sortable=False, filter=False, resizable=True, wrapText=True,
        )
        review_builder.configure_selection(
            selection_mode="multiple",
            use_checkbox=True,
            header_checkbox=True,
            header_checkbox_filtered_only=True,
            pre_selected_rows=preselected_review_rows,
            rowMultiSelectWithClick=True,
        )
        review_builder.configure_column("Order", pinned="left", width=80)
        review_builder.configure_column("Item", pinned="left", width=165, tooltipValueGetter=review_tooltip)
        review_builder.configure_column("Description", width=390, tooltipValueGetter=review_tooltip)
        review_builder.configure_column("Issue Type", width=210)
        review_builder.configure_column("Details", width=330)
        review_builder.configure_column("Draft Count Check", width=245)
        for count_column in [
            "Exact Item Rows in Draft", "Exact Description Rows in Draft",
            "Quote Demand", "Suggested Rows To Add",
        ]:
            if count_column in review_grid_df.columns:
                review_builder.configure_column(count_column, width=145)
        review_builder.configure_column(
            "Decision",
            editable=True,
            width=175,
            cellEditor="agSelectCellEditor",
            cellEditorParams={
                "values": [
                    "Needs Review", "Known Non-Schedule", "Map Later",
                    "Custom Item Required", "Added to Draft", "Resolved",
                ]
            },
        )
        review_builder.configure_column("Engineer Note", editable=True, width=250)
        review_builder.configure_column("_Review Index", hide=True)
        review_builder.configure_grid_options(
            rowHeight=44,
            headerHeight=48,
            tooltipShowDelay=2000,
            tooltipHideDelay=12000,
        )
        review_grid_response = AgGrid(
            review_grid_df,
            gridOptions=review_builder.build(),
            data_return_mode=DataReturnMode.CUSTOM,
            custom_jscode_for_grid_return=review_return,
            update_on=[("cellValueChanged", 250), ("selectionChanged", 400)],
            allow_unsafe_jscode=True,
            fit_columns_on_grid_load=False,
            height=620,
            theme="light",
            key=f"review_items_grid_{int(st.session_state.get('review_grid_revision', 0))}",
            custom_css={
                ".ag-header": {"background-color": "#EAF1FF", "color": "#071D49", "font-weight": "700"},
                ".ag-root-wrapper": {"border": "1px solid #B8C8E2", "border-radius": "12px"},
                ".ag-row-selected": {"outline": "2px solid #C5222E !important", "outline-offset": "-2px"},
            },
        )

        edited_review_rows = review_grid_response.get("editedRows", [])
        edited_review_frame = (
            edited_review_rows.copy()
            if isinstance(edited_review_rows, pd.DataFrame)
            else pd.DataFrame(edited_review_rows)
        )
        if not edited_review_frame.empty:
            review_next = review_edit.copy()
            review_index_lookup = {str(index): index for index in review_next.index}
            for _, edited_row in edited_review_frame.iterrows():
                target_index = review_index_lookup.get(str(edited_row.get("_Review Index", "")))
                if target_index is None:
                    continue
                for column in ["Decision", "Engineer Note"]:
                    if column in edited_row:
                        review_next.at[target_index, column] = edited_row.get(column, "")
            st.session_state["review_df"] = review_next
            review_edit = review_next
            review_df = review_next

        response_review_indices = review_grid_response.get("selectedReviewIndices", None)
        if response_review_indices is not None:
            st.session_state["review_grid_selected_indices"] = [
                str(value) for value in response_review_indices if str(value).strip()
            ]
        selected_review_index_strings = list(
            st.session_state.get("review_grid_selected_indices", [])
        )
        review_index_lookup = {str(index): index for index in review_edit.index}
        selected_review_indices = [
            review_index_lookup[value]
            for value in selected_review_index_strings
            if value in review_index_lookup
        ]
        selected_review_df = review_edit.loc[selected_review_indices].copy() \
            if selected_review_indices else pd.DataFrame(columns=review_edit.columns)
        pending_review_rows, pending_report = schedule_rows_for_review_items(
            schedule_df, selected_review_df, review_edit
        )

        with review_toolbar_placeholder.container(border=True):
            toolbar_cols = st.columns([1.15, 1.3, 2.4, 1.15])
            toolbar_cols[0].markdown(
                f"**{len(selected_review_indices)} selected**  \n"
                f"{len(pending_review_rows)} missing schedule row"
                f"{'s' if len(pending_review_rows) != 1 else ''}"
            )
            review_place = toolbar_cols[1].selectbox(
                "Place in draft",
                ["Bottom", "Top", "After row", "Before row", "As child of row"],
                key="review_insert_mode",
            )
            schedule_target_options = row_options(schedule_df)
            schedule_target_label = ""
            target_required = review_place not in {"Top", "Bottom"}
            if target_required:
                schedule_target_label = toolbar_cols[2].selectbox(
                    "Draft destination",
                    list(schedule_target_options.keys()),
                    key="review_insert_target",
                ) if schedule_target_options else ""
            else:
                toolbar_cols[2].caption(
                    "Rows added from Review remain red in the web draft and final Excel schedule."
                )
            can_add_review = (
                bool(selected_review_indices)
                and not pending_review_rows.empty
                and (not target_required or bool(schedule_target_label))
            )
            if toolbar_cols[3].button(
                "Add to Draft",
                type="primary",
                disabled=not can_add_review,
                use_container_width=True,
            ):
                review_mode_map = {
                    "Bottom": "bottom", "Top": "top", "After row": "after",
                    "Before row": "before", "As child of row": "child",
                }
                target_id = schedule_target_options.get(schedule_target_label, "")
                commit_schedule(insert_schedule_rows(
                    schedule_df,
                    pending_review_rows,
                    mode=review_mode_map[review_place],
                    target_row_id=target_id,
                ))
                updated_review = review_edit.copy()
                for index in selected_review_indices:
                    updated_review.at[index, "Decision"] = "Added to Draft"
                st.session_state["review_df"] = updated_review
                st.session_state["review_grid_selected_indices"] = []
                st.session_state["review_grid_revision"] = int(
                    st.session_state.get("review_grid_revision", 0)
                ) + 1
                added_summary = ", ".join(
                    f"{entry['Item']}: {entry['Before']} → {entry['After']}"
                    for entry in pending_report if entry.get("Added")
                )
                st.session_state["review_insert_notice"] = (
                    f"Added {len(pending_review_rows)} missing row(s) to the draft. Count check: {added_summary}."
                )
                clear_grid_selection()
                st.rerun()

            if selected_review_indices and pending_review_rows.empty:
                st.info(
                    "The selected exact item/description quantities are already represented in the draft, "
                    "so no duplicate rows will be added."
                )

with quote_tab:
    st.subheader("Quote Extract")
    st.dataframe(quote_df, use_container_width=True, hide_index=True, height=680)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<div class="avw-section-title"><span class="avw-step">3</span>Export Excel</div>', unsafe_allow_html=True)
st.markdown('<p class="avw-section-caption">The exported workbook keeps the generated schedule editable with formulas and review sheets.</p>', unsafe_allow_html=True)

output_xlsx = cached_output_workbook_bytes(
    meta,
    quote_df if quote_df is not None else pd.DataFrame(),
    schedule_df if schedule_df is not None else pd.DataFrame(columns=SCHEDULE_COLUMNS),
    review_df if review_df is not None else pd.DataFrame(),
    rules=rules,
)
customer_slug = slugify(meta.get("customer") or "customer")
invoice = meta.get("invoice_number") or datetime.now().strftime("%Y%m%d")
filename = f"AVW_Electrical_Schedule_{customer_slug}_{invoice}_edited.xlsx"
st.download_button(
    "Download Edited Excel Schedule",
    data=output_xlsx,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.markdown(
    '<div class="avw-footer-note">Made with ❤️ by Aroon Kumar</div>',
    unsafe_allow_html=True,
)
