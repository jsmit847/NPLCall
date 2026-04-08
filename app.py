from __future__ import annotations

import hashlib
import html
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from workbook_utils import (
    EDITABLE_HEADERS,
    GRID_CONTEXT_HEADERS,
    LONG_TEXT_HEADERS,
    PREVIOUS_WEEKLY_COMMENT_HEADER,
    READONLY_CONTEXT_HEADERS,
    SECTION_MAP,
    THIS_WEEK_COMMENT_HEADER,
    apply_detail_edit,
    build_comment_text,
    build_filtered_view,
    build_picklist_options,
    build_updated_workbook,
    editable_text,
    format_deal_label,
    make_editor_df,
    materialize_export_editor_df,
    metric_currency,
    metric_date,
    metric_int,
    metric_pct,
    parse_deal_workbook,
    preview_text,
    workbook_change_set,
)

APP_TZ = ZoneInfo("America/New_York")

KEY_FIELD_SPECS = [
    ("Location", editable_text),
    ("Bridge / Term", editable_text),
    ("Type", editable_text),
    ("Segment", editable_text),
    ("Financing", editable_text),
    ("Loan Commitment", metric_currency),
    ("Remaining Commitment", metric_currency),
    ("Number Of Units", metric_int),
    ("Salesforce Valuation Date", metric_date),
    ("Maturity Date", metric_date),
    ("Next Payment Date", metric_date),
    ("12/31 NPL", editable_text),
    ("3/31 NPL", editable_text),
    ("6/30 NPL", editable_text),
    ("3/27 DQ Status", editable_text),
    ("MBA Loan List Convention", editable_text),
]

OVERVIEW_COLUMNS = [
    "Deal Number",
    "Deal Name",
    "Asset Manager",
    "Bridge / Term",
    "Segment",
    "3/27 UPB",
    "Salesforce Implied As-Is LTV",
    "Salesforce ARV LTV",
    "6/30 NPL",
    "3/27 DQ Status",
    "Expected Final Resolution",
    "Resolution Timing",
    THIS_WEEK_COMMENT_HEADER,
    "Comments",
]

BULK_COLUMN_ORDER = [
    "_excel_row",
    "Deal Number",
    "Deal Name",
    "Asset Manager",
    "6/30 NPL",
    "3/27 DQ Status",
    "Resolution Likelihood",
    "Expected Final Resolution",
    "Resolution Timing",
    "Liquidity Event Timing",
    "Mod Needed Prior to Final Resolution (Y/N)",
    "FCL Needed Prior to Final Resolution (Y/N)",
    "Pref Deal (Y/N)",
    "CV Inspected in Last 6 months? (Y/N)",
    "New Valuation in 1Q26 (Y/N)",
    "Valuation Type",
    "Is Appraisal Finalized? (Y/N)",
    "FCL Date",
    "Pref Amount ($)",
    "Liquidity Event ($)",
    "As Is Appraised Value",
    "As Stabilized Value",
    "Report Date",
    "Date of Valuation",
    "Third Party Offer Amount",
    "Resolution for 6/30 Performance",
    "Addt'l NPL Comment",
    "Third Party Offer Note",
    PREVIOUS_WEEKLY_COMMENT_HEADER,
    THIS_WEEK_COMMENT_HEADER,
    "3/27 Salesforce AM Comment",
]


@st.cache_data(show_spinner=False)
def cached_parse_workbook(file_bytes: bytes):
    return parse_deal_workbook(file_bytes)



def current_comment_label() -> str:
    return datetime.now(APP_TZ).strftime("%m/%d/%y")



def inject_app_css(presentation_style: bool) -> None:
    spacing_css = """
    .block-container {
        padding-top: 0.85rem;
        padding-bottom: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: 100%;
    }
    """
    if presentation_style:
        spacing_css = """
        .block-container {
            padding-top: 0.45rem;
            padding-bottom: 0.65rem;
            padding-left: 0.7rem;
            padding-right: 0.7rem;
            max-width: 100%;
        }
        section[data-testid="stSidebar"] {
            width: 300px !important;
        }
        """

    st.markdown(
        f"""
        <style>
        {spacing_css}
        div[data-testid="stMetric"] {{
            background: white;
            border: 1px solid #e5e7eb;
            border-radius: 0.6rem;
            padding: 0.2rem 0.55rem 0.35rem 0.55rem;
        }}
        div[data-testid="stMetricValue"] {{
            font-size: 1.05rem;
        }}
        .app-title {{
            font-size: 1rem;
            font-weight: 700;
            margin-bottom: 0.2rem;
        }}
        .app-subtitle {{
            color: #4b5563;
            font-size: 0.86rem;
            margin-bottom: 0.7rem;
        }}
        .deal-banner {{
            border: 1px solid #d1d5db;
            background: #f8fafc;
            border-radius: 0.75rem;
            padding: 0.55rem 0.7rem;
            margin: 0.1rem 0 0.55rem 0;
        }}
        .deal-banner-title {{
            font-size: 1rem;
            font-weight: 700;
            line-height: 1.25;
            margin-bottom: 0.2rem;
        }}
        .deal-banner-subtitle {{
            font-size: 0.82rem;
            color: #4b5563;
        }}
        .badge {{
            display: inline-block;
            padding: 0.12rem 0.45rem;
            border-radius: 999px;
            background: #e5e7eb;
            font-size: 0.73rem;
            font-weight: 600;
            margin-right: 0.28rem;
            margin-top: 0.15rem;
        }}
        .kv-card {{
            border: 1px solid #e5e7eb;
            border-radius: 0.6rem;
            padding: 0.45rem 0.55rem;
            margin-bottom: 0.4rem;
            min-height: 3.5rem;
            background: white;
        }}
        .kv-label {{
            font-size: 0.72rem;
            color: #6b7280;
            margin-bottom: 0.2rem;
            text-transform: uppercase;
            letter-spacing: 0.02em;
        }}
        .kv-value {{
            font-size: 0.9rem;
            font-weight: 600;
            line-height: 1.2;
            word-break: break-word;
        }}
        .comment-card {{
            border: 1px solid #e5e7eb;
            border-radius: 0.7rem;
            padding: 0.55rem 0.65rem;
            min-height: 10.5rem;
            background: white;
        }}
        .comment-card-title {{
            font-size: 0.76rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.02em;
            color: #6b7280;
            margin-bottom: 0.35rem;
        }}
        .comment-card-body {{
            font-size: 0.87rem;
            line-height: 1.35;
            white-space: pre-wrap;
            word-break: break-word;
        }}
        .queue-caption {{
            color: #4b5563;
            font-size: 0.8rem;
            margin: 0.2rem 0 0.45rem 0;
        }}
        .present-note {{
            color: #374151;
            font-size: 0.82rem;
            margin-top: -0.15rem;
            margin-bottom: 0.45rem;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )



def init_state(file_hash: str, parsed: dict) -> None:
    st.session_state["file_hash"] = file_hash
    st.session_state["editor_df"] = make_editor_df(parsed["deals_df"])
    st.session_state["selected_row"] = int(parsed["deals_df"].iloc[0]["_excel_row"])
    st.session_state.setdefault("comment_label", current_comment_label())



def display_value(header: str, value) -> str:
    if header in {"3/27 UPB", "Salesforce As-Is Valuation", "Loan Commitment", "Salesforce ARV", "Remaining Commitment"}:
        return metric_currency(value)
    if header in {"Salesforce Implied As-Is LTV", "Salesforce ARV LTV", "12/31 Px"}:
        return metric_pct(value)
    if "Date" in header:
        return metric_date(value)
    if header == "Number Of Units":
        return metric_int(value)
    text = editable_text(value).strip()
    return text or "-"



def render_comment_card(title: str, text: str, height_lines: int = 9) -> None:
    body = html.escape(text.strip() if text and text.strip() else "-")
    min_height = 9.5 + max(0, height_lines - 9) * 0.5
    st.markdown(
        f"""
        <div class="comment-card" style="min-height:{min_height}rem;">
            <div class="comment-card-title">{html.escape(title)}</div>
            <div class="comment-card-body">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def render_banner(selected: pd.Series, queue_text: str | None = None) -> None:
    badges = [
        editable_text(selected.get("Bridge / Term")) or "-",
        editable_text(selected.get("Segment")) or "-",
        f"6/30 NPL: {editable_text(selected.get('6/30 NPL')) or '-'}",
        f"DQ: {editable_text(selected.get('3/27 DQ Status')) or '-'}",
    ]
    badge_html = "".join(f'<span class="badge">{html.escape(b)}</span>' for b in badges)
    subtitle = (
        f"Deal {editable_text(selected.get('Deal Number'))} | "
        f"Asset Manager: {editable_text(selected.get('Asset Manager')) or '-'}"
    )
    if queue_text:
        subtitle = f"{subtitle} | {queue_text}"
    st.markdown(
        f"""
        <div class="deal-banner">
            <div class="deal-banner-title">{html.escape(editable_text(selected.get('Deal Name')) or '-')}</div>
            <div class="deal-banner-subtitle">{html.escape(subtitle)}</div>
            <div>{badge_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def render_metric_rows(selected: pd.Series) -> None:
    first_row = st.columns(6)
    first_row[0].metric("UPB", metric_currency(selected.get("3/27 UPB")))
    first_row[1].metric("As-Is Value", metric_currency(selected.get("Salesforce As-Is Valuation")))
    first_row[2].metric("As-Is LTV", metric_pct(selected.get("Salesforce Implied As-Is LTV")))
    first_row[3].metric("ARV", metric_currency(selected.get("Salesforce ARV")))
    first_row[4].metric("ARV LTV", metric_pct(selected.get("Salesforce ARV LTV")))
    first_row[5].metric("Remaining Commit.", metric_currency(selected.get("Remaining Commitment")))

    second_row = st.columns(6)
    second_row[0].metric("Commitment", metric_currency(selected.get("Loan Commitment")))
    second_row[1].metric("Units", metric_int(selected.get("Number Of Units")))
    second_row[2].metric("Maturity", metric_date(selected.get("Maturity Date")))
    second_row[3].metric("Next Pmt", metric_date(selected.get("Next Payment Date")))
    second_row[4].metric("12/31 NPL", editable_text(selected.get("12/31 NPL")) or "-")
    second_row[5].metric("3/31 NPL", editable_text(selected.get("3/31 NPL")) or "-")



def render_key_value_grid(selected: pd.Series, presentation_mode: bool = False) -> None:
    cols_per_row = 4 if presentation_mode else 3
    field_specs = KEY_FIELD_SPECS
    for start in range(0, len(field_specs), cols_per_row):
        cols = st.columns(cols_per_row)
        for idx, (header, formatter) in enumerate(field_specs[start : start + cols_per_row]):
            value = formatter(selected.get(header))
            with cols[idx]:
                st.markdown(
                    f"""
                    <div class="kv-card">
                        <div class="kv-label">{html.escape(header)}</div>
                        <div class="kv-value">{html.escape(value or '-')}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )



def select_options_for_value(header: str, current_value: str, picklist_options: dict[str, list[str]]) -> list[str] | None:
    base_options = picklist_options.get(header)
    if not base_options:
        return None
    current_clean = editable_text(current_value).strip()
    if current_clean and current_clean not in base_options:
        return ["", current_clean, *[value for value in base_options if value != ""]]
    return base_options



def render_input_widget(header: str, current_value: str, picklist_options: dict[str, list[str]], key_name: str) -> str:
    options = select_options_for_value(header, current_value, picklist_options)
    current_text = editable_text(current_value)
    if options:
        try:
            current_index = options.index(current_text)
        except ValueError:
            current_index = 0
        return st.selectbox(header, options, index=current_index, key=key_name)
    if header in LONG_TEXT_HEADERS:
        return st.text_area(header, current_text, height=120, key=key_name)
    return st.text_input(header, current_text, key=key_name)



def render_update_form(
    selected: pd.Series,
    picklist_options: dict[str, list[str]],
    comment_date_label: str,
    key_prefix: str,
    presentation_mode: bool = False,
):
    excel_row = int(selected["_excel_row"])
    current_previous_comment = editable_text(selected.get(PREVIOUS_WEEKLY_COMMENT_HEADER))
    current_week_comment = editable_text(selected.get(THIS_WEEK_COMMENT_HEADER))
    salesforce_comment = editable_text(selected.get("3/27 Salesforce AM Comment"))
    preview_comment = build_comment_text(current_previous_comment, current_week_comment, comment_date_label)

    st.markdown("#### AM update workspace")
    st.markdown(
        f'<div class="present-note">This week\'s comment will be prepended as <strong>{html.escape(comment_date_label)}</strong> when you download the workbook, so the exported file keeps the same single <strong>Comments</strong> column layout.</div>',
        unsafe_allow_html=True,
    )

    with st.form(f"{key_prefix}_form_{excel_row}"):
        c1, c2, c3 = st.columns(3)
        with c1:
            render_comment_card("Salesforce AM Comment", salesforce_comment)
        with c2:
            render_comment_card("Previous Weekly Comment", current_previous_comment)
        with c3:
            this_week_comment = st.text_area(
                THIS_WEEK_COMMENT_HEADER,
                current_week_comment,
                height=182 if presentation_mode else 170,
                key=f"{key_prefix}_this_week_comment_{excel_row}",
            )

        preview_comment = build_comment_text(current_previous_comment, this_week_comment, comment_date_label)
        render_comment_card("Export Preview for Comments Cell", preview_comment, height_lines=10)

        updated_values = {THIS_WEEK_COMMENT_HEADER: this_week_comment}
        cols_per_row = 3 if presentation_mode else 2

        for section_title, headers in SECTION_MAP.items():
            st.markdown(f"##### {section_title}")
            short_headers = [header for header in headers if header not in LONG_TEXT_HEADERS]
            long_headers = [header for header in headers if header in LONG_TEXT_HEADERS]

            for start in range(0, len(short_headers), cols_per_row):
                cols = st.columns(cols_per_row)
                for idx, header in enumerate(short_headers[start : start + cols_per_row]):
                    with cols[idx]:
                        updated_values[header] = render_input_widget(
                            header,
                            editable_text(selected.get(header)),
                            picklist_options,
                            key_name=f"{key_prefix}_{excel_row}_{header}",
                        )

            for header in long_headers:
                updated_values[header] = render_input_widget(
                    header,
                    editable_text(selected.get(header)),
                    picklist_options,
                    key_name=f"{key_prefix}_{excel_row}_{header}",
                )

        saved = st.form_submit_button("Save app edits for this deal", use_container_width=True)
        if saved:
            return updated_values

    return None



def render_property_detail(selected: pd.Series, properties_df: pd.DataFrame, expanded: bool = False) -> None:
    if properties_df.empty or "Deal Number" not in properties_df.columns:
        return
    related = properties_df[properties_df["Deal Number"] == selected["Deal Number"]].copy()
    if related.empty:
        return

    property_columns = [
        column
        for column in [
            "Servicer ID",
            "Servicer",
            "SF Yardi ID",
            "Asset ID",
            "Address",
            "City",
            "State",
            "# Units",
            "Financing",
            "First Funded Date",
            "Maturity Date",
            "3/27 UPB",
            "Salesforce As-Is Valuation",
            "Salesforce ARV",
            "Days Past Due",
        ]
        if column in related.columns
    ]

    with st.expander(f"Related property detail ({len(related)})", expanded=expanded):
        st.dataframe(related[property_columns], use_container_width=True, hide_index=True)



def render_deal_workspace(
    selected: pd.Series,
    properties_df: pd.DataFrame,
    picklist_options: dict[str, list[str]],
    comment_date_label: str,
    key_prefix: str,
    presentation_mode: bool = False,
    queue_text: str | None = None,
):
    render_banner(selected, queue_text=queue_text)
    render_metric_rows(selected)
    st.markdown("#### Deal datapoints")
    render_key_value_grid(selected, presentation_mode=presentation_mode)
    updates = render_update_form(
        selected=selected,
        picklist_options=picklist_options,
        comment_date_label=comment_date_label,
        key_prefix=key_prefix,
        presentation_mode=presentation_mode,
    )
    render_property_detail(selected, properties_df, expanded=not presentation_mode)
    return updates



def build_overview_table(filtered_df: pd.DataFrame, changed_rows: set[int]) -> pd.DataFrame:
    overview_df = filtered_df[OVERVIEW_COLUMNS].copy()
    overview_df.insert(3, "Updated This Session", overview_df.index.map(lambda idx: "Y" if int(filtered_df.loc[idx, "_excel_row"]) in changed_rows else ""))
    overview_df["3/27 UPB"] = overview_df["3/27 UPB"].apply(metric_currency)
    overview_df["Salesforce Implied As-Is LTV"] = overview_df["Salesforce Implied As-Is LTV"].apply(metric_pct)
    overview_df["Salesforce ARV LTV"] = overview_df["Salesforce ARV LTV"].apply(metric_pct)
    overview_df[THIS_WEEK_COMMENT_HEADER] = overview_df[THIS_WEEK_COMMENT_HEADER].apply(lambda x: preview_text(x, 70))
    overview_df["Comments"] = overview_df["Comments"].apply(lambda x: preview_text(x, 110))
    return overview_df.reset_index(drop=True)



def render_overview(filtered_df: pd.DataFrame, changed_rows: set[int]) -> None:
    needs_comment_count = int(filtered_df[THIS_WEEK_COMMENT_HEADER].fillna("").astype(str).str.strip().eq("").sum())
    changed_in_view = len(set(filtered_df["_excel_row"]).intersection(changed_rows))

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Deals in view", f"{len(filtered_df):,}")
    c2.metric("UPB in view", metric_currency(filtered_df["3/27 UPB"].fillna(0).sum()))
    c3.metric("Asset Managers", f"{filtered_df['Asset Manager'].astype(str).nunique():,}")
    c4.metric("6/30 NPL = Y", f"{filtered_df[filtered_df['6/30 NPL'].astype(str).str.upper().eq('Y')].shape[0]:,}")
    c5.metric("Needs this-week comment", f"{needs_comment_count:,}")
    c6.metric("Edited this session", f"{changed_in_view:,}")

    st.dataframe(build_overview_table(filtered_df, changed_rows), use_container_width=True, hide_index=True, height=560)



def build_bulk_column_config(picklist_options: dict[str, list[str]]) -> dict:
    config: dict = {
        "Deal Number": st.column_config.TextColumn("Deal Number", width="small"),
        "Deal Name": st.column_config.TextColumn("Deal Name", width="medium"),
        "Asset Manager": st.column_config.TextColumn("Asset Manager", width="small"),
        PREVIOUS_WEEKLY_COMMENT_HEADER: st.column_config.TextColumn(PREVIOUS_WEEKLY_COMMENT_HEADER, width="large"),
        THIS_WEEK_COMMENT_HEADER: st.column_config.TextColumn(THIS_WEEK_COMMENT_HEADER, width="large"),
        "3/27 Salesforce AM Comment": st.column_config.TextColumn("3/27 Salesforce AM Comment", width="large"),
        "Resolution for 6/30 Performance": st.column_config.TextColumn("Resolution for 6/30 Performance", width="large"),
        "Addt'l NPL Comment": st.column_config.TextColumn("Addt'l NPL Comment", width="large"),
        "Third Party Offer Note": st.column_config.TextColumn("Third Party Offer Note", width="medium"),
    }
    for header, options in picklist_options.items():
        if header in BULK_COLUMN_ORDER:
            config[header] = st.column_config.SelectboxColumn(header, options=options, width="small")
    return config



def next_blank_comment_row(option_rows: list[int], editor_df: pd.DataFrame, current_row: int) -> int | None:
    if not option_rows:
        return None
    comment_lookup = editor_df.set_index("_excel_row")[THIS_WEEK_COMMENT_HEADER].to_dict()
    if current_row not in option_rows:
        current_row = option_rows[0]
    start_idx = option_rows.index(current_row) + 1
    search_order = option_rows[start_idx:] + option_rows[:start_idx]
    for row_num in search_order:
        if not editable_text(comment_lookup.get(row_num, "")).strip():
            return row_num
    return None



def select_current_deal(options: pd.DataFrame, editor_df: pd.DataFrame, key_prefix: str) -> pd.Series:
    option_rows = options["_excel_row"].tolist()
    if st.session_state.get("selected_row") not in option_rows:
        st.session_state["selected_row"] = option_rows[0]

    current_idx = option_rows.index(st.session_state["selected_row"])
    next_blank_row = next_blank_comment_row(option_rows, editor_df, st.session_state["selected_row"])

    c1, c2, c3, c4 = st.columns([1, 1.2, 1.6, 6])
    with c1:
        if st.button("Prev", key=f"{key_prefix}_prev", disabled=current_idx == 0, use_container_width=True):
            st.session_state["selected_row"] = option_rows[current_idx - 1]
            st.rerun()
    with c2:
        if st.button("Next", key=f"{key_prefix}_next", disabled=current_idx == len(option_rows) - 1, use_container_width=True):
            st.session_state["selected_row"] = option_rows[current_idx + 1]
            st.rerun()
    with c3:
        if st.button(
            "Next open comment",
            key=f"{key_prefix}_next_blank",
            disabled=next_blank_row is None,
            use_container_width=True,
        ):
            st.session_state["selected_row"] = next_blank_row
            st.rerun()
    with c4:
        selected_row = st.selectbox(
            "Deal queue",
            option_rows,
            format_func=lambda row_num: format_deal_label(options[options["_excel_row"] == row_num].iloc[0]),
            index=option_rows.index(st.session_state["selected_row"]),
            key=f"{key_prefix}_deal_select",
        )
        st.session_state["selected_row"] = selected_row

    st.markdown(
        f'<div class="queue-caption">Queue position: {current_idx + 1} of {len(option_rows)} deals in the current filtered review queue.</div>',
        unsafe_allow_html=True,
    )
    return options[options["_excel_row"] == st.session_state["selected_row"]].iloc[0]



def main() -> None:
    st.set_page_config(page_title="NPL Deal Review", layout="wide", initial_sidebar_state="expanded")

    presentation_style = st.session_state.get("presentation_style", False)
    inject_app_css(presentation_style)

    st.markdown('<div class="app-title">NPL Deal Review</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subtitle">Upload the latest workbook first. The app then presents each deal, keeps the Asset Manager directly beside the deal context, supports faster AM updates with standardized picklists, and downloads an updated workbook that preserves the original Excel layout.</div>',
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader("Step 1 - Upload the current workbook", type=["xlsx"])
    if uploaded_file is None:
        st.info("Once a workbook is uploaded, the review queue, presentation view, editing controls, and workbook download will appear here.")
        return

    file_bytes = uploaded_file.getvalue()
    file_hash = hashlib.md5(file_bytes).hexdigest()
    parsed = cached_parse_workbook(file_bytes)

    if st.session_state.get("file_hash") != file_hash:
        init_state(file_hash, parsed)

    comment_date_label = st.session_state.get("comment_label", current_comment_label())
    base_deals = parsed["deals_df"]
    properties_df = parsed["properties_df"]
    picklist_options = build_picklist_options(base_deals)
    editor_df = st.session_state["editor_df"]
    display_df = build_filtered_view(base_deals, editor_df, comment_date_label)

    export_editor_df = materialize_export_editor_df(editor_df, comment_date_label)
    changes = workbook_change_set(export_editor_df, parsed["original_editor_values"])
    changed_rows = {row for row, _ in changes.keys()}

    visible_count = int((~display_df["_workbook_hidden"]).sum())
    hidden_count = int(display_df["_workbook_hidden"].sum())
    st.caption(
        f"Workbook inspection: main deal sheet = {parsed['deal_sheet']}; property detail sheet = {parsed['property_sheet'] or 'not found'}; total deals = {len(display_df)}; workbook-visible rows = {visible_count}; workbook-hidden rows = {hidden_count}."
    )

    with st.sidebar:
        st.header("Review controls")
        st.session_state["presentation_style"] = st.toggle("Presentation styling", value=st.session_state.get("presentation_style", False))
        st.session_state["comment_label"] = st.text_input("This-week comment date label", value=comment_date_label)
        show_hidden = st.checkbox("Include workbook-hidden rows", value=False)
        needs_weekly_comment_only = st.checkbox("Needs this-week comment only", value=False)
        session_edits_only = st.checkbox("Session-edited deals only", value=False)

        asset_managers = sorted([x for x in display_df["Asset Manager"].dropna().astype(str).unique().tolist() if x])
        selected_ams = st.multiselect("Asset Manager", asset_managers)

        bridge_term_values = sorted([x for x in display_df["Bridge / Term"].dropna().astype(str).unique().tolist() if x])
        selected_bt = st.multiselect("Bridge / Term", bridge_term_values)

        segment_values = sorted([x for x in display_df["Segment"].dropna().astype(str).unique().tolist() if x])
        selected_segment = st.multiselect("Segment", segment_values)

        npl_values = sorted([x for x in display_df["6/30 NPL"].dropna().astype(str).unique().tolist() if x])
        selected_npl = st.multiselect("6/30 NPL", npl_values)

        dq_values = sorted([x for x in display_df["3/27 DQ Status"].dropna().astype(str).unique().tolist() if x])
        selected_dq = st.multiselect("3/27 DQ Status", dq_values)

        search = st.text_input("Search deal number or name")

    if st.session_state.get("presentation_style") != presentation_style:
        st.rerun()

    filtered_df = display_df.copy()
    if not show_hidden:
        filtered_df = filtered_df[~filtered_df["_workbook_hidden"]]
    if needs_weekly_comment_only:
        filtered_df = filtered_df[filtered_df[THIS_WEEK_COMMENT_HEADER].fillna("").astype(str).str.strip().eq("")]
    if session_edits_only:
        filtered_df = filtered_df[filtered_df["_excel_row"].isin(changed_rows)]
    if selected_ams:
        filtered_df = filtered_df[filtered_df["Asset Manager"].astype(str).isin(selected_ams)]
    if selected_bt:
        filtered_df = filtered_df[filtered_df["Bridge / Term"].astype(str).isin(selected_bt)]
    if selected_segment:
        filtered_df = filtered_df[filtered_df["Segment"].astype(str).isin(selected_segment)]
    if selected_npl:
        filtered_df = filtered_df[filtered_df["6/30 NPL"].astype(str).isin(selected_npl)]
    if selected_dq:
        filtered_df = filtered_df[filtered_df["3/27 DQ Status"].astype(str).isin(selected_dq)]
    if search:
        search_lower = search.lower()
        filtered_df = filtered_df[
            filtered_df["Deal Number"].astype(str).str.lower().str.contains(search_lower)
            | filtered_df["Deal Name"].astype(str).str.lower().str.contains(search_lower)
        ]

    if filtered_df.empty:
        st.warning("No deals match the current filters.")
        return

    options = filtered_df.sort_values(["Asset Manager", "Deal Name", "Deal Number"]).reset_index(drop=True)

    tab_overview, tab_review, tab_bulk, tab_presentation = st.tabs(
        ["Queue overview", "Deal review", "Bulk edit", "Presentation mode"]
    )

    with tab_overview:
        render_overview(filtered_df, changed_rows)

    with tab_review:
        selected = select_current_deal(options, editor_df, key_prefix="standard")
        queue_text = f"{options[options['_excel_row'] == st.session_state['selected_row']].index[0] + 1} of {len(options)}"
        detail_updates = render_deal_workspace(
            selected=selected,
            properties_df=properties_df,
            picklist_options=picklist_options,
            comment_date_label=st.session_state.get("comment_label", comment_date_label),
            key_prefix="standard",
            presentation_mode=False,
            queue_text=queue_text,
        )
        if detail_updates:
            st.session_state["editor_df"] = apply_detail_edit(st.session_state["editor_df"], int(selected["_excel_row"]), detail_updates)
            st.rerun()

    with tab_bulk:
        st.write("Bulk-edit the manual AM update fields across the currently filtered deals. Prior weekly comment and Salesforce AM comment stay read-only for context while this-week comment stays editable.")
        bulk_subset = st.session_state["editor_df"][st.session_state["editor_df"]["_excel_row"].isin(filtered_df["_excel_row"])].copy()
        bulk_subset = bulk_subset.sort_values(["Asset Manager", "Deal Name", "Deal Number"]).reset_index(drop=True)
        bulk_subset = bulk_subset[[column for column in BULK_COLUMN_ORDER if column in bulk_subset.columns]].copy()

        disabled_columns = ["_excel_row", *GRID_CONTEXT_HEADERS, PREVIOUS_WEEKLY_COMMENT_HEADER, *READONLY_CONTEXT_HEADERS]
        edited_subset = st.data_editor(
            bulk_subset,
            use_container_width=True,
            hide_index=True,
            disabled=disabled_columns,
            num_rows="fixed",
            column_config=build_bulk_column_config(picklist_options),
            key="bulk_editor",
            height=620,
        )

        editable_bulk_headers = [header for header in BULK_COLUMN_ORDER if header in EDITABLE_HEADERS or header == THIS_WEEK_COMMENT_HEADER]
        full_editor = st.session_state["editor_df"].set_index("_excel_row")
        edited_lookup = edited_subset.set_index("_excel_row")
        for row_num in edited_lookup.index:
            for header in editable_bulk_headers:
                if header in edited_lookup.columns:
                    full_editor.loc[row_num, header] = editable_text(edited_lookup.loc[row_num, header])
        st.session_state["editor_df"] = full_editor.reset_index()

    with tab_presentation:
        selected = select_current_deal(options, editor_df, key_prefix="presentation")
        queue_text = f"{options[options['_excel_row'] == st.session_state['selected_row']].index[0] + 1} of {len(options)}"
        presentation_updates = render_deal_workspace(
            selected=selected,
            properties_df=properties_df,
            picklist_options=picklist_options,
            comment_date_label=st.session_state.get("comment_label", comment_date_label),
            key_prefix="presentation",
            presentation_mode=True,
            queue_text=queue_text,
        )
        if presentation_updates:
            st.session_state["editor_df"] = apply_detail_edit(st.session_state["editor_df"], int(selected["_excel_row"]), presentation_updates)
            st.rerun()

    export_editor_df = materialize_export_editor_df(st.session_state["editor_df"], st.session_state.get("comment_label", comment_date_label))
    changes = workbook_change_set(export_editor_df, parsed["original_editor_values"])
    changed_rows = {row for row, _ in changes.keys()}
    changed_deals = len(changed_rows)

    st.divider()
    left, middle, right = st.columns([2.3, 1.2, 1.2])
    with left:
        st.write(f"Pending edits in app view: **{len(changes)} cells across {changed_deals} deal(s)**")
        st.caption("Workbook export keeps the original workbook structure and writes the combined weekly comment back into the same Comments column.")
    with middle:
        if st.button("Reset all app edits", use_container_width=True):
            st.session_state["editor_df"] = make_editor_df(parsed["deals_df"])
            st.rerun()
    with right:
        base_name = uploaded_file.name.rsplit(".", 1)[0]
        download_name = f"{base_name} - AM Updates.xlsx"
        updated_bytes = (
            build_updated_workbook(
                file_bytes=file_bytes,
                sheet_path=parsed["deal_sheet_path"],
                changes=changes,
                cell_meta=parsed["cell_meta"],
            )
            if changes
            else file_bytes
        )
        st.download_button(
            "Download updated workbook",
            data=updated_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
