from __future__ import annotations

import hashlib
from datetime import date
from typing import Any

import pandas as pd
import streamlit as st

from workbook_utils import (
    COMMENT_TEMPLATE_OPTIONS,
    GRID_CONTEXT_HEADERS,
    LONG_TEXT_FIELDS,
    MEETING_STATUS_OPTIONS,
    SECTION_MAP,
    apply_detail_edit,
    build_display_view,
    build_picklists,
    build_updated_workbook,
    build_weekly_comment_text,
    editable_text,
    format_comment_entry_date,
    format_deal_label,
    make_editor_df,
    metric_currency,
    metric_date,
    metric_pct,
    normalize_number,
    parse_deal_workbook,
    refresh_editor_derived_fields,
    sort_deals,
    workbook_change_set,
)


BINARY_FIELDS = {
    "Mod Needed Prior to Final Resolution (Y/N)",
    "FCL Needed Prior to Final Resolution (Y/N)",
    "Pref Deal (Y/N)",
    "CV Inspected in Last 6 months? (Y/N)",
    "New Valuation in 1Q26 (Y/N)",
    "Is Appraisal Finalized? (Y/N)",
}

LIKELIHOOD_OPTIONS = ["", "High", "Medium", "Low"]
DISCUSSION_OPTIONS = [False, True]


@st.cache_data(show_spinner=False)
def cached_parse_workbook(file_bytes: bytes):
    return parse_deal_workbook(file_bytes)


@st.cache_data(show_spinner=False)
def cached_picklists(deals_df: pd.DataFrame):
    return build_picklists(deals_df)


def apply_base_css() -> None:
    st.markdown(
        """
        <style>
            .block-container {padding-top: 0.55rem; padding-bottom: 0.9rem; max-width: 100%;}
            h1, h2, h3 {margin-top: 0.2rem !important; margin-bottom: 0.25rem !important;}
            div[data-testid="stMetricValue"] {font-size: 1.1rem !important;}
            div[data-testid="stMetricLabel"] {font-size: 0.76rem !important;}
            .deal-header {
                border: 1px solid #d7dbe2;
                border-radius: 14px;
                padding: 0.8rem 1rem;
                margin-bottom: 0.55rem;
                background: linear-gradient(180deg, #fbfcfe 0%, #f6f8fb 100%);
            }
            .deal-title {
                font-size: 2.15rem;
                font-weight: 900;
                line-height: 1.03;
                margin-right: 0.65rem;
            }
            .deal-number {
                font-size: 1.2rem;
                font-weight: 850;
                color: #344054;
                margin-right: 0.55rem;
            }
            .am-badge {
                display: inline-block;
                padding: 0.34rem 0.8rem;
                border-radius: 999px;
                border: 1px solid #bfd2ff;
                background: #edf4ff;
                font-size: 1.18rem;
                font-weight: 900;
                color: #163b7a;
            }
            .deal-subtitle {
                margin-top: 0.25rem;
                color: #5c6570;
                font-size: 0.8rem;
                line-height: 1.2;
            }
            .hero-card, .panel-card, .comment-card, .mini-card {
                border: 1px solid #e5e7eb;
                border-radius: 12px;
                background: #ffffff;
            }
            .hero-card {padding: 0.8rem 0.95rem; margin-bottom: 0.5rem;}
            .panel-card {padding: 0.7rem 0.85rem; margin-bottom: 0.5rem;}
            .comment-card {padding: 0.75rem 0.9rem; margin-bottom: 0.5rem; background: #fbfcfe;}
            .mini-card {padding: 0.45rem 0.55rem; margin-bottom: 0.38rem;}
            .section-label, .mini-label {
                color: #6b7280;
                text-transform: uppercase;
                letter-spacing: 0.02em;
                font-size: 0.72rem;
                margin-bottom: 0.15rem;
            }
            .hero-value {font-size: 1rem; line-height: 1.35; white-space: pre-wrap;}
            .comment-value {font-size: 0.95rem; line-height: 1.35; white-space: pre-wrap; font-weight: 600;}
            .mini-value {font-size: 0.92rem; line-height: 1.25; font-weight: 650;}
            .status-chip {
                display: inline-block;
                padding: 0.18rem 0.42rem;
                border-radius: 999px;
                border: 1px solid #d1d5db;
                margin-right: 0.25rem;
                margin-bottom: 0.2rem;
                font-size: 0.75rem;
                background: #f9fafb;
                font-weight: 650;
            }
            .banner {
                border-radius: 12px;
                padding: 0.65rem 0.8rem;
                margin-bottom: 0.55rem;
                border: 1px solid #dbeafe;
                background: #f8fbff;
                font-size: 0.92rem;
            }
            .toolbar {margin-bottom: 0.45rem;}
            header, footer {visibility: hidden;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def safe_text(value: Any) -> str:
    text = editable_text(value)
    return text if text else "-"


def field_options(field: str, picklists: dict[str, list[str]], current_value: str) -> list[str]:
    options = [""] + picklists.get(field, [])
    current = (current_value or "").strip()
    if current and current not in options:
        options.append(current)
    return options


def selected_index(options: list[str], current_value: str) -> int:
    current = (current_value or "").strip()
    if current in options:
        return options.index(current)
    return 0


def numeric_sum(series: pd.Series) -> float:
    total = 0.0
    for value in series:
        number = normalize_number(value)
        if number is not None:
            total += number
    return total


def init_state(file_hash: str, parsed: dict[str, Any]) -> None:
    st.session_state["file_hash"] = file_hash
    st.session_state["editor_df"] = make_editor_df(parsed["deals_df"])
    st.session_state["selected_row"] = None
    st.session_state.setdefault("comment_date", date.today())
    st.session_state.setdefault("reviewed_rows", set())
    st.session_state.setdefault("commented_rows", set())


def stash_uploaded_workbook(uploaded_file: Any) -> None:
    if uploaded_file is None:
        return
    st.session_state["workbook_bytes"] = uploaded_file.getvalue()
    st.session_state["workbook_name"] = uploaded_file.name


def clear_uploaded_workbook() -> None:
    for key in [
        "workbook_bytes",
        "workbook_name",
        "file_hash",
        "editor_df",
        "selected_row",
        "reviewed_rows",
        "commented_rows",
    ]:
        st.session_state.pop(key, None)


def get_active_workbook() -> tuple[bytes | None, str | None]:
    with st.sidebar:
        st.subheader("Workbook")
        sidebar_upload = st.file_uploader("Upload / replace workbook", type=["xlsx"], key="sidebar_workbook_upload")
        if sidebar_upload is not None:
            stash_uploaded_workbook(sidebar_upload)
        if st.session_state.get("workbook_name"):
            st.caption(f"Current workbook: {st.session_state['workbook_name']}")
            if st.button("Clear workbook", use_container_width=True):
                clear_uploaded_workbook()
                st.rerun()

    if st.session_state.get("workbook_bytes") is None:
        st.markdown("### Upload workbook")
        st.write("Choose the latest NPL workbook to open the review workspace.")
        main_upload = st.file_uploader("Select workbook", type=["xlsx"], key="main_workbook_upload")
        if main_upload is not None:
            stash_uploaded_workbook(main_upload)
            st.rerun()
        return None, None
    return st.session_state.get("workbook_bytes"), st.session_state.get("workbook_name")


def apply_filters(display_df: pd.DataFrame) -> tuple[pd.DataFrame, date, str, dict[str, Any]]:
    with st.sidebar:
        st.header("Workspace")
        workspace_mode = st.radio("Main view", ["Review workspace", "Overview", "Bulk update"], index=0)
        meeting_mode = st.toggle("Meeting mode", value=True, help="Hides extra chrome and keeps the review area presentation-first.")
        comment_date = st.date_input("Current week comment date", value=st.session_state.get("comment_date", date.today()))
        st.session_state["comment_date"] = comment_date

        st.header("Filters")
        show_hidden = st.checkbox("Include workbook-hidden rows", value=False)
        asset_managers = sorted([x for x in display_df["Asset Manager"].dropna().astype(str).unique().tolist() if x])
        am_spotlight = st.toggle("AM spotlight", value=False, help="Focus the queue on one asset manager for live review.")
        spotlight_am = st.selectbox("Spotlight AM", [""] + asset_managers, disabled=not am_spotlight)
        selected_ams = st.multiselect("Asset Manager", asset_managers, default=[spotlight_am] if am_spotlight and spotlight_am else [])
        selected_bt = st.multiselect(
            "Bridge / Term",
            sorted([x for x in display_df["Bridge / Term"].dropna().astype(str).unique().tolist() if x]),
        )
        selected_segment = st.multiselect(
            "Segment",
            sorted([x for x in display_df["Segment"].dropna().astype(str).unique().tolist() if x]),
        )
        queue_view = st.selectbox(
            "Queue focus",
            ["All deals", "Open only", "Missing current-week comment", "Needs discussion", "Has data flags"],
        )
        sort_mode = st.selectbox(
            "Sort queue",
            ["Asset Manager / Deal", "UPB desc", "Flag count desc", "Last comment date desc"],
        )
        search = st.text_input("Search deal number, name, or location")

    filtered_df = display_df.copy()
    if not show_hidden:
        filtered_df = filtered_df[~filtered_df["_workbook_hidden"]]
    if selected_ams:
        filtered_df = filtered_df[filtered_df["Asset Manager"].astype(str).isin(selected_ams)]
    if selected_bt:
        filtered_df = filtered_df[filtered_df["Bridge / Term"].astype(str).isin(selected_bt)]
    if selected_segment:
        filtered_df = filtered_df[filtered_df["Segment"].astype(str).isin(selected_segment)]
    if search:
        needle = search.lower()
        filtered_df = filtered_df[
            filtered_df["Deal Number"].astype(str).str.lower().str.contains(needle)
            | filtered_df["Deal Name"].astype(str).str.lower().str.contains(needle)
            | filtered_df["Location"].astype(str).str.lower().str.contains(needle)
        ]

    if queue_view == "Open only":
        filtered_df = filtered_df[filtered_df["Review Status"].astype(str).eq("Open")]
    elif queue_view == "Missing current-week comment":
        filtered_df = filtered_df[filtered_df["This Week Comment"].fillna("").astype(str).str.strip().eq("")]
    elif queue_view == "Needs discussion":
        filtered_df = filtered_df[filtered_df["Needs Discussion"].fillna(False)]
    elif queue_view == "Has data flags":
        filtered_df = filtered_df[filtered_df["Flag Count"] > 0]

    sorted_df = sort_deals(filtered_df, sort_mode).reset_index(drop=True)
    controls = {
        "workspace_mode": workspace_mode,
        "meeting_mode": meeting_mode,
        "show_hidden": show_hidden,
        "am_spotlight": am_spotlight,
        "spotlight_am": spotlight_am,
        "sort_mode": sort_mode,
    }
    return sorted_df, comment_date, queue_view, controls


def segmented_choice(label: str, options: list[Any], current_value: Any, key: str) -> Any:
    current = current_value if current_value in options else options[0]
    if hasattr(st, "segmented_control"):
        return st.segmented_control(label, options=options, selection_mode="single", default=current, key=key)
    return st.radio(label, options=options, index=options.index(current), horizontal=True, key=key)


def render_queue_navigation(sorted_df: pd.DataFrame) -> None:
    option_rows = sorted_df["_excel_row"].tolist()
    if not option_rows:
        return
    if st.session_state.get("selected_row") not in option_rows:
        st.session_state["selected_row"] = option_rows[0]

    current_idx = option_rows.index(st.session_state["selected_row"])
    reviewed_count = len(st.session_state.get("reviewed_rows", set()).intersection(set(option_rows)))
    commented_count = len(st.session_state.get("commented_rows", set()).intersection(set(option_rows)))

    nav1, nav2, nav3, nav4 = st.columns([0.8, 0.8, 3.9, 1.5])
    with nav1:
        if st.button("Previous", disabled=current_idx == 0, use_container_width=True):
            st.session_state["selected_row"] = option_rows[current_idx - 1]
            st.rerun()
    with nav2:
        if st.button("Next", disabled=current_idx == len(option_rows) - 1, use_container_width=True):
            st.session_state["selected_row"] = option_rows[current_idx + 1]
            st.rerun()
    with nav3:
        selected_row = st.selectbox(
            "Deal queue",
            option_rows,
            format_func=lambda row_num: format_deal_label(sorted_df[sorted_df["_excel_row"] == row_num].iloc[0]),
            index=option_rows.index(st.session_state["selected_row"]),
            label_visibility="collapsed",
        )
        st.session_state["selected_row"] = selected_row
    with nav4:
        st.markdown(
            f"<div class='banner'><strong>Session progress</strong><br>Reviewed: {reviewed_count}/{len(option_rows)}<br>Commented: {commented_count}</div>",
            unsafe_allow_html=True,
        )


def render_deal_header(selected: pd.Series, queue_position: int, queue_total: int) -> None:
    subtitle = f"{safe_text(selected.get('Location'))} · {safe_text(selected.get('Bridge / Term'))} · Queue {queue_position} of {queue_total}"
    st.markdown(
        f"""
        <div class="deal-header">
            <div>
                <span class="deal-title">{safe_text(selected.get('Deal Name'))}</span>
                <span class="deal-number">#{safe_text(selected.get('Deal Number'))}</span>
                <span class="am-badge">AM: {safe_text(selected.get('Asset Manager'))}</span>
            </div>
            <div class="deal-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_strip(selected: pd.Series) -> None:
    cols = st.columns(7)
    cols[0].metric("UPB", metric_currency(selected.get("Current UPB")))
    cols[1].metric("As-Is", metric_currency(selected.get("Salesforce As-Is Valuation")))
    cols[2].metric("As-Is LTV", metric_pct(selected.get("Salesforce Implied As-Is LTV")))
    cols[3].metric("ARV", metric_currency(selected.get("Salesforce ARV")))
    cols[4].metric("ARV LTV", metric_pct(selected.get("Salesforce ARV LTV")))
    cols[5].metric("DQ", safe_text(selected.get("Current DQ Status")))
    cols[6].metric("Last comment", metric_date(selected.get("Last Comment Date")))
    st.markdown(
        f"<div class='comment-card'><div class='section-label'>Previous weekly comment</div><div class='comment-value'>{safe_text(selected.get('Previous Weekly Comment'))}</div></div>",
        unsafe_allow_html=True,
    )


def render_status_chips(selected: pd.Series) -> None:
    chip_text = []
    for field in [
        "Current DQ Status",
        "6/30 NPL",
        "Resolution Likelihood",
        "Expected Final Resolution",
        "Resolution Timing",
        "Liquidity Event Timing",
        "Review Status",
    ]:
        value = safe_text(selected.get(field))
        if value != "-":
            label = field.replace("Current ", "").replace("Expected ", "").replace("Timing", "Tm")
            chip_text.append(f"<span class='status-chip'>{label}: {value}</span>")
    if bool(selected.get("Needs Discussion", False)):
        chip_text.append("<span class='status-chip'>Needs discussion: Yes</span>")
    if chip_text:
        st.markdown("".join(chip_text), unsafe_allow_html=True)


def render_mini_card(title: str, value: Any) -> None:
    st.markdown(
        f"<div class='mini-card'><div class='mini-label'>{title}</div><div class='mini-value'>{safe_text(value)}</div></div>",
        unsafe_allow_html=True,
    )


def render_context_column(selected: pd.Series) -> None:
    render_mini_card("Location", selected.get("Location"))
    render_mini_card("Type", selected.get("Type"))
    render_mini_card("Segment", selected.get("Segment"))
    render_mini_card("Financing", selected.get("Financing"))
    render_mini_card("Commitment", metric_currency(selected.get("Loan Commitment")))
    render_mini_card("Remaining commitment", metric_currency(selected.get("Remaining Commitment")))
    render_mini_card("Units", selected.get("Number Of Units"))
    render_mini_card("Maturity", metric_date(selected.get("Maturity Date")))
    render_mini_card("Next payment", metric_date(selected.get("Next Payment Date")))
    render_mini_card("Convention", selected.get("MBA Loan List Convention"))
    if selected.get("Flag Count", 0) > 0:
        st.warning(selected.get("Flag Summary") or "This deal has data quality flags.")
    else:
        st.success("No data flags detected for the current row.")


def render_secondary_detail_popovers(selected: pd.Series, properties_df: pd.DataFrame, preview_text: str) -> None:
    labels = ["Property detail", "Prior history", "Salesforce AM comment", "Export preview"]
    if hasattr(st, "popover"):
        pop1, pop2, pop3, pop4 = st.columns(4)
        with pop1:
            with st.popover(labels[0], use_container_width=True):
                render_property_table(selected, properties_df)
        with pop2:
            with st.popover(labels[1], use_container_width=True):
                st.text_area("Comment history", editable_text(selected.get("Comment History")), disabled=True, height=260)
        with pop3:
            with st.popover(labels[2], use_container_width=True):
                st.text_area("Current Salesforce AM Comment", editable_text(selected.get("Current Salesforce AM Comment")), disabled=True, height=220)
        with pop4:
            with st.popover(labels[3], use_container_width=True):
                st.text_area("Comments export preview", preview_text, disabled=True, height=260)
    else:
        with st.expander("Secondary detail", expanded=False):
            render_property_table(selected, properties_df)
            st.text_area("Comment history", editable_text(selected.get("Comment History")), disabled=True, height=180)
            st.text_area("Current Salesforce AM Comment", editable_text(selected.get("Current Salesforce AM Comment")), disabled=True, height=140)
            st.text_area("Comments export preview", preview_text, disabled=True, height=180)


def render_property_table(selected: pd.Series, properties_df: pd.DataFrame) -> None:
    if pd.notna(selected.get("Property Count")):
        units = normalize_number(selected.get("Property Units"))
        units_text = f"{int(units):,}" if units is not None else "-"
        st.caption(
            f"Property roll-up: {int(selected.get('Property Count', 0))} properties · Units {units_text} · Max DPD {safe_text(selected.get('Max Days Past Due'))}"
        )
    if not properties_df.empty and "Deal Number" in properties_df.columns:
        related = properties_df[properties_df["Deal Number"] == selected["Deal Number"]].copy()
        if not related.empty:
            show_cols = [
                col for col in ["Address", "City", "State", "# Units", "Current UPB", "Salesforce As-Is Valuation", "Salesforce ARV", "Days Past Due"] if col in related.columns
            ]
            st.dataframe(related[show_cols], use_container_width=True, hide_index=True)
        else:
            st.caption("No property rows found for this deal in the property sheet.")


def editable_field_input(field: str, current_value: str, picklists: dict[str, list[str]], row_key: str) -> Any:
    if field == "Resolution Likelihood":
        return segmented_choice(field, LIKELIHOOD_OPTIONS, current_value, f"{row_key}_{field}")
    if field in BINARY_FIELDS:
        current = current_value if current_value in {"", "Y", "N", "Yes", "No", "N/A"} else ""
        options = ["", "Y", "N", "N/A"]
        return segmented_choice(field, options, current, f"{row_key}_{field}")
    if field in picklists:
        options = field_options(field, picklists, current_value)
        return st.selectbox(field, options, index=selected_index(options, current_value), key=f"{row_key}_{field}")
    if field in LONG_TEXT_FIELDS or len(current_value) > 80:
        return st.text_area(field, current_value, height=86, key=f"{row_key}_{field}")
    return st.text_input(field, current_value, key=f"{row_key}_{field}")


def render_validation_banner(display_df: pd.DataFrame) -> None:
    missing_comment = int(display_df["This Week Comment"].fillna("").astype(str).str.strip().eq("").sum()) if "This Week Comment" in display_df.columns else 0
    needs_discussion = int(display_df["Needs Discussion"].fillna(False).sum()) if "Needs Discussion" in display_df.columns else 0
    flagged = int((display_df["Flag Count"] > 0).sum()) if "Flag Count" in display_df.columns else 0
    msg = f"Queue checks: {missing_comment} deal(s) missing a current-week comment, {needs_discussion} marked needs discussion, {flagged} with data flags."
    level = st.info
    if flagged > 0 or missing_comment > 0:
        level = st.warning
    level(msg)


def render_review_form(selected: pd.Series, picklists: dict[str, list[str]], properties_df: pd.DataFrame, comment_date: date, meeting_mode: bool) -> dict[str, Any] | None:
    render_metric_strip(selected)
    render_status_chips(selected)

    initial_preview = editable_text(selected.get("Comments Preview"))
    render_secondary_detail_popovers(selected, properties_df, initial_preview)

    with st.form(f"detail_form_{int(selected['_excel_row'])}"):
        updates: dict[str, Any] = {}
        row_key = f"row_{int(selected['_excel_row'])}"

        summary_col, manual_col, right_col = st.columns([1.15, 1.35, 1.0])

        with summary_col:
            st.markdown("<div class='hero-card'><div class='section-label'>Deal review snapshot</div>", unsafe_allow_html=True)
            snapshot_lines = [
                f"Location / Type: {safe_text(selected.get('Location'))} · {safe_text(selected.get('Type'))}",
                f"Segment / Financing: {safe_text(selected.get('Segment'))} · {safe_text(selected.get('Financing'))}",
                f"Current UPB / As-Is / ARV: {metric_currency(selected.get('Current UPB'))} · {metric_currency(selected.get('Salesforce As-Is Valuation'))} · {metric_currency(selected.get('Salesforce ARV'))}",
                f"Maturity / Next payment: {metric_date(selected.get('Maturity Date'))} · {metric_date(selected.get('Next Payment Date'))}",
                f"Resolution path: {safe_text(selected.get('Expected Resolution Type'))} / {safe_text(selected.get('Expected Final Resolution'))}",
                f"Additional NPL comment: {safe_text(selected.get("Addt'l NPL Comment"))}",
            ]
            st.markdown(f"<div class='hero-value'>{'<br>'.join(snapshot_lines)}</div></div>", unsafe_allow_html=True)

            st.markdown("<div class='comment-card'><div class='section-label'>Current week update</div>", unsafe_allow_html=True)
            template_options = field_options("Comment Template", {"Comment Template": COMMENT_TEMPLATE_OPTIONS}, "")
            comment_template = st.selectbox("Quick comment helper", template_options, index=0, key=f"{row_key}_comment_template")
            current_week = st.text_area(
                "This week comment",
                editable_text(selected.get("This Week Comment")),
                height=180 if meeting_mode else 150,
                help="This will be prepended to the existing Comments history on export.",
                key=f"{row_key}_this_week",
            )
            built_comment = build_weekly_comment_text(current_week, comment_template)
            updates["This Week Comment"] = built_comment
            st.markdown("</div>", unsafe_allow_html=True)

        with manual_col:
            st.markdown("### Manual updates")
            for section_title, fields in SECTION_MAP.items():
                present_fields = [field for field in fields if field in selected.index]
                if not present_fields:
                    continue
                with st.container(border=True):
                    st.markdown(f"**{section_title}**")
                    for field in present_fields:
                        updates[field] = editable_field_input(field, editable_text(selected.get(field)), picklists, row_key)

        with right_col:
            st.markdown("### Meeting controls")
            updates["Review Status"] = segmented_choice(
                "Meeting status",
                MEETING_STATUS_OPTIONS,
                str(selected.get("Review Status", "Open")),
                f"{row_key}_review_status",
            )
            updates["Needs Discussion"] = segmented_choice(
                "Needs discussion",
                [False, True],
                bool(selected.get("Needs Discussion", False)),
                f"{row_key}_needs_discussion",
            )
            render_context_column(selected)
            st.markdown("<div class='panel-card'><div class='section-label'>Valuation / offer snapshot</div>", unsafe_allow_html=True)
            valuation_lines = [
                f"Valuation Type: {safe_text(selected.get('Valuation Type'))}",
                f"As-Is Value: {metric_currency(selected.get('As Is Appraised Value'))}",
                f"As Stabilized: {metric_currency(selected.get('As Stabilized Value'))}",
                f"Valuation Date: {metric_date(selected.get('Date of Valuation'))}",
                f"Finalized?: {safe_text(selected.get('Is Appraisal Finalized? (Y/N)'))}",
                f"Third Party Offer: {safe_text(selected.get('Third Party Offer Amount'))}",
                f"Offer Note: {safe_text(selected.get('Third Party Offer Note'))}",
            ]
            st.markdown(f"<div class='hero-value'>{'<br>'.join(valuation_lines)}</div></div>", unsafe_allow_html=True)

        preview_text = (
            f"{format_comment_entry_date(comment_date)} - {built_comment}\n{editable_text(selected.get('Comment History'))}".strip()
            if built_comment
            else editable_text(selected.get("Comments Preview"))
        )
        st.caption(
            f"Export preview will prepend: {format_comment_entry_date(comment_date)} - {built_comment}" if built_comment else "No current-week comment entered yet."
        )
        st.text_area("Comments export preview", preview_text, disabled=True, height=160 if meeting_mode else 130)

        submitted = st.form_submit_button("Apply deal edits", use_container_width=True)
        if submitted:
            return updates
    return None


def render_overview(filtered_df: pd.DataFrame, snapshot_label: str | None) -> None:
    st.subheader("Portfolio overview")
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Deals in view", f"{len(filtered_df):,}")
    col2.metric("Total UPB", metric_currency(numeric_sum(filtered_df["Current UPB"])))
    col3.metric("Asset Managers", f"{filtered_df['Asset Manager'].nunique():,}")
    col4.metric("Blank current comments", f"{filtered_df['This Week Comment'].fillna('').astype(str).str.strip().eq('').sum():,}")
    col5.metric("Deals with flags", f"{(filtered_df['Flag Count'] > 0).sum():,}")
    if snapshot_label:
        st.caption(f"Workbook snapshot detected from the active sheet: {snapshot_label}")
    am_summary = (
        filtered_df.groupby("Asset Manager", dropna=False)
        .agg(
            Deals=("Deal Number", "count"),
            Total_UPB=("Current UPB", lambda s: numeric_sum(s)),
            Blank_Current_Comments=("This Week Comment", lambda s: s.fillna("").astype(str).str.strip().eq("").sum()),
            Needs_Discussion=("Needs Discussion", "sum"),
            Flagged=("Flag Count", lambda s: (s > 0).sum()),
        )
        .reset_index()
    )
    am_summary["Total UPB"] = am_summary["Total_UPB"].apply(metric_currency)
    am_summary = am_summary.drop(columns=["Total_UPB"]).rename(columns={"Blank_Current_Comments": "Blank current comments", "Needs_Discussion": "Needs discussion"})
    st.dataframe(am_summary, use_container_width=True, hide_index=True)


def render_bulk_edit(filtered_df: pd.DataFrame, picklists: dict[str, list[str]]) -> pd.DataFrame | None:
    st.subheader("Bulk update grid")
    st.caption("Use this for queue clean-up. Prior comment remains read-only; current-week comment is what gets prepended on export.")
    default_cols = [
        "Resolution Likelihood",
        "Expected Resolution Type",
        "Expected Final Resolution",
        "Resolution Timing",
        "Liquidity Event Timing",
        "Pref Deal (Y/N)",
        "Valuation Type",
        "Previous Weekly Comment",
        "This Week Comment",
        "Review Status",
        "Needs Discussion",
    ]
    available_bulk = [col for col in default_cols if col in filtered_df.columns]
    selected_cols = st.multiselect("Columns shown in grid", available_bulk, default=available_bulk)
    grid_columns = [col for col in [*GRID_CONTEXT_HEADERS, *selected_cols] if col in filtered_df.columns]
    editor_view = filtered_df[["_excel_row", *grid_columns]].copy()
    disabled_columns = ["_excel_row", *GRID_CONTEXT_HEADERS, "Previous Weekly Comment"]
    column_config: dict[str, Any] = {
        "Needs Discussion": st.column_config.CheckboxColumn("Needs Discussion"),
        "Review Status": st.column_config.SelectboxColumn("Review Status", options=MEETING_STATUS_OPTIONS),
    }
    for field, options in picklists.items():
        if field in editor_view.columns and field not in {"Comment Template"}:
            column_config[field] = st.column_config.SelectboxColumn(field, options=options)
    if "This Week Comment" in editor_view.columns:
        column_config["This Week Comment"] = st.column_config.TextColumn("This Week Comment", width="large")
    edited = st.data_editor(
        editor_view,
        use_container_width=True,
        hide_index=True,
        disabled=disabled_columns,
        num_rows="fixed",
        column_config=column_config,
        key="bulk_editor_v8",
    )
    if st.button("Apply bulk grid changes", use_container_width=True):
        return edited
    return None


def main() -> None:
    st.set_page_config(page_title="NPL Deal Review", layout="wide", initial_sidebar_state="expanded")
    apply_base_css()

    file_bytes, workbook_name = get_active_workbook()
    if file_bytes is None:
        return

    file_hash = hashlib.md5(file_bytes).hexdigest()
    parsed = cached_parse_workbook(file_bytes)
    picklists = cached_picklists(parsed["deals_df"])

    if st.session_state.get("file_hash") != file_hash:
        init_state(file_hash, parsed)

    comment_date = st.session_state.get("comment_date", date.today())
    st.session_state["editor_df"] = refresh_editor_derived_fields(st.session_state["editor_df"], comment_date)

    display_df = build_display_view(
        base_deals=parsed["deals_df"],
        editor_df=st.session_state["editor_df"],
        property_summary=parsed["property_summary"],
        entry_date=comment_date,
        picklists=picklists,
    )

    sorted_df, comment_date, queue_view, controls = apply_filters(display_df)
    st.session_state["editor_df"] = refresh_editor_derived_fields(st.session_state["editor_df"], comment_date)

    if sorted_df.empty:
        st.warning("No deals match the current filters.")
        return

    if controls["workspace_mode"] != "Review workspace":
        st.title("NPL deal review")
        st.caption("Open overview or bulk update from the sidebar when needed. Review workspace remains the main operating screen.")

    if controls["workspace_mode"] == "Overview":
        render_validation_banner(sorted_df)
        render_overview(sorted_df, parsed.get("snapshot_label"))
    elif controls["workspace_mode"] == "Bulk update":
        render_validation_banner(sorted_df)
        edited_subset = render_bulk_edit(sorted_df, picklists)
        if edited_subset is not None:
            full_editor = st.session_state["editor_df"].set_index("_excel_row")
            edited_lookup = edited_subset.set_index("_excel_row")
            for row_num in edited_lookup.index:
                if row_num not in full_editor.index:
                    continue
                for header in edited_lookup.columns:
                    if header in {"Comment History", "Comments Preview", "Previous Weekly Comment", "Last Comment Date"}:
                        continue
                    if header in full_editor.columns:
                        full_editor.loc[row_num, header] = edited_lookup.loc[row_num, header]
            st.session_state["editor_df"] = refresh_editor_derived_fields(full_editor.reset_index(), comment_date)
            st.success("Bulk grid changes applied in the app view.")
            st.rerun()
    else:
        render_queue_navigation(sorted_df)
        selected = sorted_df[sorted_df["_excel_row"] == st.session_state["selected_row"]].iloc[0]
        queue_position = sorted_df.index[sorted_df["_excel_row"] == st.session_state["selected_row"]][0] + 1
        queue_total = len(sorted_df)
        render_deal_header(selected, queue_position, queue_total)
        updates = render_review_form(selected, picklists, parsed["properties_df"], comment_date, controls["meeting_mode"])
        if updates:
            st.session_state["editor_df"] = apply_detail_edit(st.session_state["editor_df"], int(selected["_excel_row"]), updates, comment_date)
            st.session_state.setdefault("reviewed_rows", set()).add(int(selected["_excel_row"]))
            if str(updates.get("This Week Comment", "")).strip():
                st.session_state.setdefault("commented_rows", set()).add(int(selected["_excel_row"]))
            st.rerun()

    st.divider()
    changes = workbook_change_set(st.session_state["editor_df"], parsed["original_editor_values"], comment_date)
    changed_deals = len({row for row, _ in changes.keys()})
    render_validation_banner(display_df)

    left, right = st.columns([2.2, 1.2])
    with left:
        st.write(f"Pending edits in the app view: **{len(changes)} cells across {changed_deals} deal(s)**")
        if queue_view != "All deals":
            st.caption(f"Queue focus currently applied: {queue_view}")
    with right:
        if st.button("Reset all app edits", use_container_width=True):
            st.session_state["editor_df"] = make_editor_df(parsed["deals_df"])
            st.session_state["reviewed_rows"] = set()
            st.session_state["commented_rows"] = set()
            st.rerun()

    updated_bytes = build_updated_workbook(
        file_bytes=file_bytes,
        sheet_path=parsed["deal_sheet_path"],
        changes=changes,
        cell_meta=parsed["cell_meta"],
    ) if changes else file_bytes

    base_name = (workbook_name or "NPL Workbook").rsplit(".", 1)[0]
    download_name = f"{base_name} - AM Updates.xlsx"
    st.download_button(
        "Download updated workbook",
        data=updated_bytes,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
