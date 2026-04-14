from __future__ import annotations

import hashlib
import re
from datetime import date
from typing import Any

import pandas as pd
import streamlit as st

from workbook_utils import (
    COMMENT_TEMPLATE_OPTIONS,
    COMMENT_STORAGE_FIELD,
    GRID_CONTEXT_HEADERS,
    LONG_TEXT_FIELDS,
    MANUAL_ENTRY_PICKLIST_FIELDS,
    NUMERIC_OR_TEXT_FIELDS,
    SECTION_MAP,
    apply_detail_edit,
    build_display_view,
    build_picklists,
    build_updated_workbook,
    build_weekly_comment_text,
    editable_text,
    format_deal_label,
    format_comment_entry_date,
    make_editor_df,
    metric_currency,
    metric_date,
    metric_pct,
    normalize_money_text,
    normalize_number,
    parse_deal_workbook,
    refresh_editor_derived_fields,
    sort_deals,
    workbook_change_set,
)


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
            .block-container {
                padding-top: 0.7rem;
                padding-bottom: 1rem;
                max-width: 100%;
            }
            [data-testid="stSidebar"] {display: none !important;}
            [data-testid="collapsedControl"] {display: none !important;}
            section[data-testid="stSidebar"] {display: none !important;}
            h1 {font-size: 1.35rem !important; margin-bottom: 0.1rem !important;}
            h2 {font-size: 1.05rem !important; margin-top: 0.35rem !important;}
            h3 {font-size: 0.95rem !important; margin-top: 0.25rem !important;}
            div[data-testid="stMetricValue"] {font-size: 1.25rem !important; font-weight: 900 !important;}
            div[data-testid="stMetricLabel"] {font-size: 0.86rem !important; font-weight: 800 !important;}
            .side-panel {
                border: 1px solid #dfe6ef;
                border-radius: 14px;
                padding: 0.9rem 0.9rem 0.7rem 0.9rem;
                background: #fbfcfe;
                position: sticky;
                top: 0.75rem;
            }
            .deal-header {
                border: 1px solid #d7dbe2;
                border-radius: 12px;
                padding: 0.7rem 0.9rem;
                margin-bottom: 0.6rem;
                background: #fbfcfe;
            }
            .deal-title {
                font-size: 1.95rem;
                font-weight: 950;
                line-height: 1.08;
                margin-right: 0.55rem;
            }
            .deal-number {
                font-size: 1.08rem;
                font-weight: 850;
                color: #475467;
                margin-right: 0.5rem;
            }
            .am-badge {
                display: inline-block;
                padding: 0.28rem 0.7rem;
                border-radius: 999px;
                border: 1px solid #c9daf8;
                background: #eef4ff;
                font-size: 1.12rem;
                font-weight: 950;
            }
            .deal-subtitle {
                margin-top: 0.2rem;
                color: #5c6570;
                font-size: 0.78rem;
                line-height: 1.2;
            }
            .small-card {
                border: 1px solid #e5e7eb;
                border-radius: 10px;
                padding: 0.45rem 0.6rem;
                margin-bottom: 0.45rem;
                background: #ffffff;
            }
            .small-card .label {
                color: #6b7280;
                font-size: 0.72rem;
                text-transform: uppercase;
                letter-spacing: 0.02em;
            }
            .small-card .value {
                font-size: 0.92rem;
                font-weight: 600;
                margin-top: 0.12rem;
            }
            .presenter-panel {
                border: 1px solid #dfe6ef;
                border-radius: 12px;
                padding: 0.8rem 0.95rem;
                background: white;
                height: 100%;
            }
            .kpi-grid {
                display: grid;
                grid-template-columns: repeat(7, minmax(110px, 1fr));
                gap: 0.45rem;
                margin-bottom: 0.45rem;
            }
            .kpi-box {
                border: 1px solid #dfe6ef;
                border-radius: 12px;
                padding: 0.72rem 0.78rem;
                background: white;
                box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
            }
            .kpi-box .label {
                color: #667085;
                font-size: 0.8rem;
                text-transform: uppercase;
                letter-spacing: 0.03em;
                font-weight: 800;
            }
            .kpi-box .value {
                margin-top: 0.18rem;
                font-size: 1.42rem;
                font-weight: 950;
                line-height: 1.05;
            }
            .kpi-comment {
                border: 1px solid #e5e7eb;
                border-radius: 10px;
                padding: 0.55rem 0.7rem;
                background: #ffffff;
                margin-top: 0.45rem;
                min-height: 5.4rem;
            }
            .kpi-comment .label {
                color: #6b7280;
                font-size: 0.72rem;
                text-transform: uppercase;
                letter-spacing: 0.02em;
                margin-bottom: 0.2rem;
            }
            .kpi-comment .value {
                font-size: 0.95rem;
                line-height: 1.32;
                white-space: pre-wrap;
                font-weight: 600;
            }
            .presenter-section-title {
                font-size: 0.86rem;
                font-weight: 700;
                margin: 0.1rem 0 0.35rem 0;
            }
            .snapshot-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 0.85rem;
                margin-top: 0.35rem;
            }
            .snapshot-card {
                border: 1px solid #dce5f0;
                border-radius: 14px;
                padding: 0.75rem 0.85rem;
                background: #fcfdff;
                min-height: 100%;
            }
            .snapshot-section-title {
                color: #0f172a;
                font-size: 0.9rem;
                text-transform: uppercase;
                letter-spacing: 0.03em;
                margin-bottom: 0.4rem;
                font-weight: 800;
            }
            .snapshot-row {
                display: grid;
                grid-template-columns: 10rem 1fr;
                gap: 0.35rem;
                padding: 0.16rem 0;
                border-top: 1px solid #edf2f7;
            }
            .snapshot-row:first-child {
                border-top: none;
                padding-top: 0;
            }
            .snapshot-label {
                color: #667085;
                font-size: 0.73rem;
                text-transform: uppercase;
                letter-spacing: 0.02em;
                font-weight: 700;
            }
            .snapshot-value {
                font-size: 1.03rem;
                line-height: 1.35;
                font-weight: 800;
                color: #0f172a;
                white-space: pre-wrap;
            }
            .action-bar {
                border: 1px solid #dfe6ef;
                border-radius: 12px;
                padding: 0.6rem 0.75rem;
                background: #f8fbff;
                margin-top: 0.35rem;
            }
            .status-chip {
                display: inline-block;
                padding: 0.14rem 0.4rem;
                border-radius: 999px;
                border: 1px solid #d1d5db;
                margin-right: 0.25rem;
                margin-bottom: 0.2rem;
                font-size: 0.75rem;
                background: #f9fafb;
                font-weight: 600;
            }
            header, footer {visibility: hidden;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def safe_text(value: Any) -> str:
    text = editable_text(value)
    return text if text else "-"


def infer_quarter_label(workbook_name: str | None, snapshot_label: str | None) -> str:
    name = (workbook_name or "").strip()
    patterns = [
        r"(\b[1-4]Q\d{2}\b)",
        r"(\b[1-4]Q\d{4}\b)",
    ]
    for pattern in patterns:
        match = re.search(pattern, name, flags=re.IGNORECASE)
        if match:
            return match.group(1).upper()
    today = date.today()
    q = ((today.month - 1) // 3) + 1
    yy = str(today.year)[-2:]
    return f"{q}Q{yy}"


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
        "sidebar_workbook_upload",
    ]:
        st.session_state.pop(key, None)


def get_active_workbook(container) -> tuple[bytes | None, str | None]:
    with container:
        st.markdown("### Workbook")
        if st.session_state.get("workbook_name"):
            st.caption(f"Current workbook: {st.session_state['workbook_name']}")
            if st.button("Clear workbook", width="stretch", key="clear_workbook_btn"):
                clear_uploaded_workbook()
                st.rerun()

        upload = st.file_uploader(
            "Upload workbook",
            type=["xlsx"],
            key="sidebar_workbook_upload",
            help="Choose the latest NPL workbook to open the review workspace.",
        )
        if upload is not None:
            stash_uploaded_workbook(upload)

        if st.session_state.get("workbook_bytes") is None:
            st.info("Upload the workbook to begin.")
            return None, None

    return st.session_state.get("workbook_bytes"), st.session_state.get("workbook_name")


def apply_filters(display_df: pd.DataFrame, controls_container) -> tuple[pd.DataFrame, date, str, str, str]:
    with controls_container:
        st.markdown("### Review controls")
        st.caption("Choose presentation order and deal scope here.")

        workspace_mode = st.radio(
            "Workspace",
            ["Review workspace", "Overview", "Bulk update"],
            index=0,
            key="workspace_mode_radio",
        )
        comment_date = st.date_input(
            "Current week comment date",
            value=st.session_state.get("comment_date", date.today()),
            key="comment_date_input",
        )
        st.session_state["comment_date"] = comment_date
        show_hidden = st.checkbox(
            "Include workbook-hidden rows",
            value=False,
            key="show_hidden_checkbox",
        )

        asset_managers = sorted([x for x in display_df["Asset Manager"].dropna().astype(str).unique().tolist() if x])

        presentation_order = st.selectbox(
            "Presentation order",
            ["Workbook file order", "Asset Manager order"],
            index=0,
            key="presentation_order_select",
            help="Choose whether the queue follows the uploaded workbook order or groups deals by Asset Manager.",
        )

        am_scope = st.selectbox(
            "Show deals for",
            ["All asset managers"] + asset_managers if asset_managers else ["All asset managers"],
            index=0,
            key="am_scope_select",
            help="Pick one Asset Manager here to review only that AM's deals.",
        )

        selected_bt = st.multiselect(
            "Bridge / Term",
            sorted([x for x in display_df["Bridge / Term"].dropna().astype(str).unique().tolist() if x]),
            key="bridge_term_multiselect",
        )
        selected_segment = st.multiselect(
            "Segment",
            sorted([x for x in display_df["Segment"].dropna().astype(str).unique().tolist() if x]),
            key="segment_multiselect",
        )
        queue_view = st.selectbox(
            "Queue focus",
            [
                "All deals",
                "Missing current-week comment",
                "Needs discussion",
                "Has data flags",
            ],
            key="queue_focus_select",
        )
        manual_sort = st.selectbox(
            "Secondary sort",
            [
                "Use presentation order only",
                "Deal Name",
                "UPB desc",
                "Flag count desc",
                "Last comment date desc",
            ],
            index=0,
            key="secondary_sort_select",
        )
        search = st.text_input(
            "Search deal number, name, or location",
            key="deal_search_input",
        )

    filtered_df = display_df.copy()
    if not show_hidden:
        filtered_df = filtered_df[~filtered_df["_workbook_hidden"]]

    if am_scope != "All asset managers":
        filtered_df = filtered_df[filtered_df["Asset Manager"].astype(str).eq(am_scope)]

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

    if queue_view == "Missing current-week comment":
        filtered_df = filtered_df[filtered_df["This Week Comment"].fillna("").astype(str).str.strip().eq("")]
    elif queue_view == "Needs discussion":
        filtered_df = filtered_df[filtered_df["Needs Discussion"].fillna(False)]
    elif queue_view == "Has data flags":
        filtered_df = filtered_df[filtered_df["Flag Count"] > 0]

    selected_sort = "Workbook file order" if presentation_order == "Workbook file order" else "Asset Manager / Deal"
    if manual_sort == "Deal Name":
        selected_sort = "Deal Name"
    elif manual_sort == "UPB desc":
        selected_sort = "UPB desc"
    elif manual_sort == "Flag count desc":
        selected_sort = "Flag count desc"
    elif manual_sort == "Last comment date desc":
        selected_sort = "Last comment date desc"

    sorted_df = sort_deals(filtered_df, selected_sort).reset_index(drop=True)
    return sorted_df, comment_date, queue_view, workspace_mode, selected_sort


def render_queue_navigation(sorted_df: pd.DataFrame) -> None:
    option_rows = sorted_df["_excel_row"].tolist()
    if not option_rows:
        return

    if st.session_state.get("selected_row") not in option_rows:
        st.session_state["selected_row"] = option_rows[0]

    current_idx = option_rows.index(st.session_state["selected_row"])

    nav1, nav2, nav3 = st.columns([1, 1, 5])
    with nav1:
        if st.button("Previous", disabled=current_idx == 0, width="stretch"):
            st.session_state["selected_row"] = option_rows[current_idx - 1]
            st.rerun()
    with nav2:
        if st.button("Next", disabled=current_idx == len(option_rows) - 1, width="stretch"):
            st.session_state["selected_row"] = option_rows[current_idx + 1]
            st.rerun()
    with nav3:
        selected_row = st.selectbox(
            "Deal queue",
            option_rows,
            format_func=lambda row_num: format_deal_label(sorted_df[sorted_df["_excel_row"] == row_num].iloc[0]),
            index=option_rows.index(st.session_state["selected_row"]),
            key="deal_queue_select",
        )
        st.session_state["selected_row"] = selected_row


def render_deal_header(selected: pd.Series, queue_position: int, queue_total: int) -> None:
    subtitle = (
        f"{safe_text(selected.get('Location'))} · {safe_text(selected.get('Bridge / Term'))} · "
        f"Queue {queue_position} of {queue_total}"
    )
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
    items = [
        ("UPB", metric_currency(selected.get("Current UPB"))),
        ("As-Is", metric_currency(selected.get("Salesforce As-Is Valuation"))),
        ("As-Is LTV", metric_pct(selected.get("Salesforce Implied As-Is LTV"))),
        ("ARV", metric_currency(selected.get("Salesforce ARV"))),
        ("ARV LTV", metric_pct(selected.get("Salesforce ARV LTV"))),
        ("DQ", safe_text(selected.get("Current DQ Status"))),
        ("Last comment date", metric_date(selected.get("Last Comment Date"))),
    ]
    html = ["<div class='kpi-grid'>"]
    for label, value in items:
        html.append(
            f"<div class='kpi-box'><div class='label'>{label}</div><div class='value'>{value}</div></div>"
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)
    st.markdown(
        f"<div class='kpi-comment'><div class='label'>Previous weekly comment</div><div class='value'>{safe_text(selected.get('Previous Weekly Comment'))}</div></div>",
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
    ]:
        value = safe_text(selected.get(field))
        if value != "-":
            chip_text.append(
                f"<span class='status-chip'>{field.replace('Current ', '').replace('Expected ', '').replace('Timing', 'Tm')}: {value}</span>"
            )
    if chip_text:
        st.markdown("".join(chip_text), unsafe_allow_html=True)


def render_small_kv_card(title: str, value: Any) -> None:
    st.markdown(
        f"<div class='small-card'><div class='label'>{title}</div><div class='value'>{safe_text(value)}</div></div>",
        unsafe_allow_html=True,
    )


def render_context_column(selected: pd.Series) -> None:
    render_small_kv_card("Location", selected.get("Location"))
    render_small_kv_card("Type", selected.get("Type"))
    render_small_kv_card("Segment", selected.get("Segment"))
    render_small_kv_card("Financing", selected.get("Financing"))
    render_small_kv_card("Commitment", metric_currency(selected.get("Loan Commitment")))
    render_small_kv_card("Remaining commitment", metric_currency(selected.get("Remaining Commitment")))
    render_small_kv_card("Units", selected.get("Number Of Units"))
    render_small_kv_card("Maturity", metric_date(selected.get("Maturity Date")))
    render_small_kv_card("Next payment", metric_date(selected.get("Next Payment Date")))
    render_small_kv_card("Convention", selected.get("MBA Loan List Convention"))

    if selected.get("Flag Count", 0) > 0:
        st.warning(selected.get("Flag Summary") or "This deal has data quality flags.")
    else:
        st.success("No data flags detected for the current row.")


def render_property_summary(selected: pd.Series, properties_df: pd.DataFrame) -> None:
    if pd.notna(selected.get("Property Count")):
        units = normalize_number(selected.get("Property Units"))
        units_text = f"{int(units):,}" if units is not None else "-"
        st.caption(
            f"Property roll-up: {int(selected.get('Property Count', 0))} properties"
            f" · Units {units_text}"
            f" · Max DPD {safe_text(selected.get('Max Days Past Due'))}"
        )

    if not properties_df.empty and "Deal Number" in properties_df.columns:
        related = properties_df[properties_df["Deal Number"] == selected["Deal Number"]].copy()
        if not related.empty:
            show_cols = [
                col
                for col in [
                    "Address",
                    "City",
                    "State",
                    "# Units",
                    "Current UPB",
                    "Salesforce As-Is Valuation",
                    "Salesforce ARV",
                    "As Is Appraised Value",
                    "As Stabilized Value",
                    "Third Party Offer Amount",
                    "Days Past Due",
                ]
                if col in related.columns
            ]
            display_related = related[show_cols].copy()
            for col in display_related.columns:
                col_lower = str(col).lower()
                if any(token in col_lower for token in ["upb", "valuation", "value", "amount", "offer"]):
                    display_related[col] = display_related[col].apply(metric_currency)
            if "# Units" in display_related.columns:
                display_related["# Units"] = display_related["# Units"].apply(
                    lambda v: f"{int(v):,}" if normalize_number(v) is not None else safe_text(v)
                )
            if "Days Past Due" in display_related.columns:
                display_related["Days Past Due"] = display_related["Days Past Due"].apply(
                    lambda v: f"{int(v):,}" if normalize_number(v) is not None else safe_text(v)
                )
            with st.expander("Property detail", expanded=False):
                st.dataframe(display_related, width="stretch", hide_index=True)


def editable_field_input(field: str, current_value: str, picklists: dict[str, list[str]]) -> Any:
    if field in MANUAL_ENTRY_PICKLIST_FIELDS and field in picklists:
        options = field_options(field, picklists, current_value)
        custom_option = "Type custom value"
        select_options = options + ([custom_option] if custom_option not in options else [])
        default_value = current_value if current_value in select_options else custom_option
        default_index = selected_index(select_options, default_value)
        selected_option = st.selectbox(
            f"{field} options",
            select_options,
            index=default_index,
            key=f"{field}_select_{abs(hash((field, current_value)))}",
        )
        if selected_option == custom_option:
            return st.text_input(
                field,
                current_value,
                key=f"{field}_custom_{abs(hash((field, current_value, 'custom')))}",
            )
        return selected_option

    if field in NUMERIC_OR_TEXT_FIELDS:
        typed_value = st.text_input(
            field,
            normalize_money_text(current_value),
            placeholder="$0",
            help="Dollar values are normalized to whole dollars on save.",
        )
        preview = normalize_money_text(typed_value)
        if preview and preview != typed_value.strip():
            st.caption(f"Formatted value: {preview}")
        return typed_value

    if field in picklists:
        options = field_options(field, picklists, current_value)
        return st.selectbox(field, options, index=selected_index(options, current_value))

    if field in LONG_TEXT_FIELDS or len(current_value) > 80:
        return st.text_area(field, current_value, height=90)

    return st.text_input(field, current_value)


def render_snapshot_grid(selected: pd.Series) -> None:
    sections = [
        (
            "Loan snapshot",
            [
                ("Location", safe_text(selected.get("Location"))),
                ("Type / Segment", f"{safe_text(selected.get('Type'))} · {safe_text(selected.get('Segment'))}"),
                ("Financing", safe_text(selected.get("Financing"))),
                ("UPB", metric_currency(selected.get("Current UPB"))),
                ("Commitment", metric_currency(selected.get("Loan Commitment"))),
                ("Remaining commitment", metric_currency(selected.get("Remaining Commitment"))),
                ("Maturity", metric_date(selected.get("Maturity Date"))),
                ("Next payment", metric_date(selected.get("Next Payment Date"))),
                ("Units", safe_text(selected.get("Number Of Units"))),
                ("Property count", safe_text(selected.get("Property Count"))),
            ],
        ),
        (
            "Resolution + valuation / offer",
            [
                ("Likelihood", safe_text(selected.get("Resolution Likelihood"))),
                ("Resolution type", safe_text(selected.get("Expected Resolution Type"))),
                ("Final resolution", safe_text(selected.get("Expected Final Resolution"))),
                ("Resolution timing", safe_text(selected.get("Resolution Timing"))),
                ("Liquidity timing", safe_text(selected.get("Liquidity Event Timing"))),
                ("6/30 performance", safe_text(selected.get("Resolution for 6/30 Performance"))),
                ("Valuation type", safe_text(selected.get("Valuation Type"))),
                ("Valuation date", metric_date(selected.get("Date of Valuation"))),
                ("As-is appraisal", metric_currency(selected.get("As Is Appraised Value"))),
                ("As stabilized", metric_currency(selected.get("As Stabilized Value"))),
                ("Third-party offer", metric_currency(selected.get("Third Party Offer Amount"))),
                ("Offer note", safe_text(selected.get("Third Party Offer Note"))),
            ],
        ),
    ]
    html_parts = ["<div class='snapshot-grid'>"]
    for section_title, rows in sections:
        html_parts.append(f"<div class='snapshot-card'><div class='snapshot-section-title'>{section_title}</div>")
        for label, value in rows:
            html_parts.append(
                f"<div class='snapshot-row'><div class='snapshot-label'>{label}</div><div class='snapshot-value'>{value}</div></div>"
            )
        html_parts.append("</div>")
    html_parts.append("</div>")
    st.markdown("".join(html_parts), unsafe_allow_html=True)


def render_review_form(
    selected: pd.Series,
    picklists: dict[str, list[str]],
    properties_df: pd.DataFrame,
    comment_date: date,
) -> dict[str, Any] | None:
    render_metric_strip(selected)
    render_status_chips(selected)

    with st.form(f"detail_form_{int(selected['_excel_row'])}"):
        updates: dict[str, Any] = {}

        st.markdown("<div class='presenter-panel'>", unsafe_allow_html=True)
        st.markdown("<div class='presenter-section-title'>Deal review snapshot</div>", unsafe_allow_html=True)
        render_snapshot_grid(selected)
        st.markdown("</div>", unsafe_allow_html=True)
        render_property_summary(selected, properties_df)

        left, middle, right = st.columns([0.95, 1.35, 1.35])

        with left:
            st.subheader("Context")
            render_context_column(selected)
            updates["Needs Discussion"] = st.checkbox(
                "Needs discussion",
                value=bool(selected.get("Needs Discussion", False)),
            )

        with middle:
            st.subheader("Manual updates")
            for section_title, fields in SECTION_MAP.items():
                present_fields = [field for field in fields if field in selected.index and field != COMMENT_STORAGE_FIELD]
                if not present_fields:
                    continue
                st.markdown(f"**{section_title}**")
                for field in present_fields:
                    updates[field] = editable_field_input(field, editable_text(selected.get(field)), picklists)

        with right:
            st.subheader("Weekly comments")
            st.text_area(
                "Previous weekly comment (from Addt'l NPL Comment)",
                editable_text(selected.get("Previous Weekly Comment")),
                disabled=True,
                height=90,
            )
            with st.expander("Full prior comment history", expanded=False):
                st.text_area(
                    "Comment history",
                    editable_text(selected.get("Comment History")),
                    disabled=True,
                    height=220,
                )
            with st.expander("Salesforce AM comment", expanded=False):
                st.text_area(
                    "Current Salesforce AM Comment",
                    editable_text(selected.get("Current Salesforce AM Comment")),
                    disabled=True,
                    height=220,
                )

            template_options = field_options("Comment Template", {"Comment Template": COMMENT_TEMPLATE_OPTIONS}, "")
            comment_template = st.selectbox("Quick comment helper", template_options, index=0)
            current_week = st.text_area(
                "This week comment",
                editable_text(selected.get("This Week Comment")),
                height=150,
                help="This will be prepended to the existing Addt'l NPL Comment history on save/export.",
            )
            built_comment = build_weekly_comment_text(current_week, comment_template)
            updates["This Week Comment"] = built_comment
            if built_comment:
                preview_text = f"{format_comment_entry_date(comment_date)} - {built_comment}\n{editable_text(selected.get('Comment History'))}".strip()
                st.caption(f"Addt'l NPL Comment will save as: {format_comment_entry_date(comment_date)} - {built_comment}")
            else:
                preview_text = editable_text(selected.get("Comments Preview"))
                st.caption("No current-week comment entered yet.")
            st.text_area(
                "Addt'l NPL Comment preview",
                preview_text,
                disabled=True,
                height=170,
            )

        st.markdown("<div class='action-bar'>", unsafe_allow_html=True)
        action_left, action_mid, action_right = st.columns(3)
        with action_left:
            submitted = st.form_submit_button("Apply deal edits", width="stretch")
        with action_mid:
            submitted_next = st.form_submit_button("Apply edits + next deal", width="stretch")
        with action_right:
            submitted_next_blank = st.form_submit_button("Apply edits + next blank comment", width="stretch")
        st.markdown("</div>", unsafe_allow_html=True)

        if submitted or submitted_next or submitted_next_blank:
            if submitted_next:
                updates["_advance_mode"] = "next"
            elif submitted_next_blank:
                updates["_advance_mode"] = "next_blank"
            else:
                updates["_advance_mode"] = "stay"
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
    am_summary = am_summary.drop(columns=["Total_UPB"]).rename(
        columns={
            "Blank_Current_Comments": "Blank current comments",
            "Needs_Discussion": "Needs discussion",
        }
    )
    st.markdown("**AM workload**")
    st.dataframe(am_summary, width="stretch", hide_index=True)

    st.markdown("**Current queue**")
    queue_cols = [
        "Deal Number",
        "Deal Name",
        "Asset Manager",
        "Current DQ Status",
        "Current UPB",
        "Resolution Likelihood",
        "Expected Final Resolution",
        "Previous Weekly Comment",
        "This Week Comment",
        "Flag Summary",
    ]
    show_cols = [col for col in queue_cols if col in filtered_df.columns]
    st.dataframe(filtered_df[show_cols], width="stretch", hide_index=True)


def render_bulk_edit(filtered_df: pd.DataFrame, picklists: dict[str, list[str]]) -> pd.DataFrame | None:
    st.subheader("Bulk update grid")
    st.caption("Use this for fast queue clean-up. The prior comment is read-only; current-week comment is what gets prepended on export.")

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
        "Needs Discussion",
    ]
    available_bulk = [col for col in default_cols if col in filtered_df.columns]
    selected_cols = st.multiselect("Columns shown in grid", available_bulk, default=available_bulk)

    grid_columns = [col for col in [*GRID_CONTEXT_HEADERS, *selected_cols] if col in filtered_df.columns]
    editor_view = filtered_df[["_excel_row", *grid_columns]].copy()

    disabled_columns = [
        "_excel_row",
        *GRID_CONTEXT_HEADERS,
        "Previous Weekly Comment",
    ]

    column_config: dict[str, Any] = {
        "Needs Discussion": st.column_config.CheckboxColumn("Needs Discussion"),
    }
    for field, options in picklists.items():
        if field in editor_view.columns and field not in {"Comment Template"}:
            column_config[field] = st.column_config.SelectboxColumn(field, options=options)
    if "This Week Comment" in editor_view.columns:
        column_config["This Week Comment"] = st.column_config.TextColumn("This Week Comment", width="large")
    if "Previous Weekly Comment" in editor_view.columns:
        column_config["Previous Weekly Comment"] = st.column_config.TextColumn("Previous Weekly Comment", width="large")

    edited = st.data_editor(
        editor_view,
        width="stretch",
        hide_index=True,
        disabled=disabled_columns,
        num_rows="fixed",
        column_config=column_config,
        key="bulk_editor_v3",
    )

    if st.button("Apply bulk grid changes", width="stretch"):
        return edited
    return None


def main() -> None:
    st.set_page_config(page_title="NPL Deal Review", layout="wide", initial_sidebar_state="collapsed")
    apply_base_css()

    page_left, page_right = st.columns([1.1, 4.9], gap="large")

    with page_left:
        st.markdown('<div class="side-panel">', unsafe_allow_html=True)
        file_bytes, workbook_name = get_active_workbook(st.container())
        st.markdown("</div>", unsafe_allow_html=True)

    if file_bytes is None:
        with page_right:
            st.title("NPL deal review")
            st.caption("Upload the latest workbook from the left control panel to begin.")
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

    visible_count = int((~display_df["_workbook_hidden"]).sum())
    hidden_count = int(display_df["_workbook_hidden"].sum())

    with page_left:
        st.markdown('<div class="side-panel">', unsafe_allow_html=True)
        sorted_df, comment_date, queue_view, workspace_mode, selected_sort = apply_filters(display_df, st.container())
        st.markdown("</div>", unsafe_allow_html=True)

    st.session_state["editor_df"] = refresh_editor_derived_fields(st.session_state["editor_df"], comment_date)

    with page_right:
        if workspace_mode != "Review workspace":
            st.title("NPL deal review")
            st.caption("Compact review workspace for live Teams screenshare, AM updates, and export back into the original workbook layout.")
            with st.expander("Workbook details", expanded=False):
                st.caption(
                    f"Deal sheet = {parsed['deal_sheet']}; property sheet = {parsed['property_sheet'] or 'not found'}; "
                    f"total deals = {len(display_df)}; workbook-visible rows = {visible_count}; workbook-hidden rows = {hidden_count}."
                )

        flash_message = st.session_state.pop("flash_message", None)
        if flash_message:
            st.success(flash_message)

        if sorted_df.empty:
            st.warning("No deals match the current filters.")
            return

        if workspace_mode == "Review workspace":
            order_label = "workbook file order" if selected_sort == "Workbook file order" else selected_sort.lower()
            st.caption(f"Current review queue: {order_label} · {len(sorted_df):,} deal(s) in scope")

        render_queue_navigation(sorted_df)

        selected = sorted_df[sorted_df["_excel_row"] == st.session_state["selected_row"]].iloc[0]
        queue_position = sorted_df.index[sorted_df["_excel_row"] == st.session_state["selected_row"]][0] + 1
        queue_total = len(sorted_df)

        if workspace_mode == "Overview":
            render_overview(sorted_df, parsed.get("snapshot_label"))
        elif workspace_mode == "Bulk update":
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
            render_deal_header(selected, queue_position, queue_total)
            updates = render_review_form(selected, picklists, parsed["properties_df"], comment_date)
            if updates:
                advance_mode = updates.pop("_advance_mode", "stay")
                st.session_state["editor_df"] = apply_detail_edit(
                    st.session_state["editor_df"],
                    int(selected["_excel_row"]),
                    updates,
                    comment_date,
                )
                option_rows = sorted_df["_excel_row"].tolist()
                current_idx = option_rows.index(int(selected["_excel_row"]))
                next_row = None

                if advance_mode == "next" and current_idx < len(option_rows) - 1:
                    next_row = option_rows[current_idx + 1]
                elif advance_mode == "next_blank":
                    refreshed_after_save = build_display_view(
                        base_deals=parsed["deals_df"],
                        editor_df=st.session_state["editor_df"],
                        property_summary=parsed["property_summary"],
                        entry_date=comment_date,
                        picklists=picklists,
                    )
                    queue_after_save = sort_deals(
                        refreshed_after_save[refreshed_after_save["_excel_row"].isin(option_rows)],
                        selected_sort,
                    )
                    blanks = queue_after_save[
                        queue_after_save["This Week Comment"].fillna("").astype(str).str.strip().eq("")
                    ]["_excel_row"].tolist()
                    for row_num in blanks:
                        if row_num != int(selected["_excel_row"]):
                            next_row = row_num
                            break
                    if next_row is None and current_idx < len(option_rows) - 1:
                        next_row = option_rows[current_idx + 1]

                if next_row is not None:
                    st.session_state["selected_row"] = next_row
                st.session_state["flash_message"] = "Deal edits saved."
                st.rerun()

        changes = workbook_change_set(
            st.session_state["editor_df"],
            parsed["original_editor_values"],
            comment_date,
        )
        changed_deals = len({row for row, _ in changes.keys()})

        st.divider()
        summary_col, action_col = st.columns([2, 1])
        with summary_col:
            st.write(f"Pending edits in the app view: **{len(changes)} cells across {changed_deals} deal(s)**")
            summary_parts = []
            if queue_view != "All deals":
                summary_parts.append(f"queue focus: {queue_view}")
            summary_parts.append(f"order: {selected_sort.lower()}")
            st.caption(" · ".join(summary_parts))
        with action_col:
            if st.button("Reset all app edits", width="stretch"):
                st.session_state["editor_df"] = make_editor_df(parsed["deals_df"])
                st.rerun()

        updated_bytes = (
            build_updated_workbook(
                file_bytes=file_bytes,
                sheet_path=parsed["deal_sheet_path"],
                changes=changes,
                cell_meta=parsed["cell_meta"],
                sheet_name=parsed["deal_sheet"],
            )
            if changes
            else file_bytes
        )

        quarter_label = infer_quarter_label(workbook_name, parsed.get("snapshot_label"))
        download_name = f"{quarter_label} NPL Resolutions.xlsx"
        st.download_button(
            "Download updated workbook",
            data=updated_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )


if __name__ == "__main__":
    main()
