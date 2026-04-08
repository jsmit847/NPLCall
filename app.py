from __future__ import annotations

import hashlib
from datetime import date
from typing import Any

import pandas as pd
import streamlit as st

from workbook_utils import (
    COMMENT_TEMPLATE_OPTIONS,
    CORE_FIELDS,
    EXPORTABLE_FIELDS,
    GRID_CONTEXT_HEADERS,
    LONG_TEXT_FIELDS,
    MEETING_STATUS_OPTIONS,
    READONLY_CONTEXT_FIELDS,
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
    normalize_number,
    parse_deal_workbook,
    refresh_editor_derived_fields,
    row_quality_flags,
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
            .block-container {padding-top: 0.8rem; padding-bottom: 1.0rem; max-width: 100%;}
            h1 {font-size: 1.35rem !important; margin-bottom: 0.1rem !important;}
            h2 {font-size: 1.05rem !important; margin-top: 0.35rem !important;}
            h3 {font-size: 0.95rem !important; margin-top: 0.25rem !important;}
            div[data-testid="stMetricValue"] {font-size: 1.0rem !important;}
            div[data-testid="stMetricLabel"] {font-size: 0.78rem !important;}
            .deal-header {
                border: 1px solid #d7dbe2;
                border-radius: 10px;
                padding: 0.55rem 0.7rem;
                margin-bottom: 0.55rem;
                background: #fbfcfe;
            }
            .deal-title {
                font-size: 1.02rem;
                font-weight: 700;
                line-height: 1.2;
                margin-right: 0.35rem;
            }
            .am-badge {
                display: inline-block;
                padding: 0.16rem 0.45rem;
                border-radius: 999px;
                border: 1px solid #c9daf8;
                background: #eef4ff;
                font-size: 0.78rem;
                font-weight: 700;
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
                padding: 0.65rem 0.8rem;
                background: white;
            }
            .presenter-section-title {
                font-size: 0.86rem;
                font-weight: 700;
                margin: 0.1rem 0 0.35rem 0;
            }
            .presenter-text {
                font-size: 0.96rem;
                line-height: 1.35;
                white-space: pre-wrap;
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
            .stTabs [data-baseweb="tab-list"] {gap: 0.25rem;}
            .stTabs [data-baseweb="tab"] {padding-top: 0.35rem; padding-bottom: 0.35rem;}
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



def apply_filters(display_df: pd.DataFrame) -> tuple[pd.DataFrame, date, str]:
    with st.sidebar:
        st.header("Review filters")
        comment_date = st.date_input("Current week comment date", value=st.session_state.get("comment_date", date.today()))
        st.session_state["comment_date"] = comment_date
        show_hidden = st.checkbox("Include workbook-hidden rows", value=False)

        asset_managers = sorted([x for x in display_df["Asset Manager"].dropna().astype(str).unique().tolist() if x])
        selected_ams = st.multiselect("Asset Manager", asset_managers)
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
            [
                "All deals",
                "Open only",
                "Missing current-week comment",
                "Needs discussion",
                "Has data flags",
            ],
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

    sort_lookup = {
        "Asset Manager / Deal": "Asset Manager / Deal",
        "UPB desc": "UPB desc",
        "Flag count desc": "Flag count desc",
        "Last comment date desc": "Last comment date desc",
    }
    sorted_df = sort_deals(filtered_df, sort_lookup[sort_mode]).reset_index(drop=True)
    return sorted_df, comment_date, queue_view



def render_queue_navigation(sorted_df: pd.DataFrame) -> None:
    option_rows = sorted_df["_excel_row"].tolist()
    if not option_rows:
        return

    if st.session_state.get("selected_row") not in option_rows:
        st.session_state["selected_row"] = option_rows[0]

    current_idx = option_rows.index(st.session_state["selected_row"])

    def next_index_with(predicate) -> int | None:
        for idx in range(current_idx + 1, len(option_rows)):
            row = sorted_df.iloc[idx]
            if predicate(row):
                return idx
        return None

    nav1, nav2, nav3, nav4, nav5 = st.columns([1, 1, 1, 1.2, 4])
    with nav1:
        if st.button("Previous", disabled=current_idx == 0, use_container_width=True):
            st.session_state["selected_row"] = option_rows[current_idx - 1]
            st.rerun()
    with nav2:
        if st.button("Next", disabled=current_idx == len(option_rows) - 1, use_container_width=True):
            st.session_state["selected_row"] = option_rows[current_idx + 1]
            st.rerun()
    with nav3:
        next_open = next_index_with(lambda row: row.get("Review Status", "Open") == "Open")
        if st.button("Next open", disabled=next_open is None, use_container_width=True):
            st.session_state["selected_row"] = option_rows[next_open]
            st.rerun()
    with nav4:
        next_comment = next_index_with(lambda row: safe_text(row.get("This Week Comment")) == "-")
        if st.button("Next blank comment", disabled=next_comment is None, use_container_width=True):
            st.session_state["selected_row"] = option_rows[next_comment]
            st.rerun()
    with nav5:
        selected_row = st.selectbox(
            "Deal queue",
            option_rows,
            format_func=lambda row_num: format_deal_label(sorted_df[sorted_df["_excel_row"] == row_num].iloc[0]),
            index=option_rows.index(st.session_state["selected_row"]),
        )
        st.session_state["selected_row"] = selected_row



def render_deal_header(selected: pd.Series, queue_position: int, queue_total: int) -> None:
    subtitle = (
        f"Deal #{safe_text(selected.get('Deal Number'))} · {safe_text(selected.get('Location'))} · "
        f"{safe_text(selected.get('Bridge / Term'))} · Queue {queue_position} of {queue_total}"
    )
    st.markdown(
        f"""
        <div class="deal-header">
            <div><span class="deal-title">{safe_text(selected.get('Deal Name'))}</span>
            <span class="am-badge">AM: {safe_text(selected.get('Asset Manager'))}</span></div>
            <div class="deal-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def render_metric_strip(selected: pd.Series) -> None:
    cols = st.columns(8)
    cols[0].metric("UPB", metric_currency(selected.get("Current UPB")))
    cols[1].metric("As-Is", metric_currency(selected.get("Salesforce As-Is Valuation")))
    cols[2].metric("As-Is LTV", metric_pct(selected.get("Salesforce Implied As-Is LTV")))
    cols[3].metric("ARV", metric_currency(selected.get("Salesforce ARV")))
    cols[4].metric("ARV LTV", metric_pct(selected.get("Salesforce ARV LTV")))
    cols[5].metric("DQ", safe_text(selected.get("Current DQ Status")))
    cols[6].metric("6/30 NPL", safe_text(selected.get("6/30 NPL")))
    cols[7].metric("Last comment", metric_date(selected.get("Last Comment Date")))



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
                    "Days Past Due",
                ]
                if col in related.columns
            ]
            with st.expander("Property detail", expanded=False):
                st.dataframe(related[show_cols], use_container_width=True, hide_index=True)



def editable_field_input(field: str, current_value: str, picklists: dict[str, list[str]]) -> Any:
    if field in picklists:
        options = field_options(field, picklists, current_value)
        return st.selectbox(field, options, index=selected_index(options, current_value))
    if field in LONG_TEXT_FIELDS or len(current_value) > 80:
        return st.text_area(field, current_value, height=90)
    return st.text_input(field, current_value)



def render_review_form(
    selected: pd.Series,
    picklists: dict[str, list[str]],
    properties_df: pd.DataFrame,
    comment_date: date,
) -> dict[str, Any] | None:
    render_metric_strip(selected)
    with st.form(f"detail_form_{int(selected['_excel_row'])}"):
        left, middle, right = st.columns([1.0, 1.2, 1.35])
        updates: dict[str, Any] = {}

        with left:
            st.subheader("Context")
            render_context_column(selected)
            review_status_options = field_options("Review Status", {"Review Status": MEETING_STATUS_OPTIONS}, str(selected.get("Review Status", "Open")))
            updates["Review Status"] = st.selectbox(
                "Meeting status",
                review_status_options,
                index=selected_index(review_status_options, str(selected.get("Review Status", "Open"))),
            )
            updates["Needs Discussion"] = st.checkbox(
                "Needs discussion",
                value=bool(selected.get("Needs Discussion", False)),
            )
            render_property_summary(selected, properties_df)

        with middle:
            st.subheader("Manual updates")
            for section_title, fields in SECTION_MAP.items():
                present_fields = [field for field in fields if field in selected.index]
                if not present_fields:
                    continue
                st.markdown(f"**{section_title}**")
                for field in present_fields:
                    updates[field] = editable_field_input(field, editable_text(selected.get(field)), picklists)

        with right:
            st.subheader("Comments")
            st.text_area(
                "Previous weekly comment",
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
                help="This will be prepended to the existing Comments history on export.",
            )
            built_comment = build_weekly_comment_text(current_week, comment_template)
            updates["This Week Comment"] = built_comment
            st.caption(f"Export preview will prepend: {format_comment_entry_date(comment_date)} - {built_comment}" if built_comment else "No current-week comment entered yet.")
            st.text_area(
                "Comments export preview",
                editable_text(selected.get("Comments Preview")) if not built_comment else f"{format_comment_entry_date(comment_date)} - {built_comment}\n{editable_text(selected.get('Comment History'))}".strip(),
                disabled=True,
                height=170,
            )

        submitted = st.form_submit_button("Apply deal edits", use_container_width=True)
        if submitted:
            return updates
    return None



def render_overview(filtered_df: pd.DataFrame, snapshot_label: str | None, picklists: dict[str, list[str]]) -> None:
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
            Open=("Review Status", lambda s: (s.astype(str) == "Open").sum()),
        )
        .reset_index()
        .rename(columns={"Asset Manager": "Asset Manager"})
    )
    am_summary["Total UPB"] = am_summary["Total_UPB"].apply(metric_currency)
    am_summary = am_summary.drop(columns=["Total_UPB"]).rename(columns={
        "Blank_Current_Comments": "Blank current comments",
        "Needs_Discussion": "Needs discussion",
    })

    st.markdown("**AM workload**")
    st.dataframe(am_summary, use_container_width=True, hide_index=True)

    qos_findings = []
    valuation_values = set(filtered_df.get("Valuation Type", pd.Series(dtype=str)).dropna().astype(str))
    if "Drive By BPO" in valuation_values and "Drive-By BPO" in valuation_values:
        qos_findings.append("Valuation Type has inconsistent variants for Drive-By BPO; a controlled picklist helps keep exports uniform.")
    fcl_values = set(filtered_df.get("FCL Needed Prior to Final Resolution (Y/N)", pd.Series(dtype=str)).dropna().astype(str))
    if len(fcl_values) > 4:
        qos_findings.append("The FCL field is carrying mixed meanings (Y/N plus Foreclosure / Payoff / 0). The app now pushes this toward a standard dropdown.")
    offer_values = filtered_df.get("Third Party Offer Amount", pd.Series(dtype=str)).dropna().astype(str)
    if any(v.strip().isdigit() for v in offer_values) and any(v.strip().lower() in {"no", "yes", "n/a", "na"} for v in offer_values):
        qos_findings.append("Third Party Offer Amount currently mixes numbers and yes/no text. Keeping offer note separate reduces rework during review.")
    if qos_findings:
        st.info("\n\n".join(qos_findings[:3]))

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
        "Review Status",
        "Flag Summary",
    ]
    show_cols = [col for col in queue_cols if col in filtered_df.columns]
    st.dataframe(filtered_df[show_cols], use_container_width=True, hide_index=True)



def render_bulk_edit(filtered_df: pd.DataFrame, picklists: dict[str, list[str]], comment_date: date) -> pd.DataFrame | None:
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
        "Review Status",
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
        "Review Status": st.column_config.SelectboxColumn("Review Status", options=MEETING_STATUS_OPTIONS),
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
        use_container_width=True,
        hide_index=True,
        disabled=disabled_columns,
        num_rows="fixed",
        column_config=column_config,
        key="bulk_editor_v3",
    )

    if st.button("Apply bulk grid changes", use_container_width=True):
        return edited
    return None



def render_presentation_mode(selected: pd.Series, queue_position: int, queue_total: int, properties_df: pd.DataFrame) -> None:
    st.subheader("Presentation mode")
    st.caption("Use this tab while screen-sharing in Teams. It is intentionally denser and cleaner than the edit view.")
    render_deal_header(selected, queue_position, queue_total)

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
            chip_text.append(f"<span class='status-chip'>{field.replace('Current ', '').replace('Expected ', '').replace('Timing', 'Tm')}: {value}</span>")
    if chip_text:
        st.markdown("".join(chip_text), unsafe_allow_html=True)

    left, right = st.columns([1.8, 1.1])
    with left:
        st.markdown("<div class='presenter-panel'>", unsafe_allow_html=True)
        st.markdown("<div class='presenter-section-title'>Where the deal stands</div>", unsafe_allow_html=True)
        standup_lines = [
            f"Asset Manager: {safe_text(selected.get('Asset Manager'))}",
            f"Location / Type: {safe_text(selected.get('Location'))} · {safe_text(selected.get('Type'))}",
            f"Current UPB / As-Is / ARV: {metric_currency(selected.get('Current UPB'))} · {metric_currency(selected.get('Salesforce As-Is Valuation'))} · {metric_currency(selected.get('Salesforce ARV'))}",
            f"Maturity / Next payment: {metric_date(selected.get('Maturity Date'))} · {metric_date(selected.get('Next Payment Date'))}",
            f"Resolution path: {safe_text(selected.get('Expected Resolution Type'))} / {safe_text(selected.get('Expected Final Resolution'))}",
            f"Additional NPL comment: {safe_text(selected.get("Addt'l NPL Comment"))}",
        ]
        st.markdown(f"<div class='presenter-text'>{'<br>'.join(standup_lines)}</div>", unsafe_allow_html=True)
        st.markdown("<div class='presenter-section-title'>Latest prior comment</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='presenter-text'>{safe_text(selected.get('Previous Weekly Comment'))}</div>", unsafe_allow_html=True)
        if safe_text(selected.get("This Week Comment")) != "-":
            st.markdown("<div class='presenter-section-title'>Current week draft</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='presenter-text'>{safe_text(selected.get('This Week Comment'))}</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown("<div class='presenter-panel'>", unsafe_allow_html=True)
        st.markdown("<div class='presenter-section-title'>Valuation & offer snapshot</div>", unsafe_allow_html=True)
        valuation_lines = [
            f"Valuation Type: {safe_text(selected.get('Valuation Type'))}",
            f"As-Is Value: {metric_currency(selected.get('As Is Appraised Value'))}",
            f"As Stabilized: {metric_currency(selected.get('As Stabilized Value'))}",
            f"Valuation Date: {metric_date(selected.get('Date of Valuation'))}",
            f"Finalized?: {safe_text(selected.get('Is Appraisal Finalized? (Y/N)'))}",
            f"Third Party Offer: {safe_text(selected.get('Third Party Offer Amount'))}",
            f"Offer Note: {safe_text(selected.get('Third Party Offer Note'))}",
            f"Meeting status: {safe_text(selected.get('Review Status'))}",
        ]
        st.markdown(f"<div class='presenter-text'>{'<br>'.join(valuation_lines)}</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        render_property_summary(selected, properties_df)



def main() -> None:
    st.set_page_config(page_title="NPL Deal Review", layout="wide", initial_sidebar_state="collapsed")
    apply_base_css()

    st.title("NPL deal review")
    st.caption(
        "Upload the latest workbook first. The app then opens a compact deal queue, shows Asset Manager directly next to the deal, lets AMs enter this week's updates, and exports back into the original workbook layout."
    )

    uploaded_file = st.file_uploader("Step 1 — upload the latest workbook", type=["xlsx"])
    if uploaded_file is None:
        st.info("After upload, the review queue, presentation mode, bulk editor, and workbook download will appear here.")
        return

    file_bytes = uploaded_file.getvalue()
    file_hash = hashlib.md5(file_bytes).hexdigest()
    parsed = cached_parse_workbook(file_bytes)
    picklists = cached_picklists(parsed["deals_df"])

    if st.session_state.get("file_hash") != file_hash:
        init_state(file_hash, parsed)

    comment_date = st.session_state.get("comment_date", date.today())
    refreshed_editor = refresh_editor_derived_fields(st.session_state["editor_df"], comment_date)
    st.session_state["editor_df"] = refreshed_editor

    display_df = build_display_view(
        base_deals=parsed["deals_df"],
        editor_df=st.session_state["editor_df"],
        property_summary=parsed["property_summary"],
        entry_date=comment_date,
        picklists=picklists,
    )

    visible_count = int((~display_df["_workbook_hidden"]).sum())
    hidden_count = int(display_df["_workbook_hidden"].sum())
    st.caption(
        f"Workbook inspection: deal sheet = {parsed['deal_sheet']}; property sheet = {parsed['property_sheet'] or 'not found'}; total deals = {len(display_df)}; workbook-visible rows = {visible_count}; workbook-hidden rows = {hidden_count}."
    )

    sorted_df, comment_date, queue_view = apply_filters(display_df)
    st.session_state["editor_df"] = refresh_editor_derived_fields(st.session_state["editor_df"], comment_date)

    if sorted_df.empty:
        st.warning("No deals match the current filters.")
        return

    render_queue_navigation(sorted_df)

    selected = sorted_df[sorted_df["_excel_row"] == st.session_state["selected_row"]].iloc[0]
    queue_position = sorted_df.index[sorted_df["_excel_row"] == st.session_state["selected_row"]][0] + 1
    queue_total = len(sorted_df)

    tab_overview, tab_review, tab_bulk, tab_present = st.tabs(["Overview", "Review", "Bulk update", "Presentation mode"])

    with tab_overview:
        render_overview(sorted_df, parsed.get("snapshot_label"), picklists)

    with tab_review:
        render_deal_header(selected, queue_position, queue_total)
        updates = render_review_form(selected, picklists, parsed["properties_df"], comment_date)
        if updates:
            st.session_state["editor_df"] = apply_detail_edit(
                st.session_state["editor_df"],
                int(selected["_excel_row"]),
                updates,
                comment_date,
            )
            st.rerun()

    with tab_bulk:
        edited_subset = render_bulk_edit(sorted_df, picklists, comment_date)
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

    with tab_present:
        render_presentation_mode(selected, queue_position, queue_total, parsed["properties_df"])

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
        if queue_view != "All deals":
            st.caption(f"Queue focus currently applied: {queue_view}")
    with action_col:
        if st.button("Reset all app edits", use_container_width=True):
            st.session_state["editor_df"] = make_editor_df(parsed["deals_df"])
            st.rerun()

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

    base_name = uploaded_file.name.rsplit(".", 1)[0]
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
