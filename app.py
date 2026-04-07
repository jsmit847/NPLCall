from __future__ import annotations

import hashlib

import streamlit as st

from workbook_utils import (
    EDITABLE_HEADERS,
    GRID_CONTEXT_HEADERS,
    LONG_TEXT_HEADERS,
    READONLY_CONTEXT_HEADERS,
    SECTION_MAP,
    apply_detail_edit,
    build_filtered_view,
    build_updated_workbook,
    editable_text,
    format_deal_label,
    make_editor_df,
    metric_currency,
    metric_date,
    metric_pct,
    parse_deal_workbook,
    workbook_change_set,
)


@st.cache_data(show_spinner=False)
def cached_parse_workbook(file_bytes: bytes):
    return parse_deal_workbook(file_bytes)



def render_overview(filtered_df):
    st.subheader("Workbook summary")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Deals in current app view", f"{len(filtered_df):,}")
    col2.metric("Total 3/27 UPB", metric_currency(filtered_df["3/27 UPB"].fillna(0).sum()))
    col3.metric("Unique Asset Managers", f"{filtered_df['Asset Manager'].nunique():,}")
    npl_count = filtered_df[filtered_df["6/30 NPL"].astype(str).str.upper().eq("Y")].shape[0]
    col4.metric("6/30 NPL = Y", f"{npl_count:,}")

    st.dataframe(
        filtered_df[
            [
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
                "Comments",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )



def render_deal_detail(selected, properties_df):
    st.subheader(f"{selected['Deal Name']}")
    st.caption(f"Deal {selected['Deal Number']} | Asset Manager: {selected['Asset Manager']}")

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("UPB", metric_currency(selected.get("3/27 UPB")))
    m2.metric("As-Is Value", metric_currency(selected.get("Salesforce As-Is Valuation")))
    m3.metric("As-Is LTV", metric_pct(selected.get("Salesforce Implied As-Is LTV")))
    m4.metric("ARV", metric_currency(selected.get("Salesforce ARV")))
    m5.metric("ARV LTV", metric_pct(selected.get("Salesforce ARV LTV")))
    m6.metric("6/30 NPL", editable_text(selected.get("6/30 NPL")) or "-")

    st.markdown("### Core datapoints")
    left, right = st.columns(2)
    with left:
        for header in [
            "Location",
            "Bridge / Term",
            "Type",
            "Segment",
            "Financing",
            "Loan Commitment",
            "Remaining Commitment",
            "Number Of Units",
        ]:
            st.write(f"**{header}:** {editable_text(selected.get(header)) or '-'}")
    with right:
        for header in [
            "Maturity Date",
            "Next Payment Date",
            "12/31 NPL",
            "3/31 NPL",
            "6/30 NPL",
            "3/27 DQ Status",
            "MBA Loan List Convention",
        ]:
            value = selected.get(header)
            display = metric_date(value) if "Date" in header else (editable_text(value) or "-")
            st.write(f"**{header}:** {display}")

    st.markdown("### Manual update editor")
    with st.form(f"detail_form_{int(selected['_excel_row'])}"):
        updated_values = {}
        for section_title, headers in SECTION_MAP.items():
            st.markdown(f"#### {section_title}")
            for header in headers:
                current_value = editable_text(selected.get(header))
                if header == "3/27 Salesforce AM Comment":
                    st.text_area(header, current_value, disabled=True, height=160)
                    continue
                if header in LONG_TEXT_HEADERS:
                    updated_values[header] = st.text_area(header, current_value, height=140)
                else:
                    updated_values[header] = st.text_input(header, current_value)
        saved = st.form_submit_button("Save deal edits")
        if saved:
            st.success("Deal edits saved in the app view.")
            return updated_values

    if not properties_df.empty and "Deal Number" in properties_df.columns:
        related = properties_df[properties_df["Deal Number"] == selected["Deal Number"]].copy()
        if not related.empty:
            st.markdown("### Related property detail")
            property_columns = [
                c for c in [
                    "Servicer ID",
                    "Servicer",
                    "SF Yardi ID",
                    "Asset ID",
                    "Address",
                    "City",
                    "State",
                    "# Units",
                    "3/27 UPB",
                    "Salesforce As-Is Valuation",
                    "Salesforce ARV",
                    "Days Past Due",
                ] if c in related.columns
            ]
            st.dataframe(related[property_columns], use_container_width=True, hide_index=True)

    return None



def init_state(file_hash, parsed):
    st.session_state["file_hash"] = file_hash
    st.session_state["editor_df"] = make_editor_df(parsed["deals_df"])



def main():
    st.set_page_config(page_title="Deal KPI Review & AM Update App", layout="wide")
    st.title("Deal KPI Review & AM Update App")
    st.write(
        "Upload the current workbook first. The app will then present each deal, show the Asset Manager, allow AM manual updates, and let you download an updated workbook that keeps the original Excel layout."
    )

    uploaded_file = st.file_uploader("Step 1 - Upload the current workbook", type=["xlsx"])
    if uploaded_file is None:
        st.info("Once a workbook is uploaded, the deal presentation and editing workflow will appear here.")
        return

    file_bytes = uploaded_file.getvalue()
    file_hash = hashlib.md5(file_bytes).hexdigest()
    parsed = cached_parse_workbook(file_bytes)

    if st.session_state.get("file_hash") != file_hash:
        init_state(file_hash, parsed)

    base_deals = parsed["deals_df"]
    properties_df = parsed["properties_df"]
    editor_df = st.session_state["editor_df"]
    display_df = build_filtered_view(base_deals, editor_df)

    visible_count = int((~display_df["_workbook_hidden"]).sum())
    hidden_count = int(display_df["_workbook_hidden"].sum())
    st.caption(
        f"Workbook inspection: main deal sheet = {parsed['deal_sheet']}; property detail sheet = {parsed['property_sheet'] or 'not found'}; total deals = {len(display_df)}; workbook-hidden rows = {hidden_count}; workbook-visible rows = {visible_count}."
    )

    with st.sidebar:
        st.header("Filters")
        show_hidden = st.checkbox("Include workbook-hidden rows", value=False)
        asset_managers = sorted([x for x in display_df["Asset Manager"].dropna().astype(str).unique().tolist() if x])
        selected_ams = st.multiselect("Asset Manager", asset_managers)
        selected_bt = st.multiselect("Bridge / Term", sorted(display_df["Bridge / Term"].dropna().astype(str).unique().tolist()))
        selected_segment = st.multiselect("Segment", sorted(display_df["Segment"].dropna().astype(str).unique().tolist()))
        search = st.text_input("Search deal number or name")

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
        search_lower = search.lower()
        filtered_df = filtered_df[
            filtered_df["Deal Number"].astype(str).str.lower().str.contains(search_lower)
            | filtered_df["Deal Name"].astype(str).str.lower().str.contains(search_lower)
        ]

    if filtered_df.empty:
        st.warning("No deals match the current filters.")
        return

    tab_overview, tab_detail, tab_bulk = st.tabs(["Overview", "Deal detail", "Bulk edit"])

    with tab_overview:
        render_overview(filtered_df)

    with tab_detail:
        options = filtered_df.sort_values(["Asset Manager", "Deal Name", "Deal Number"]).reset_index(drop=True)
        option_rows = options["_excel_row"].tolist()
        default_row = option_rows[0]
        if st.session_state.get("selected_row") not in option_rows:
            st.session_state["selected_row"] = default_row

        nav_left, nav_mid, nav_right = st.columns([1, 4, 1])
        current_idx = option_rows.index(st.session_state["selected_row"])

        with nav_left:
            if st.button("Previous", disabled=current_idx == 0):
                st.session_state["selected_row"] = option_rows[current_idx - 1]
                st.rerun()
        with nav_right:
            if st.button("Next", disabled=current_idx == len(option_rows) - 1):
                st.session_state["selected_row"] = option_rows[current_idx + 1]
                st.rerun()
        with nav_mid:
            selected_row = st.selectbox(
                "Select a deal",
                option_rows,
                format_func=lambda row_num: format_deal_label(options[options["_excel_row"] == row_num].iloc[0]),
                index=option_rows.index(st.session_state["selected_row"]),
            )
            st.session_state["selected_row"] = selected_row

        selected = options[options["_excel_row"] == st.session_state["selected_row"]].iloc[0]
        detail_updates = render_deal_detail(selected, properties_df)
        if detail_updates:
            st.session_state["editor_df"] = apply_detail_edit(st.session_state["editor_df"], int(selected["_excel_row"]), detail_updates)
            st.rerun()

    with tab_bulk:
        st.write("Edit the AM update columns across the currently filtered deals. Salesforce AM Comment remains read-only for context.")
        bulk_subset = st.session_state["editor_df"][st.session_state["editor_df"]["_excel_row"].isin(filtered_df["_excel_row"])].copy()
        bulk_subset = bulk_subset.sort_values(["Asset Manager", "Deal Name", "Deal Number"]).reset_index(drop=True)
        disabled_columns = ["_excel_row", *GRID_CONTEXT_HEADERS, *READONLY_CONTEXT_HEADERS]
        edited_subset = st.data_editor(
            bulk_subset,
            use_container_width=True,
            hide_index=True,
            disabled=disabled_columns,
            num_rows="fixed",
            key="bulk_editor",
        )
        full_editor = st.session_state["editor_df"].set_index("_excel_row")
        edited_lookup = edited_subset.set_index("_excel_row")
        for row_num in edited_lookup.index:
            for header in EDITABLE_HEADERS:
                full_editor.loc[row_num, header] = edited_lookup.loc[row_num, header]
        st.session_state["editor_df"] = full_editor.reset_index()

    changes = workbook_change_set(st.session_state["editor_df"], parsed["original_editor_values"])
    changed_deals = len({row for row, _ in changes.keys()})
    st.divider()
    left, right = st.columns([2, 1])
    with left:
        st.write(f"Pending edits in app view: **{len(changes)} cells across {changed_deals} deal(s)**")
    with right:
        if st.button("Reset all app edits"):
            st.session_state["editor_df"] = make_editor_df(parsed["deals_df"])
            st.rerun()

    updated_bytes = build_updated_workbook(
        file_bytes=file_bytes,
        sheet_path=parsed["deal_sheet_path"],
        changes=changes,
        cell_meta=parsed["cell_meta"],
    ) if changes else file_bytes

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
