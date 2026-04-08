from __future__ import annotations

import io
import math
import re
from collections import Counter
from datetime import date, datetime
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_SPACE = "http://www.w3.org/XML/1998/namespace"

PREVIOUS_WEEKLY_COMMENT_HEADER = "Previous Weekly Comment"
THIS_WEEK_COMMENT_HEADER = "This Week Comment"

EDITABLE_HEADERS = [
    "Resolution Likelihood",
    "Resolution for 6/30 Performance",
    "Addt'l NPL Comment",
    "Mod Needed Prior to Final Resolution (Y/N)",
    "FCL Needed Prior to Final Resolution (Y/N)",
    "FCL Date",
    "Expected Final Resolution",
    "Resolution Timing",
    "Liquidity Event Timing",
    "Pref Deal (Y/N)",
    "Pref Amount ($)",
    "Liquidity Event ($)",
    "CV Inspected in Last 6 months? (Y/N)",
    "New Valuation in 1Q26 (Y/N)",
    "Valuation Type",
    "As Is Appraised Value",
    "As Stabilized Value",
    "Report Date",
    "Date of Valuation",
    "Is Appraisal Finalized? (Y/N)",
    "Third Party Offer Amount",
    "Third Party Offer Note",
    "Comments",
]

READONLY_CONTEXT_HEADERS = [
    "3/27 Salesforce AM Comment",
]

LONG_TEXT_HEADERS = {
    "Resolution for 6/30 Performance",
    "Addt'l NPL Comment",
    "Third Party Offer Note",
    "3/27 Salesforce AM Comment",
    "Comments",
    PREVIOUS_WEEKLY_COMMENT_HEADER,
    THIS_WEEK_COMMENT_HEADER,
}

DATE_OR_TEXT_HEADERS = {
    "FCL Date",
    "Report Date",
    "Date of Valuation",
}

NUMERIC_OR_TEXT_HEADERS = {
    "Pref Amount ($)",
    "Liquidity Event ($)",
    "As Is Appraised Value",
    "As Stabilized Value",
    "Third Party Offer Amount",
}

GRID_CONTEXT_HEADERS = [
    "Deal Number",
    "Deal Name",
    "Asset Manager",
    "6/30 NPL",
    "3/27 DQ Status",
]

SECTION_MAP = {
    "6/30 NPL Projections": [
        "Resolution Likelihood",
        "Resolution for 6/30 Performance",
        "Addt'l NPL Comment",
    ],
    "AM Manual Updates - Resolutions": [
        "Mod Needed Prior to Final Resolution (Y/N)",
        "FCL Needed Prior to Final Resolution (Y/N)",
        "FCL Date",
        "Expected Final Resolution",
        "Resolution Timing",
        "Liquidity Event Timing",
        "Pref Deal (Y/N)",
        "Pref Amount ($)",
        "Liquidity Event ($)",
    ],
    "AM Manual Updates - Quarterly Valuations": [
        "CV Inspected in Last 6 months? (Y/N)",
        "New Valuation in 1Q26 (Y/N)",
        "Valuation Type",
        "As Is Appraised Value",
        "As Stabilized Value",
        "Report Date",
        "Date of Valuation",
        "Is Appraisal Finalized? (Y/N)",
        "Third Party Offer Amount",
        "Third Party Offer Note",
    ],
}

PICKLIST_BASE_OPTIONS = {
    "Resolution Likelihood": ["", "High", "Medium", "Low", "Resolved"],
    "Mod Needed Prior to Final Resolution (Y/N)": ["", "Y", "N", "N/A"],
    "FCL Needed Prior to Final Resolution (Y/N)": ["", "Y", "N", "Foreclosure", "Payoff", "N/A"],
    "Expected Final Resolution": [
        "",
        "REO Sale",
        "DPO",
        "Payoff",
        "Receivership Sale",
        "Foreclosure",
        "Payoff with Pref",
        "Note Sale",
        "Modification",
        "Mod/FC",
        "Rcvr/FC",
        "TBD",
        "Mod/Paydowns",
        "Title Insurance",
        "JV with new developer",
        "Liquidation of Portfolio",
        "Payoffs/FC",
        "Sale",
        "Sale of Asset",
        "Bankruptcy Sale",
        "Discounted Payoff",
        "Refi",
        "N/A",
    ],
    "Pref Deal (Y/N)": ["", "Y", "N", "N/A"],
    "CV Inspected in Last 6 months? (Y/N)": ["", "Y", "N", "N/A"],
    "New Valuation in 1Q26 (Y/N)": ["", "Y", "N", "N/A"],
    "Valuation Type": [
        "",
        "Appraisal",
        "Drive By BPO",
        "Drive-By BPO",
        "Interior BPO",
        "BOV",
        "Ext-BPO",
        "Exterior Appraisal",
        "BPO/Appraisal",
        "Drive-By BOV",
        "BPO",
        "N/A",
    ],
    "Is Appraisal Finalized? (Y/N)": ["", "Y", "N", "N/A"],
}


HEADER_ALIASES = {
    "CV Inspected in Last 6 months? (Y/N)": "CV Inspected in Last 6 months? (Y/N)",
    "Third Party Offer Note": "Third Party Offer Note",
    "As Is Appraised Value": "As Is Appraised Value",
    "As Stabilized Value": "As Stabilized Value",
}


# ---------------------------------------------------------------------------
# Header / workbook parsing helpers
# ---------------------------------------------------------------------------

def sanitize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text



def normalize_header(value: Any) -> str:
    text = sanitize_header(value)
    return HEADER_ALIASES.get(text, text)



def get_sheet_info(file_bytes: bytes) -> dict[str, dict[str, str]]:
    with ZipFile(io.BytesIO(file_bytes)) as zf:
        workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rel_map = {
        rel.get("Id"): rel.get("Target")
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship")
    }

    sheet_info: dict[str, dict[str, str]] = {}
    sheets_elem = workbook_root.find(f"{{{MAIN_NS}}}sheets")
    if sheets_elem is None:
        return sheet_info

    for sheet in sheets_elem:
        name = sheet.get("name", "")
        rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
        target = rel_map.get(rid, "")
        if target and not target.startswith("xl/"):
            target = f"xl/{target}"
        sheet_info[name] = {
            "state": sheet.get("state", "visible"),
            "rid": rid,
            "path": target,
        }
    return sheet_info



def get_hidden_rows(file_bytes: bytes, sheet_path: str) -> set[int]:
    with ZipFile(io.BytesIO(file_bytes)) as zf:
        root = ET.fromstring(zf.read(sheet_path))
    hidden_rows: set[int] = set()
    sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        return hidden_rows
    for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
        if row.get("hidden") == "1":
            hidden_rows.add(int(row.get("r")))
    return hidden_rows



def detect_sheet_names(file_bytes: bytes) -> dict[str, str | None]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    sheet_info = get_sheet_info(file_bytes)

    deal_candidates: list[tuple[int, str]] = []
    property_candidates: list[tuple[int, str]] = []

    for ws in wb.worksheets:
        state = sheet_info.get(ws.title, {}).get("state", "visible")
        visible_bonus = 100 if state == "visible" else 0

        row2: list[str] = []
        try:
            row2 = [normalize_header(v) for v in next(ws.iter_rows(min_row=2, max_row=2, values_only=True))]
        except StopIteration:
            pass

        row4: list[str] = []
        try:
            row4 = [normalize_header(v) for v in next(ws.iter_rows(min_row=4, max_row=4, values_only=True))]
        except StopIteration:
            pass

        if {"Deal Number", "Asset Manager", "Deal Name"}.issubset(set(row2)):
            score = visible_bonus
            if "Loan List" in ws.title:
                score += 10
            deal_candidates.append((score, ws.title))

        if {"Deal Number", "Servicer ID", "Address"}.issubset(set(row4)):
            score = visible_bonus
            if "Property List" in ws.title:
                score += 10
            property_candidates.append((score, ws.title))

    deal_sheet = max(deal_candidates)[1] if deal_candidates else None
    property_sheet = max(property_candidates)[1] if property_candidates else None
    return {"deal_sheet": deal_sheet, "property_sheet": property_sheet}



def parse_deal_workbook(file_bytes: bytes) -> dict[str, Any]:
    sheet_names = detect_sheet_names(file_bytes)
    deal_sheet = sheet_names["deal_sheet"]
    property_sheet = sheet_names["property_sheet"]
    if not deal_sheet:
        raise ValueError("Could not find the main deal sheet in the uploaded workbook.")

    wb_values = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    wb_formula = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=False)
    sheet_info = get_sheet_info(file_bytes)

    deals_values_ws = wb_values[deal_sheet]
    deals_formula_ws = wb_formula[deal_sheet]
    deal_sheet_path = sheet_info[deal_sheet]["path"]
    hidden_rows = get_hidden_rows(file_bytes, deal_sheet_path)

    raw_headers = [v for v in next(deals_values_ws.iter_rows(min_row=2, max_row=2, values_only=True))]
    clean_headers = [normalize_header(v) for v in raw_headers]
    col_letters = {clean_headers[idx]: get_column_letter(idx + 1) for idx in range(len(clean_headers)) if clean_headers[idx]}

    records: list[dict[str, Any]] = []
    original_editor_values: dict[int, dict[str, str]] = {}
    cell_meta: dict[tuple[int, str], dict[str, Any]] = {}

    value_rows = deals_values_ws.iter_rows(min_row=3, max_row=deals_values_ws.max_row, values_only=False)
    formula_rows = deals_formula_ws.iter_rows(min_row=3, max_row=deals_formula_ws.max_row, values_only=False)

    for value_row, formula_row in zip(value_rows, formula_rows):
        excel_row = value_row[0].row
        if value_row[0].value in (None, ""):
            continue

        record: dict[str, Any] = {
            "_excel_row": excel_row,
            "_workbook_hidden": excel_row in hidden_rows,
        }
        editor_values_for_row: dict[str, str] = {}

        for idx, (value_cell, formula_cell) in enumerate(zip(value_row, formula_row)):
            if idx >= len(clean_headers):
                continue
            header = clean_headers[idx]
            if not header:
                continue
            value = value_cell.value
            record[header] = value
            if header in EDITABLE_HEADERS or header in READONLY_CONTEXT_HEADERS:
                editor_values_for_row[header] = editable_text(value)
                cell_meta[(excel_row, header)] = {
                    "coord": formula_cell.coordinate,
                    "data_type": formula_cell.data_type,
                    "number_format": formula_cell.number_format,
                    "original_formula": formula_cell.value if formula_cell.data_type == "f" else None,
                    "raw_value": formula_cell.value,
                    "column_letter": col_letters[header],
                }

        records.append(record)
        original_editor_values[excel_row] = editor_values_for_row

    deals_df = pd.DataFrame(records)
    if deals_df.empty:
        raise ValueError("The main deal sheet was found, but no deal rows could be read.")

    properties_df = pd.DataFrame()
    if property_sheet:
        properties_ws = wb_values[property_sheet]
        prop_headers = [normalize_header(v) for v in next(properties_ws.iter_rows(min_row=4, max_row=4, values_only=True))]
        property_records: list[dict[str, Any]] = []
        for row in properties_ws.iter_rows(min_row=5, max_row=properties_ws.max_row, values_only=True):
            if row[0] in (None, ""):
                continue
            property_records.append({
                prop_headers[i]: row[i] for i in range(min(len(prop_headers), len(row))) if prop_headers[i]
            })
        if property_records:
            properties_df = pd.DataFrame(property_records)

    return {
        "deal_sheet": deal_sheet,
        "property_sheet": property_sheet,
        "deal_sheet_path": deal_sheet_path,
        "deals_df": deals_df,
        "properties_df": properties_df,
        "cell_meta": cell_meta,
        "original_editor_values": original_editor_values,
        "sheet_info": sheet_info,
        "clean_headers": clean_headers,
    }


# ---------------------------------------------------------------------------
# Display / edit helpers
# ---------------------------------------------------------------------------

def editable_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    return str(value)



def metric_currency(value: Any) -> str:
    if value is None or value == "":
        return "-"
    try:
        return f"${float(value):,.0f}"
    except Exception:
        return str(value)



def metric_pct(value: Any) -> str:
    if value is None or value == "":
        return "-"
    try:
        return f"{float(value):.1%}"
    except Exception:
        return str(value)



def metric_date(value: Any) -> str:
    if value in (None, ""):
        return "-"
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, date):
        return value.strftime("%m/%d/%Y")
    return str(value)



def metric_int(value: Any) -> str:
    if value in (None, ""):
        return "-"
    try:
        return f"{int(float(value)):,}"
    except Exception:
        return str(value)



def preview_text(value: Any, limit: int = 90) -> str:
    text = editable_text(value).strip()
    if not text:
        return ""
    collapsed = re.sub(r"\s+", " ", text)
    if len(collapsed) <= limit:
        return collapsed
    return f"{collapsed[: limit - 1].rstrip()}…"



def quarter_code_options(start_year: int = 2024, end_year: int = 2028) -> list[str]:
    values = [""]
    for year in range(start_year, end_year + 1):
        yy = str(year)[-2:]
        for quarter in range(1, 5):
            values.append(f"{quarter}Q{yy}")
    values.extend(["TBD", "N/A", "NA", "NEW"])
    return values



def merge_options(base_options: list[str], observed_values: list[str]) -> list[str]:
    merged: list[str] = []
    seen: set[str] = set()
    for value in [*base_options, *observed_values]:
        text = editable_text(value).strip()
        if text in seen:
            continue
        merged.append(text)
        seen.add(text)
    return merged



def build_picklist_options(deals_df: pd.DataFrame) -> dict[str, list[str]]:
    options: dict[str, list[str]] = {}
    for header, base in PICKLIST_BASE_OPTIONS.items():
        observed: list[str] = []
        if header in deals_df.columns:
            series = deals_df[header].apply(editable_text).str.strip()
            counts = Counter(value for value in series.tolist() if value)
            observed = [value for value, _ in counts.most_common()]
        if header in {"Resolution Timing", "Liquidity Event Timing"}:
            options[header] = merge_options(quarter_code_options(), observed)
        else:
            options[header] = merge_options(base, observed)

    for header in ["Resolution Timing", "Liquidity Event Timing"]:
        if header not in options:
            options[header] = quarter_code_options()

    return options



def format_deal_label(row: pd.Series) -> str:
    return f"{row['Deal Number']} | {row['Deal Name']} | AM: {row['Asset Manager']}"



def make_editor_df(deals_df: pd.DataFrame) -> pd.DataFrame:
    editor_df = deals_df[["_excel_row", *GRID_CONTEXT_HEADERS, *EDITABLE_HEADERS, *READONLY_CONTEXT_HEADERS]].copy()
    for header in EDITABLE_HEADERS + READONLY_CONTEXT_HEADERS:
        editor_df[header] = editor_df[header].apply(editable_text)
    editor_df[PREVIOUS_WEEKLY_COMMENT_HEADER] = editor_df["Comments"].apply(editable_text)
    editor_df[THIS_WEEK_COMMENT_HEADER] = ""
    return editor_df



def build_comment_text(previous_comment: Any, this_week_comment: Any, as_of_label: str) -> str:
    previous_raw = editable_text(previous_comment)
    previous_clean = previous_raw.strip()
    current_clean = editable_text(this_week_comment).strip()
    if not current_clean:
        return previous_raw
    current_block = f"{as_of_label} - {current_clean}" if as_of_label else current_clean
    if not previous_clean:
        return current_block
    if previous_clean == current_block or previous_raw.startswith(f"{current_block}\n"):
        return previous_raw
    return f"{current_block}\n{previous_raw}"



def materialize_export_editor_df(editor_df: pd.DataFrame, as_of_label: str) -> pd.DataFrame:
    export_df = editor_df.copy()
    export_df["Comments"] = export_df.apply(
        lambda row: build_comment_text(
            row.get(PREVIOUS_WEEKLY_COMMENT_HEADER, ""),
            row.get(THIS_WEEK_COMMENT_HEADER, ""),
            as_of_label,
        ),
        axis=1,
    )
    return export_df



def build_filtered_view(base_deals: pd.DataFrame, editor_df: pd.DataFrame, as_of_label: str) -> pd.DataFrame:
    display_df = base_deals.copy()
    editor_lookup = editor_df.set_index("_excel_row")
    overlay_headers = EDITABLE_HEADERS + READONLY_CONTEXT_HEADERS + [PREVIOUS_WEEKLY_COMMENT_HEADER, THIS_WEEK_COMMENT_HEADER]
    for header in overlay_headers:
        if header in editor_lookup.columns:
            display_df[header] = display_df["_excel_row"].map(editor_lookup[header].to_dict())
    display_df["Comments"] = display_df.apply(
        lambda row: build_comment_text(
            row.get(PREVIOUS_WEEKLY_COMMENT_HEADER, ""),
            row.get(THIS_WEEK_COMMENT_HEADER, ""),
            as_of_label,
        ),
        axis=1,
    )
    return display_df



def normalize_for_compare(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value)



def workbook_change_set(editor_df: pd.DataFrame, original_editor_values: dict[int, dict[str, str]]) -> dict[tuple[int, str], str]:
    changes: dict[tuple[int, str], str] = {}
    row_lookup = editor_df.set_index("_excel_row")
    for excel_row, original in original_editor_values.items():
        if excel_row not in row_lookup.index:
            continue
        current = row_lookup.loc[excel_row]
        for header in EDITABLE_HEADERS:
            original_text = normalize_for_compare(original.get(header, ""))
            current_text = normalize_for_compare(current.get(header, ""))
            if current_text != original_text:
                changes[(excel_row, header)] = current_text
    return changes


# ---------------------------------------------------------------------------
# XML patch / workbook export helpers
# ---------------------------------------------------------------------------

def maybe_parse_number(text: str) -> int | float | None:
    candidate = text.strip()
    if not candidate:
        return None
    numeric_pattern = r"^\(?\$?[-+]?\d[\d,]*(\.\d+)?\)?$"
    if not re.match(numeric_pattern, candidate):
        return None
    negative = candidate.startswith("(") and candidate.endswith(")")
    cleaned = candidate.replace("$", "").replace(",", "").replace("(", "").replace(")", "")
    try:
        number = float(cleaned)
    except ValueError:
        return None
    if negative:
        number *= -1
    if number.is_integer():
        return int(number)
    return number



def maybe_parse_date(text: str) -> datetime | None:
    candidate = text.strip()
    if not candidate:
        return None
    for fmt in (
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%m/%d/%y",
        "%m-%d-%y",
        "%Y/%m/%d",
        "%b %d %Y",
        "%B %d %Y",
    ):
        try:
            return datetime.strptime(candidate, fmt)
        except ValueError:
            continue
    return None



def coerce_cell_value(text: str, header: str) -> Any:
    if text is None:
        return None
    if isinstance(text, float) and math.isnan(text):
        return None
    if not isinstance(text, str):
        return text
    if text == "":
        return None

    if header in DATE_OR_TEXT_HEADERS:
        parsed_date = maybe_parse_date(text)
        if parsed_date is not None:
            return parsed_date
        return text

    if header in NUMERIC_OR_TEXT_HEADERS:
        parsed_number = maybe_parse_number(text)
        if parsed_number is not None:
            return parsed_number
        return text

    return text



def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"



def get_or_create_cell(row_elem: ET.Element, coord: str) -> ET.Element:
    target_idx = column_index_from_string(re.sub(r"\d", "", coord))
    cell_tag = qn(MAIN_NS, "c")

    for cell in row_elem.findall(cell_tag):
        if cell.get("r") == coord:
            return cell

    new_cell = ET.Element(cell_tag, {"r": coord})
    inserted = False
    for pos, cell in enumerate(list(row_elem)):
        existing_ref = cell.get("r")
        if not existing_ref:
            continue
        existing_idx = column_index_from_string(re.sub(r"\d", "", existing_ref))
        if existing_idx > target_idx:
            row_elem.insert(pos, new_cell)
            inserted = True
            break
    if not inserted:
        row_elem.append(new_cell)
    return new_cell



def set_cell_xml_value(cell_elem: ET.Element, value: Any) -> None:
    for child in list(cell_elem):
        cell_elem.remove(child)
    cell_elem.attrib.pop("t", None)

    if value is None:
        return

    if isinstance(value, datetime):
        v = ET.SubElement(cell_elem, qn(MAIN_NS, "v"))
        v.text = str(to_excel(value))
        return

    if isinstance(value, date):
        v = ET.SubElement(cell_elem, qn(MAIN_NS, "v"))
        v.text = str(to_excel(datetime.combine(value, datetime.min.time())))
        return

    if isinstance(value, bool):
        cell_elem.set("t", "b")
        v = ET.SubElement(cell_elem, qn(MAIN_NS, "v"))
        v.text = "1" if value else "0"
        return

    if isinstance(value, (int, float)):
        v = ET.SubElement(cell_elem, qn(MAIN_NS, "v"))
        v.text = str(value)
        return

    cell_elem.set("t", "inlineStr")
    is_elem = ET.SubElement(cell_elem, qn(MAIN_NS, "is"))
    t_elem = ET.SubElement(is_elem, qn(MAIN_NS, "t"))
    t_elem.set(f"{{{XML_SPACE}}}space", "preserve")
    t_elem.text = str(value)



def patch_sheet_xml(sheet_xml: bytes, changes: dict[tuple[int, str], str], cell_meta: dict[tuple[int, str], dict[str, Any]]) -> bytes:
    root = ET.fromstring(sheet_xml)
    sheet_data = root.find(qn(MAIN_NS, "sheetData"))
    if sheet_data is None:
        return sheet_xml

    row_map = {int(row.get("r")): row for row in sheet_data.findall(qn(MAIN_NS, "row"))}

    for (excel_row, header), text_value in changes.items():
        meta = cell_meta.get((excel_row, header))
        if not meta:
            continue
        row_elem = row_map.get(excel_row)
        if row_elem is None:
            row_elem = ET.SubElement(sheet_data, qn(MAIN_NS, "row"), {"r": str(excel_row)})
            row_map[excel_row] = row_elem
        cell_elem = get_or_create_cell(row_elem, meta["coord"])
        coerced = coerce_cell_value(text_value, header)
        set_cell_xml_value(cell_elem, coerced)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)



def patch_workbook_xml(workbook_xml: bytes) -> bytes:
    root = ET.fromstring(workbook_xml)
    calc_pr = root.find(qn(MAIN_NS, "calcPr"))
    if calc_pr is None:
        calc_pr = ET.SubElement(root, qn(MAIN_NS, "calcPr"))
    calc_pr.set("calcMode", "auto")
    calc_pr.set("fullCalcOnLoad", "1")
    calc_pr.set("forceFullCalc", "1")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)



def patch_workbook_rels_xml(rels_xml: bytes) -> bytes:
    root = ET.fromstring(rels_xml)
    calcchain_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
    for rel in list(root):
        if rel.get("Type") == calcchain_type:
            root.remove(rel)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)



def patch_content_types_xml(content_xml: bytes) -> bytes:
    root = ET.fromstring(content_xml)
    calc_part = "/xl/calcChain.xml"
    for child in list(root):
        if child.get("PartName") == calc_part:
            root.remove(child)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)



def build_updated_workbook(file_bytes: bytes, sheet_path: str, changes: dict[tuple[int, str], str], cell_meta: dict[tuple[int, str], dict[str, Any]]) -> bytes:
    source = io.BytesIO(file_bytes)
    output = io.BytesIO()

    with ZipFile(source, "r") as zin, ZipFile(output, "w", compression=ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == sheet_path:
                data = patch_sheet_xml(data, changes, cell_meta)
            elif info.filename == "xl/workbook.xml":
                data = patch_workbook_xml(data)
            elif info.filename == "xl/_rels/workbook.xml.rels":
                data = patch_workbook_rels_xml(data)
            elif info.filename == "[Content_Types].xml":
                data = patch_content_types_xml(data)
            elif info.filename == "xl/calcChain.xml":
                continue

            new_info = info
            new_info.compress_type = ZIP_DEFLATED
            zout.writestr(new_info, data)

    return output.getvalue()



def apply_detail_edit(editor_df: pd.DataFrame, excel_row: int, values: dict[str, str]) -> pd.DataFrame:
    updated = editor_df.copy()
    mask = updated["_excel_row"] == excel_row
    for header, value in values.items():
        if header in updated.columns:
            updated.loc[mask, header] = editable_text(value)
    return updated
