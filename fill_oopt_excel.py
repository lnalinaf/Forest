"""
Fill the «Название ООПТ / нарушения…» column from OOPT_territories .doc lists.

Example::

    .venv/bin/python fill_oopt_excel.py \\
        --input \"test data/need_to_be_filled.xlsx\" \\
        --output \"test data/filled_oopt.xlsx\"
"""

from __future__ import annotations

import argparse
import os

from openpyxl import load_workbook
import pandas as pd

from oopt_registry import get_cached_rules, match_oopt, parse_kv_cell
from setting import SETTINGS
from process_pdf import get_target_pdf, get_doc_text, extract_info_v2
from download_data import download_unzip


def _pick_column(columns: list, must_include: list[str], exclude: tuple[str, ...] = ()) -> str | None:
    for c in columns:
        s = str(c).strip().lower()
        if any(e in s for e in exclude):
            continue
        if all(m in s for m in must_include):
            return c
    return None


def extract_hyperlinks_from_excel(path: str, id_col_name: str = "акт") -> dict:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    
    col_idx = None
    for cell in ws[1]:
        if cell.value and id_col_name.lower() in str(cell.value).strip().lower():
            col_idx = cell.column
            break
            
    if not col_idx:
        return {}
        
    links = {}
    for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, min_row=2):
        cell = row[0]
        if cell.value and cell.hyperlink:
            links[str(cell.value).strip()] = cell.hyperlink.target
    return links

def fill_dataframe_oopt(df: pd.DataFrame, hyperlinks: dict, rules_dir: str | None = None, data_dir: str | None = None) -> pd.DataFrame:
    rules_dir = rules_dir or SETTINGS.oopt_territories_dir
    rules = get_cached_rules(os.path.abspath(rules_dir))

    cols = list(df.columns)
    c_oopt = _pick_column(cols, ["оопт", "название"]) or _pick_column(cols, ["оопт"])
    c_lesn = _pick_column(cols, ["лесничество"], exclude=("участков", "участковое"))
    c_uch = _pick_column(cols, ["участковое"])
    c_kv = _pick_column(cols, ["квартал"])
    c_akt = _pick_column(cols, ["акт"])
    c_area = _pick_column(cols, ["площадь"])
    c_vid = _pick_column(cols, ["вид", "рубки"])
    c_status = _pick_column(cols, ["статус"])

    if not c_oopt:
        raise ValueError("Could not find ООПТ column (header should mention «Название ООПТ»).")

    out = df.copy()
    out[c_oopt] = out[c_oopt].astype(object)
    if c_lesn: out[c_lesn] = out[c_lesn].astype(object)
    if c_uch: out[c_uch] = out[c_uch].astype(object)
    if c_area: out[c_area] = out[c_area].astype(object)
    if c_kv: out[c_kv] = out[c_kv].astype(object)
    if c_vid: out[c_vid] = out[c_vid].astype(object)
    if c_status: out[c_status] = out[c_status].astype(object)
    
    if data_dir and not os.path.exists(data_dir):
        os.makedirs(data_dir, exist_ok=True)

    for i in out.index:
        akt_val = out.at[i, c_akt] if c_akt else None
        if pd.isna(akt_val) or not str(akt_val).strip():
            continue
            
        akt_str = str(akt_val).strip()
        link = hyperlinks.get(akt_str)
        if not link:
            continue
            
        print(f"Processing: {akt_str}")
        extracted_folder = download_unzip(akt_str, link, data_dir)
        if not extracted_folder:
            continue
            
        pdf_path = get_target_pdf(extracted_folder)
        if not pdf_path:
            print(f"No PDF found for {akt_str}")
            continue
            
        doc_text = get_doc_text(pdf_path)
        forest_name, zone_name, area, zone_1, zone_2, full_event_name, parse_status = extract_info_v2(doc_text, pdf_path=pdf_path)
        
        if parse_status == "wrong format":
            print(f"Wrong format for {akt_str}")
            if c_status:
                out.at[i, c_status] = "wrong format"
            continue

        if c_lesn: out.at[i, c_lesn] = forest_name
        if c_uch: out.at[i, c_uch] = zone_name
        if c_area: out.at[i, c_area] = area
        if c_kv and zone_1 and zone_2: out.at[i, c_kv] = f"{zone_1}({zone_2})" # Kvartal(Vydel)
        if c_vid: out.at[i, c_vid] = full_event_name
        
        uch = zone_name or forest_name
        kv = int(zone_1) if zone_1 and str(zone_1).isdigit() else None
        if kv is not None:
            m = match_oopt(rules, None, uch, kv)
            if m:
                out.at[i, c_oopt] = m

    return out



def main() -> None:
    p = argparse.ArgumentParser(description="Fill ООПТ column from territory .doc lists.")
    p.add_argument("--input", "-i", required=True, help="Input .xlsx path")
    p.add_argument("--output", "-o", help="Output .xlsx (default: overwrite input)")
    p.add_argument(
        "--oopt-dir",
        default=None,
        help="OOPT_territories folder (default: SETTINGS.oopt_territories_dir)",
    )
    args = p.parse_args()
    path_in = os.path.abspath(args.input)
    path_out = os.path.abspath(args.output) if args.output else path_in
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

    hyperlinks = extract_hyperlinks_from_excel(path_in)
    print(f"Found {len(hyperlinks)} hyperlinks.")

    df = pd.read_excel(path_in)
    filled = fill_dataframe_oopt(df, hyperlinks, rules_dir=args.oopt_dir, data_dir=data_dir)
    filled.to_excel(path_out, index=False)
    print(f"Wrote {path_out}")


if __name__ == "__main__":
    main()
