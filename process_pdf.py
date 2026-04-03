import re
from pdf2image import convert_from_path
from pytesseract import image_to_string
import os
import pandas as pd
import argparse
from setting import SETTINGS
from oopt_registry import get_cached_rules, match_oopt

# Path to your PDF file
forest_data_dir = SETTINGS.forest_data_dir
volunteer_name = SETTINGS.volunteer_name
_oopt_rules = None


def _rules(oopt_dir=None):
    global _oopt_rules
    if _oopt_rules is None:
        path = oopt_dir or os.path.abspath(SETTINGS.oopt_territories_dir)
        _oopt_rules = get_cached_rules(path)
    return _oopt_rules

initial_forest_data_xlsx = os.path.join(forest_data_dir, "Initial_forest_data.xlsx")

pattern_name_dict = {
    # High Priority: "СОМ не требуется" variants
    r'(?:мероприятия[\s\n\w,]*не[\s\n]+требуются|[сc]анитарно[\s\n\-]*оздоровительн[ых\s\n\-]+мероприяти[йя\s\n\-]+не[\s\n\-]*требует[ся\s\n]*|мероприятия по предупреждению распространения вредных организмов не требуются)': 'СОМ не требуется',
    r'мероприятия[\s\n]+не[\s\n]+требуются': 'СОМ не требуется',
    r'[сc]ом[\s\n]+не[\s\n]+требуется': 'СОМ не требуется', 
    # Standard Rubka Types
    r'[сc]плошн[аяой]+[\s\n\-]+санитарн[аяой]+[\s\n\-]+рубк[аи]': 'ССР',
    r'выборочн[аяой]+[\s\n\-]+санитарн[аяой]+[\s\n\-]+рубк[аи]': 'ВСР',
    r'рубк[аи][\s\n\-]+аварийн[ыхое]+[\s\n\-]+деревьев': 'РАД',
    r'\bРАД\b': 'РАД',
    # Others
    r'размножения[\s\n]+насекомоядных(?:\s[^\n]*)птиц[\s\n]+и[\s\n]+других[\s\n]+насекомоядных': 'Улучшение условий обитания и размножения насекомоядных год птиц и других насекомоядных животных',
    r'\b[сc]рс\b': 'ССР', 
    r'\b[сc]рв\b': 'ВСР',
    r'улучшение[\s\n]+условий[\s\n]+обитания': 'Улучшение условий обитания и размножения насекомоядных'
}

checked_status = 'Проверено'

def get_target_pdf(folder_path):
    """Finds the most relevant PDF file in the given folder path, prioritizing 'Акт' and excluding 'абрис'."""
    pdf_files = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            name_lower = file.lower()
            if file.endswith('.pdf') and 'абрис' not in name_lower and not file.endswith('.sig'):
                full_path = os.path.join(root, file)
                pdf_files.append(full_path)
    
    if not pdf_files: return None
    
    # Prioritize files starting with 'Акт' or having 'акт' in the name
    akt_files = [f for f in pdf_files if 'акт' in os.path.basename(f).lower()]
    # Sort to prefer 'Акт №...' over 'абрис к акту...' (though abris is already excluded)
    akt_files.sort(key=lambda x: (not os.path.basename(x).startswith('Акт'), len(os.path.basename(x))))
    
    return akt_files[0] if akt_files else pdf_files[0]

def get_doc_text(path_of_pdf):
    if not path_of_pdf:
        raise FileNotFoundError(f"The file '{path_of_pdf}' does not exist.")
    temp_folder = "temp_images"
    os.makedirs(temp_folder, exist_ok=True)

    all_text = []
    pages = convert_from_path(path_of_pdf, dpi=300, output_folder=temp_folder)
    for page_number, page_image in enumerate(pages, start=1):
        text = image_to_string(page_image, lang="rus")
        all_text.append(f"--- Page {page_number} ---\n{text}\n")

    document_text = "\n".join(all_text)
    for image_file in os.listdir(temp_folder):
        os.remove(os.path.join(temp_folder, image_file))
    os.rmdir(temp_folder)
    return document_text

def extract_info_v2(document_text, pdf_path=None):
    """
    Robust extraction with multiple fallback strategies.
    
    Returns: (forest_name, zone_name, area, zone_1, zone_2, full_event_name, status)
    """
    rotated_texts = []
    # --- 1. Header Detection (with Rotation Recovery) ---
    zakl_broad = r'ЗАКЛЮЧЕНИЕ'
    zakl_specific = r'ЗАКЛЮЧЕНИЕ\s+к\s+инструментальному\s+обследованию\s+(?:лесного\s+)?участка'
    
    has_zakl = re.search(zakl_broad, document_text, re.IGNORECASE)
    has_akt = re.search(r'Акт\s+лесопатологического\s+обследования', document_text, re.IGNORECASE)

    if not has_zakl and not has_akt and pdf_path:
        # Header Recovery: Try rotating the first few pages to find "ЗАКЛЮЧЕНИЕ"
        print(f"Header not found, trying rotation recovery for {pdf_path}...")
        try:
            # Using global imports: convert_from_path, pytesseract, image_to_string
            # Try first 10 pages to handle long notifications before the act
            pages = convert_from_path(pdf_path, dpi=200, first_page=1, last_page=10)
            for page in pages:
                for angle in [90, -90]:
                    rotated = page.rotate(angle, expand=True)
                    text = image_to_string(rotated, lang='rus')
                    if re.search(zakl_broad, text, re.IGNORECASE):
                        print(f"Header found after {angle} degree rotation!")
                        has_zakl = True
                        document_text += "\n" + text # Append rotated text to main search
                        rotated_texts.append(text)
                        break
                if has_zakl: break
        except Exception as e:
            print(f"Header recovery failed: {e}")

    if not has_zakl and not has_akt:
        return None, None, None, None, None, None, "wrong format"

    # Re-evaluate zakl_specific after potentially finding header via rotation
    zakl_specific_match = re.search(zakl_specific, document_text, re.IGNORECASE)

    # --- 2. Extract Forest (Лесничество) ---
    forest_match = re.search(r'насаждений\s+([\w\-\s]+(?:\(лесничество\))?)', document_text, re.IGNORECASE)
    forest_name_raw = forest_match.group(1).strip() if forest_match else None
    forest_name = re.sub(r'\s*\(лесничество\)\s*', '', forest_name_raw) if forest_name_raw else None
    # Normalize Genitive case (e.g., "Всеволожского" -> "Всеволожское")
    if forest_name:
        forest_name = re.sub(r'([А-Яа-яЁё\-]+)(?:ого|его|ое|ому|ым|ом)$', r'\1ое', forest_name)

    # --- 3. Extract Участковое лесничество, Квартал, Выдел ---
    # District (Uchastkovoe) - Support "уч. лесничество" abbreviation
    uch_match = re.search(r'(?:Участковое|уч\.?)\s+лесничество[:\s]+([\w\-]+)', document_text, re.IGNORECASE)
    zone_name = uch_match.group(1).strip() if uch_match else None

    area = None
    zone_1 = None  # Kvartal
    zone_2 = None  # Vydel
    full_event_name = 'NO MATCH'
    def is_junk_id(val, text, match_start):
        # Section headers, versions, and appendixes: 1.1, 1.2, 2.2 etc.
        if val in ['1.0', '1.1', '1.2', '1.3', '1.4', '2.0', '2.1', '2.2', '2.3', '3.1', '3.2']:
            context_before = text[max(0, match_start-50):match_start]
            context_after = text[match_start+len(val):match_start+len(val)+50]
            # Sections almost always have a trailing dot or specific keywords
            if re.search(r'Прил|Лист|Схема|подписью|№|верс', context_before, re.IGNORECASE): return True
            if re.search(r'Метод|срок|действ|к\s+акту|^\.', context_after, re.IGNORECASE): return True
            # Small decimals in the first 2 pages are almost always noise
            if match_start < 2000: return True
        return False

    # Priority 1: Strict pair match "Квартал X Выдел Y" - also support "кв." abbreviation
    kv_matches = re.finditer(r'(?:Квартал|кв\.?)\s*(\d+)', document_text, re.IGNORECASE)
    for kvm in kv_matches:
        kv_val = kvm.group(1)
        search_near = document_text[kvm.end():kvm.end()+300]
        vyd_match = re.search(r'(?:Выдел|выдел)[^\d]*(\d+[.,]?\d*)', search_near, re.IGNORECASE)
        if vyd_match:
            vyd_val = vyd_match.group(1).replace(',', '.')
            if not is_junk_id(vyd_val, document_text, kvm.end() + vyd_match.start(1)):
                # If we have an integer vydel, and then find a decimal, the decimal is BETTER (subvydel)
                if zone_1 is None or ('.' in vyd_val and (zone_2 is None or '.' not in str(zone_2))):
                    zone_1, zone_2 = kv_val, vyd_val
                    # Attempt row-based district search
                    lookback = document_text[max(0, kvm.start()-50):kvm.start()]
                    uch_row = re.search(r'([\w\-]{3,20}[сc]кое)', lookback, re.IGNORECASE)
                    if uch_row: zone_name = uch_row.group(1).strip()

    # Priority 2: Specific "лесопатологический выдел"
    if zone_2 is None or '.' not in str(zone_2):
        vyd_decs = re.finditer(r'лесопатологически[йе]\s+выдел[^\d]*(\d+[.,]\d+)', document_text, re.IGNORECASE)
        for vd in vyd_decs:
            v = vd.group(1).replace(',', '.')
            if not is_junk_id(v, document_text, vd.start(1)):
                zone_2 = v
                kv_near_m = re.search(r'Квартал[^\d]*(\d+)', document_text[max(0, vd.start()-300):vd.end()+300], re.IGNORECASE)
                if kv_near_m: 
                    zone_1 = kv_near_m.group(1)
                    # Attempt row-based district
                    lookback = document_text[max(0, kv_near_m.start()-50):kv_near_m.start()]
                    uch_row = re.search(r'([\w\-]{3,20}[сc]кое)', lookback, re.IGNORECASE)
                    if uch_row: zone_name = uch_row.group(1).strip()
                break

    # Final forest_name fallback
    forest_name = forest_name or forest_name_raw

    # --- 4. Find the ЗАКЛЮЧЕНИЕ page number ---
    zakl_match = zakl_specific_match or has_zakl
    zakl_page_num = None
    page_splits = document_text.split("--- Page ")
    for ps in page_splits:
        if re.search(zakl_specific, ps, re.IGNORECASE) or re.search(zakl_broad, ps, re.IGNORECASE):
            m_page = re.match(r'(\d+)', ps)
            if m_page:
                zakl_page_num = int(m_page.group(1))
                break

    # --- 5. Build combined search text ---
    if zakl_match:
        # Use context after zakl_match
        search_text = document_text[zakl_match.start():]
    else:
        search_text = document_text

    # --- 6. Event type ---
    for pattern, label in pattern_name_dict.items():
        if re.search(pattern, search_text, re.IGNORECASE):
            full_event_name = label
            break

    # --- 7. FALLBACK: omni-rotation OCR ---
    rotated_texts = []
    if full_event_name == 'NO MATCH' and pdf_path and zakl_page_num:
        try:
            pages = convert_from_path(pdf_path, dpi=300, first_page=zakl_page_num, last_page=zakl_page_num)
            if pages:
                page_img = pages[0]
                for rot in [90, -90]:
                    rot_img = page_img.rotate(rot, expand=True)
                    rot_text = image_to_string(rot_img, lang="rus")
                    rotated_texts.append(rot_text)
                    for pattern, label in pattern_name_dict.items():
                        if re.search(pattern, rot_text, re.IGNORECASE):
                            full_event_name = label
                            break
                    if full_event_name != 'NO MATCH': break
        except Exception as e:
            print(f"Rotation OCR failed for {pdf_path}: {e}")

    # --- 8. Area extraction ---
    combined_zakl_text = search_text
    for rt in rotated_texts:
        combined_zakl_text += "\n" + rt

    def clean_area(val):
        val = val.replace(',', '.')
        # Common OCR mistake: missing dot in small areas (04->0.4, 13->1.3, 64->6.4)
        if re.match(r'^\d\d$', val): 
            return f"{val[0]}.{val[1]}"
        return val

    def is_likely_date(val, text, match_start):
        context = text[max(0, match_start-5):match_start+len(val)+15]
        # Dates are usually X.Y.Z. Areas are X.Y га.
        if re.search(r'га\b', context, re.IGNORECASE): return False
        if re.search(r'\d{2}[.,]\d{2}[.,]\d{4}', context): return True
        if re.search(rf'{re.escape(val)}[.,]\d{{4}}', context): return True
        return False

    # Strategy for Area candidates
    area_candidates = [] # List of tuples (priority, value)

    # Patterns with explicit priorities (0 = highest)
    area_patterns = [
        (0, r'(?:на\s+площади|площадь\s+лесопатологического\s+выдела)[^\d]{0,15}(\d+[.,]?\d*)\s*га'),
        (1, r'на\s+общей\s+площади[^\d]{0,20}(\d+[.,]?\d*)'),
        (2, r'площад[ьиьь][: ]+лесопатологическ[ий\sаго]+выдел[а]?[^\d]{0,15}(\d+[.,]?\d*)\s*га'),
        (3, r'площад[ьиьь][: ]+мероприятия[^\d]{0,15}(\d+[.,]?\d*)\s*га'),
        (4, r'площад[ьиьь][^\d]{0,15}(\d+[.,]?\d*)\s*га'),
        (5, r'(\d+[.,]?\d*)\s*га')
    ]

    # Two-pass extraction: 
    # 1. Search in search_text (Conclusion) - High Priority
    # 2. If nothing found, search in document_text (Full) - Low Priority
    for scope_text in [search_text, combined_zakl_text]:
        if area_candidates: break # We are happy with candidates from the most specific section
        for priority, kp in area_patterns:
            m_areas = re.finditer(kp, scope_text, re.IGNORECASE)
            for ma in m_areas:
                val_str = ma.group(1)
                val = clean_area(val_str)
                # Filter out obvious noise (section IDs like 1.1) unless accompanied by area markers
                if val in ['1.0', '1.1', '1.2', '1.3', '1.4', '2.0', '2.1', '2.2', '2.3']:
                    if is_junk_id(val_str, scope_text, ma.start(1)): 
                        # Check if 'площадь' or 'га' is VERY close (re-verification)
                        if not re.search(r'площадь|га', scope_text[max(0, ma.start()-15):ma.end()+15], re.IGNORECASE):
                            continue

                try:
                    fval = float(val)
                    if not (2000 <= fval <= 2100) and fval < 500 and not is_likely_date(val_str, scope_text, ma.start(1)):
                        area_candidates.append((priority, val))
                except: continue

    # Selection Logic: Prioritize by pattern priority, then by "good decimal" status
    area = None
    if area_candidates:
        # Sort by priority first
        area_candidates.sort(key=lambda x: x[0])
        # Find best candidate within the highest available priority group
        best_priority = area_candidates[0][0]
        group = [c[1] for c in area_candidates if c[0] == best_priority]
        
        good_decimals = [c for c in group if '.' in c and c not in ['1.1', '1.2', '1.3', '1.4', '2.1', '2.2']]
        if good_decimals: 
            # Pick the largest value in the highest priority group to avoid picking small sub-areas
            good_decimals.sort(key=lambda x: float(x), reverse=True)
            area = good_decimals[0]
        else:
            all_decimals = [c for c in group if '.' in c]
            if all_decimals:
                all_decimals.sort(key=lambda x: float(x), reverse=True)
                area = all_decimals[0]
            else:
                group.sort(key=lambda x: float(x.replace(',', '.')), reverse=True)
                area = group[0]

    return forest_name, zone_name, area, zone_1, zone_2, full_event_name, "ok"


def main():
    parser = argparse.ArgumentParser(description="Forest Survey Data Extraction Tool")
    parser.add_argument("--input", default="test data/need_to_be_filled.xlsx", help="Path to input Excel template")
    parser.add_argument("--pdf", default="data", help="Root directory for PDF files")
    parser.add_argument("--output", default="test data/out_with_oopt.xlsx", help="Path to save output Excel")
    parser.add_argument("--oopt", default="OOPT_territories", help="Directory with OOPT reference files")
    
    args = parser.parse_args()
    
    excel_path = args.input
    pdf_root = args.pdf
    output_path = args.output
    oopt_dir = args.oopt

    if not os.path.exists(excel_path):
        print(f"Error: {excel_path} not found.")
        return

    df = pd.read_excel(excel_path)
    # Ensure columns are object type to avoid TypeError when writing strings to empty columns
    for col in ['Лесничество', 'Участковое лесничество', 'Вид рубки', 'Статус', 'площадь', 'Кварталы, выделы ']:
        if col in df.columns:
            df[col] = df[col].astype(object)
    # Ensure columns used for strings are object type to avoid float conversion errors
    string_cols = [
        'Статус', 'площадь',
        'Название ООПТ / нарушения (водоохранная зона / документальные или смысловые расхождения внутри акта) ',
        'Кварталы, выделы ', 'Вид рубки', 'Кто заполнял таблицу', 'Участковое лесничество'
    ]
    for col in string_cols:
        if col in df.columns:
            df[col] = df[col].astype(object)

    for index, row in df.iterrows():
        # Match the lowercase 'акт' column name from the template
        filename = str(row.get('акт', ''))
        if not filename or filename == 'nan': continue
        
        # Folder name in 'data/' i.e. "51-1 Куйбышевское"
        target_folder_path = os.path.join(pdf_root, filename.replace('/', '.'))
        if os.path.isdir(target_folder_path):
            pdf_path = get_target_pdf(target_folder_path)
            if not pdf_path: 
                print(f"No PDF found in {target_folder_path}")
                continue

            print(f"Processing {pdf_path}...")
            document_text = get_doc_text(pdf_path)
            extracted_data = extract_info_v2(document_text, pdf_path=pdf_path)
            forest_name, zone_name, area, zone_1, zone_2, full_event_name, status = extracted_data

            if status == "wrong format":
                full_event_name = "wrong format"

            # ---------------------------------------------------------
            # Principle: If 'РАД' or 'СОМ не требуется' - clear Area and Zones
            # ---------------------------------------------------------
            if full_event_name in ['РАД', 'СОМ не требуется']:
                area = None
                zone_1, zone_2 = None, None

            oopt_label = None
            if zone_1 is not None and str(zone_1).strip().isdigit():
                uch = zone_name or forest_name
                kv = int(zone_1)
                # Pass forest_name to match_oopt to restrict search to that forest's file
                m = match_oopt(_rules(oopt_dir), forest_name, uch, kv)
                if m: oopt_label = m

            df.loc[index, 'Статус'] = 'проверено' if full_event_name not in ['NO MATCH', 'wrong format'] else 'Проверить'
            df.loc[index, 'Название ООПТ / нарушения (водоохранная зона / документальные или смысловые расхождения внутри акта) '] = oopt_label
            
            # Match columns from the excel template precisely
            if area is not None:
                try:
                    area_val = float(area)
                    # Convert to string with comma for Russian Excel format if needed
                    df.loc[index, 'площадь'] = str(area_val).replace('.', ',')
                except:
                    df.loc[index, 'площадь'] = area
            else:
                df.loc[index, 'площадь'] = None
            
            # Format: Kvartal(Vydely)
            df.loc[index, 'Кварталы, выделы '] = f'{zone_1}({zone_2})' if zone_1 and zone_2 else None
            df.loc[index, 'Вид рубки'] = str(full_event_name)
            df.loc[index, 'Кто заполнял таблицу'] = str(volunteer_name)
            df.loc[index, 'Участковое лесничество'] = str(zone_name) if zone_name else None
            df.loc[index, 'Лесничество'] = str(forest_name) if forest_name else None
    
    # Clean up redundant columns
    if 'Status' in df.columns: df.drop(columns=['Status'], inplace=True)
    
    df.to_excel(output_path, index=False)
    print(f"Done! Saved to {output_path}")

if __name__ == "__main__":
    main()
