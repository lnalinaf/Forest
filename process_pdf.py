import re
from pdf2image import convert_from_path
from pytesseract import image_to_string
import os
import pandas as pd
from setting import SETTINGS

# Path to your PDF file
forest_data_dir = SETTINGS.forest_data_dir
volunteer_name = SETTINGS.volunteer_name
initial_forest_data_xlsx = os.path.join(forest_data_dir, "Initial_forest_data.xlsx")
# target_folder_path = "/Users/alina/Downloads/004-ТМ-41-24_47_Ленинское"
events_names = {'размножения насекомоядных птиц и других насекомоядных':
                         'Улучшение условий обитания и размножения насекомоядных год птиц и других насекомоядных животных',
                     'СОМ не требуется':'СОМ не требуется'}

location_patterns = [r'кв\.\s*(\d+).*?выд\.\s*(\d+)', r'в выделе (\d+) квартала (\d+) (\S+)', r'кв\s*(\d+)\s*выд\s*(\d+)']
event_patterns = [r'размножения насекомоядных(?:\s[^\n]*)птиц и других насекомоядных', r'([А-Я][а-я]+).*\s(СОМ не требуется)\s',
                  r'(СРС)', r'(СРВ)',
                  r'СОМ(\n|\s){0,2}не(\n|\s){0,2}требуется', r'(Улучшение условий)\s.*?(обитания и ,размножения)\s.*?(насекомоядных птиц и)\s.*?(других насекомоядных\sи\sживотных)']
pattern_name_dict = {r'СОМ(\n|\s){0,2}не(\n|\s){0,2}требуется': 'СОМ не требуется', r'размножения насекомоядных(?:\s[^\n]*)птиц и других насекомоядных':
                     'Улучшение условий обитания и размножения насекомоядных год птиц и других насекомоядных животных',
                     r'(Улучшение условий)\s.*?(обитания)\s.*?(размножения)\s.*?(насекомоядных птиц)\s.*?(других)\s.*?(насекомоядных)\s.*?(животных)':
                     'Улучшение условий обитания и размножения насекомоядных год птиц и других насекомоядных животных',
                     r'(СРС)':'СРС', r'(СРВ)':'СРВ', r"(Охрана(?:.|\s)*?местообитаний(?:.|\s)*?насекомых)":"Охрана местообитаний"}

checked_status = 'Проверено'

area_patterns =[r"обследование проведено\sна\s(.*\s)?пло.ад.\s(\d+,?\d*)\s?га", r"на\s(общей\s)?пло.ад.\s(\d+[.,]?\d*)"]


def get_target_pdf(folder_path):
    """Finds the first PDF file in the given folder path, including subfolders."""
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.pdf'):
                return os.path.join(root, file)  # Return the full path to the first PDF found
    return None


def get_doc_text(path_of_pdf):
    if not path_of_pdf:
        raise FileNotFoundError(f"The file '{path_of_pdf}' does not exist.")
    # Temporary folder to store images
    temp_folder = "temp_images"
    os.makedirs(temp_folder, exist_ok=True)

    # Perform OCR on all pages
    all_text = []
    pages = convert_from_path(path_of_pdf, dpi=300, output_folder=temp_folder)
    for page_number, page_image in enumerate(pages, start=1):
        # Perform OCR with Russian language support
        text = image_to_string(page_image, lang="rus")
        all_text.append(f"--- Page {page_number} ---\n{text}\n")

    # Combine all text into one string
    document_text = "\n".join(all_text)
    # Clean up temporary images
    for image_file in os.listdir(temp_folder):
        os.remove(os.path.join(temp_folder, image_file))
    os.rmdir(temp_folder)
    return document_text


def extract_info(document_text):
    # Forest name search
    pattern = r'насаждений\s([\w\-\s]+(?:\(лесничество\))?)'
    match = re.search(pattern, document_text)
    forest_name = match.group(1) if match else None
    if forest_name:
        forest_name = forest_name.replace('кого', 'кое')
    print("forest_name", forest_name)

    # Area
    area = None
    for area_pattern in area_patterns:
        area_match = re.search(area_pattern, document_text)
        if area_match:
            area = area_match.group(2)

    print('area S', area)

    # Location
    for i, location_pattern in enumerate(location_patterns):
        location_match = re.search(location_pattern, document_text)
        if location_match:
            if i == 1:  # For the first pattern in the list
                zone_1 = location_match.group(2)  # квартал
                zone_2 = location_match.group(1)  # выдел
            else:  # For all other patterns
                zone_1 = location_match.group(1)  # квартал
                zone_2 = location_match.group(2)  # выдел

            print('zone_1 выдел', zone_1, 'zone_2 квартал', zone_2)

            zone_name = location_match.group(3) if len(location_match.groups()) >= 3 else None
        else:
            zone_1 = zone_2 = zone_name = None
        if zone_name:
            zone_name = zone_name.replace('кого', 'кое')

        print(zone_1, zone_2, zone_name)

    # Event type
    full_event_name = 'NO MATCH'
    for event_pattern in event_patterns:
        event_match = re.search(event_pattern, document_text, flags=re.MULTILINE)
        if event_match and pattern_name_dict.get(event_pattern):
            full_event_name = pattern_name_dict.get(event_pattern)
        elif event_match:
            if zone_name is None:
                zone_name = event_match.group(1)  # Extracts 'Шидрозерское'
                print(f"Zone Name: {zone_name}")
            event = event_match.group(2)  # Extracts 'СОМ не требуется'
            print(f"Event: {event}")
            full_event_name = events_names.get(event, "STILL NO NAME")
    print('full_event_name', full_event_name)
    return forest_name, zone_name, area, zone_1, zone_2, full_event_name



def write_to_common_dataframe(data_tuple, combined_data):
    forest_name, zone_name, area, zone_1, zone_2, full_event_name = data_tuple

    # Wrap scalar values in lists to make them compatible with pandas.DataFrame
    new_data = {
        'Status': ['проверено'],
        'OOPT': [''],
        'big_zone': [forest_name],
        'small_zone': [zone_name],
        'square': [area],
        'zones': [f'{zone_2}({zone_1})'],
        'event': [full_event_name],
        'rent': [''],
        'volunteer': [volunteer_name]
    }

    # Convert the dictionary to a DataFrame
    new_df = pd.DataFrame(data=new_data)

    # Append the new DataFrame to the combined one
    combined_data = pd.concat([combined_data, new_df], ignore_index=True)

    return combined_data


def main():
    df = pd.read_excel(initial_forest_data_xlsx)
    # Iterate over each file in "Name" column
    for index, row in df.iterrows():
        filename = row['Name']
        # Construct the file path
        target_folder_path = os.path.join(forest_data_dir, filename.replace('/', '.'))
        if os.path.isdir(target_folder_path):
            pdf_path = get_target_pdf(target_folder_path)
            if not pdf_path:
                print(f'pdf does not exist for {target_folder_path.path}')
                continue

            document_text = get_doc_text(pdf_path)
            output_file_path = os.path.join(target_folder_path, "extracted_text.txt")
            with open(output_file_path, "w", encoding="utf-8") as file:
                file.write(document_text)
            extracted_data = extract_info(document_text)
            forest_name, zone_name, area, zone_1, zone_2, full_event_name = extracted_data

            df.loc[index, 'Status'] = checked_status if full_event_name!='NO MATCH' else 'Проверить'
            df.loc[index, 'OOPT'] = ''
            df.loc[index, 'forest_name'] = forest_name
            df.loc[index, 'zone_name'] = zone_name
            df.loc[index, 'square'] = area
            df.loc[index, 'zones'] = f'{zone_2}({zone_1})'
            df.loc[index, 'event'] = full_event_name
            df.loc[index, 'rent'] = ''
            df.loc[index, 'volunteer'] = volunteer_name
        else:
            print(f'No directory {target_folder_path}')
    df.to_excel(initial_forest_data_xlsx, index=False)

main()
