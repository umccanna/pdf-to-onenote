import pdfplumber
from config import get_config
import re
import getopt
import sys
from pathlib import Path 
from tabulate import tabulate
import json

def normalize_text(text: str):
    text = re.sub(r'\s+',  ' ', text).strip()
    text = re.sub(r". ,","",text)
    # remove all instances of multiple spaces
    text = text.replace("..",".")
    text = text.replace(". .",".")
    text = text.replace("\n", "")
    text = text.replace("\r", "")

    # replace unicode characters with more usable characters 
    text = text.replace("●", "*")
    text = text.replace("‑", "-")
    text = text.strip()
    
    return text

def strip_emails_and_phone_numbers_and_web_addresses(text):
    text = re.sub(r'(?:[a-zA-Z]|\d|\-|\.)*@cms\.hhs\.gov', "", text).strip()
    text = re.sub(r'\(?\d{3}\)?\-?\d{3}\-\d{4}', "", text).strip()
    text = re.sub(r'(https?|ftp):\/\/([a-zA-Z0-9\-.]+)(\/[^\s]*)?(\?[^\s]*)?', "", text).strip()

    return text

def make_directory_if_not_exists(directory: str):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)

def get_headers_based_on_table_of_contents(input_file: str):
    print('Extracting headers based on table of contents')

    [start_page, end_page] = get_config('TableOfContentsPageRange')
    new_section_regex = get_config('NewSectionRegex')

    headers = []
    with pdfplumber.open(input_file) as pdf:
        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text()
            page.flush_cache()
            
            if (i+1) >= start_page and (i+1) <= end_page:
                matches = re.findall(new_section_regex, page_text, re.MULTILINE)
                for match in matches:
                    headers.append(match.strip())
            elif (i+1) > end_page:
                break
    
    print(f'Found {len(headers)} headers in the table of contents')
    print(headers)
    return headers

def extract_text(): 
    input_file = get_config('InputPdfFile')
    output_directory = get_config('OutputFolder')
    new_section_regex = get_config('NewSectionRegex')

    make_directory_if_not_exists(output_directory)

    output_directory_path = Path(output_directory)

    should_normalize_text = get_config('NormalizeText')

    extract_based_on_table_of_contents = get_config('ExtractBasedOnTableOfContents')
    headers = None
    if extract_based_on_table_of_contents:
        headers = get_headers_based_on_table_of_contents(input_file)
        if len(headers) == 0:
            raise Exception('Unable to process file because no headers found in table of contents')
    
    print(f'Extracting text from "{input_file}" into "{output_directory}"')
    
    section_name = None
    header_name = None
    previous_section_name = None
    section_counter = 0
    with pdfplumber.open(input_file) as pdf:
        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text()
            tables = page.extract_table()
            page.flush_cache()
            
            if should_normalize_text:
                page_text = normalize_text(page_text)

            if extract_based_on_table_of_contents:
                current_header = find_header_in_page_text(headers, page_text)
                if current_header is None and section_name is None:
                    header_name = "Beginning"
                    section_name = f"{section_counter:02}. {header_name}"
                    section_counter+=1
                elif current_header is not None:
                    header_name = current_header.strip().replace("/", " or ").replace("\\", " or ")
                    previous_section_name = section_name
                    section_name = f'{section_counter:02} - {header_name}'
                    section_counter+=1
            else:
                new_section_match = re.search(new_section_regex, page_text, re.MULTILINE)
                if not new_section_match and not section_name:
                    header_name = "Beginning"
                    section_name = f"{section_counter:02}. {header_name}"
                    section_counter+=1
                elif new_section_match:
                    header_name = new_section_match.group().strip().replace("/", " or ").replace("\\", " or ")
                    previous_section_name = section_name
                    section_name = f'{section_counter:02} - {header_name}'
                    section_counter+=1

                    
            if not section_name:
                raise Exception(f'No section name is set! Page: {i+1}')
            
            if tables:
                ascii_table = tabulate(tables, headers="firstrow", tablefmt="grid")
                first_cell = tables[0][0]
                last_cell = tables[-1][-1]
                if not first_cell or not last_cell:
                    first_cell = ''
                    last_cell = ''
                else:
                    if '\n' in first_cell:
                        first_cell = first_cell.split('\n')[0]
                    if '\n' in last_cell:
                        last_cell = last_cell.split('\n')[-1]
                
                starting_index = page_text.index(first_cell)
                ending_index = page_text.index(last_cell) + len(last_cell)
                page_text = page_text[:starting_index] + '\n\n' + ascii_table + '\n\n' + page_text[ending_index:]

            stripped_text = page_text.strip()
            if not stripped_text.startswith(header_name) and header_name in stripped_text and previous_section_name:
                [before, after] = page_text.split(header_name)

                after = f'{header_name}{after}'
                write_page_text_to_file(output_directory_path, previous_section_name, before, f'{i+1}.txt')
                write_page_text_to_file(output_directory_path, section_name, after, f'{i+1}.txt')
            else:
                write_page_text_to_file(output_directory_path, section_name, page_text, f'{i+1}.txt')

def write_page_text_to_file(output_directory_path, section_name, page_text, file_name):
    sub_directory = output_directory_path / section_name
    make_directory_if_not_exists(sub_directory)

    print(f"Saving {file_name} to {sub_directory}")
    with open(sub_directory / file_name, 'w', encoding='utf8') as writer:
        writer.write(page_text)

def find_header_in_page_text(headers, page_text):
    # exclude table of contents
    if '.....' in page_text:
        return None
    
    page_lines = []
    for line in page_text.split('\n'):
        page_lines.append(line.strip())
    
    for header in headers:
        if header in page_lines:
            print(f'Page matched on header "{header}"')
            return header
        
    return None

def main():
    try:
        opts, _ = getopt.getopt(sys.argv[1:], "e", ['extract-text'])
    except getopt.GetoptError:
        print_help()
        sys.exit(2)
        
    if requesting_help(opts):
        print_help()
        sys.exit(0)        
    
    upload_ruling = get_flag(opts, '-e')
    if upload_ruling:
        extract_text()

def requesting_help(opts):
    help = next((o for o in opts if len(o) > 0 and o[0] == '-h'), None)
    return help != None

def print_help():
    print('You need help...')
    
def get_value(opts, parameter):
    value_arg = next((o for o in opts if len(o) > 0 and o[0] == parameter and len(o) == 2), None)
    return value_arg[1] if value_arg != None else None

def get_flag(opts, parameter):
    value_arg = next((o for o in opts if len(o) > 0 and o[0] == parameter), None)
    return value_arg != None

if __name__ == "__main__":
	sys.exit(main())