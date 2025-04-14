import pdfplumber
import csv
import re
import os

# Folder path and output CSV file, script will loop through each file
PDF_FOLDER = "Contracts"
OUTPUT_CSV = "output.csv"
DEBUG_LOG = "debug_log.txt"
DEBUG_MODE = True

COLUMNS = [
    "Source File", "PR Number", "PR Title", "CLIN", "SLIN", "Title", "Quantity", "Estimated Unit Price ($)",
    "Unit", "Amount ($)", "Amount Committed ($)", "Amount Reserved ($)", "Optional",
    "Not to Exceed", "Description", "POP Start Date", "POP End Date", "Group", "Line Item Type",
    "NSP", "Place of Performance"
]

VALID_UNIT = {
    "AL","BG","BJ","OC","BO","BX","CT","CS","CH","CF","CO","CY","DA","DY","A8","EA","EC",
    "EX","FL","FT","1G","GA","GL","HR","FF","HS","HE","JO","JB","KW","KH","LH","LL","LF","LP",
    "LT","LD","NL","LO","LS","M0","PU","MG","MT","MN","MW","MH","MM","MB","BZ","MJ","MO","TN",
    "OT","PH","ZP","PR","PB","PL","P1","PI","LB","PJ","PE","PO","QT","Q1","RM","RE","RT","1O",
    "ST","SU","SV","SQ","SF","SY","YT","TD","TH","NT","TO","UN","WK","WM","YD","YR"
}

def log_debug(message):
    if DEBUG_MODE:
        with open(DEBUG_LOG, "a") as log_file:
            log_file.write(message + "\n")

def get_filename_from_path(path):
    return os.path.basename(path)

# -----------------------------
# Field Extractor Functions
# -----------------------------

def extract_pr_number(lines):
    for i, line in enumerate(lines):
        if "REQUISITION NUMBER" in line.upper():
            for j in range(1, 4):
                if i + j < len(lines):
                    candidate = lines[i + j].strip()
                    match = re.search(r"RCS-[A-Z0-9\-]+", candidate)
                    if match:
                        return match.group(0)
    return "N/A"

def extract_pr_title(lines):
    return "N/A"

TEXT_FIELD_EXTRACTORS = {
    "PR Number": extract_pr_number,
    "PR Title": extract_pr_title,
}

def parse_text_content(text):
    structured_fields = {}
    lines = text.split("\n")

    for field, extractor in TEXT_FIELD_EXTRACTORS.items():
        try:
            value = extractor(lines)
            if value:
                structured_fields[field] = value
                log_debug(f"Extracted {field}: {value}")
        except Exception as e:
            log_debug(f"Error extracting {field}: {e}")

    return structured_fields

# -----------------------------
# Placeholder Line Item Field Extractors
# -----------------------------

def extract_lineitem_optional(match):
    return "N/A"

def extract_lineitem_not_to_exceed(match):
    return "N/A"

def extract_lineitem_description(match):
    return "N/A"

def extract_lineitem_pop_start_date(match):
    return "N/A"

def extract_lineitem_pop_end_date(match):
    return "N/A"

def extract_lineitem_group(match):
    return "N/A"

def extract_lineitem_line_item_type(match):
    return "N/A"

def extract_lineitem_nsp(match):
    return "N/A"

def extract_lineitem_place_of_performance(match):
    return "N/A"

def extract_lineitem_amount_committed(match):
    return "N/A"

def extract_lineitem_amount_reserved(match):
    return "N/A"

# -----------------------------
# Line Item Extractors
# -----------------------------

def extract_lineitem_clin_or_slin(match):
    code = match.group(1)
    if re.match(r"\d{4}$", code):
        return code, "N/A"
    elif re.match(r"\d{4}[A-Z]+$", code):
        return code[:4], code
    return code, "N/A"

def extract_lineitem_title(match):
    full_text = match.group(2)
    tokens = full_text.split()

    # Reverse scan to find and remove Quantity + Unit pattern (e.g., "1 LO")
    cleaned_tokens = tokens.copy()
    for i in range(len(tokens) - 2, -1, -1):
        quantity_token = tokens[i]
        unit_token = tokens[i + 1]
        if re.fullmatch(r"\d{1,4}", quantity_token) and unit_token in VALID_UNIT:
            cleaned_tokens = tokens[:i]
            break

    return " ".join(cleaned_tokens).strip()

def extract_lineitem_quantity(match, is_slin):
    text_segment = match.group(2)
    tokens = text_segment.split()

    for i in range(len(tokens) - 1):
        token = tokens[i]
        next_token = tokens[i + 1] if i + 1 < len(tokens) else ""

        # Quantity must be directly followed by a known unit
        if re.fullmatch(r"\d{1,4}", token) and next_token in VALID_UNIT:
            return token

    return "N/A"

def extract_lineitem_unit_price(match, is_slin):
    text_segment = match.group(2)
    tokens = text_segment.split()
    prices = [t for t in tokens if re.fullmatch(r"\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})", t)]
    return prices[-1].replace("$", "") if len(prices) > 1 else "N/A"

def extract_lineitem_amount(match):
    return match.group(3).replace("$", "")

def extract_lineitem_unit(match, is_slin):
    text_segment = match.group(2).split()
    for token in text_segment:
        if token in VALID_UNIT:
            return token
    return "N/A"

# -----------------------------
# Line Item Parser
# -----------------------------

def parse_line_items_from_text(text):
    lines = text.split("\n")
    line_items = []
    capture = False
    pattern = re.compile(r"^(\d{4}[A-Z]{0,2})\s+(.*?)\s+(\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)$")

    for i, line in enumerate(lines):
        if "ITEM NO." in line or "SCHEDULE OF SUPPLIES/SERVICES" in line:
            capture = True
            continue

        if capture:
            if "Continued ..." in line:
                continue

            match = pattern.match(line.strip())
            if match:
                clin, slin = extract_lineitem_clin_or_slin(match)
                is_slin = slin != "N/A"
                item = {
                    "CLIN": clin,
                    "SLIN": slin,
                    "Title": extract_lineitem_title(match),
                    "Quantity": extract_lineitem_quantity(match, is_slin),
                    "Estimated Unit Price ($)": extract_lineitem_unit_price(match, is_slin),
                    "Amount ($)": extract_lineitem_amount(match),
                    "Unit": extract_lineitem_unit(match, is_slin),
                    "Amount Committed ($)": extract_lineitem_amount_committed(match),
                    "Amount Reserved ($)": extract_lineitem_amount_reserved(match),
                    "Optional": extract_lineitem_optional(match),
                    "Not to Exceed": extract_lineitem_not_to_exceed(match),
                    "Description": extract_lineitem_description(match),
                    "POP Start Date": extract_lineitem_pop_start_date(match),
                    "POP End Date": extract_lineitem_pop_end_date(match),
                    "Group": extract_lineitem_group(match),
                    "Line Item Type": extract_lineitem_line_item_type(match),
                    "NSP": extract_lineitem_nsp(match),
                    "Place of Performance": extract_lineitem_place_of_performance(match)
                }
                line_items.append(item)
                log_debug("Line Item Extracted from Text:")
                for key, val in item.items():
                    log_debug(f"  {key:30}: {val}")
                log_debug("----------------------------------------")

    return line_items

# -----------------------------
# PDF Processing Pipeline
# -----------------------------

def extract_pdf_text(pdf_path):
    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            log_debug(f"========== PAGE {page_num} ==========")
            text = page.extract_text()
            tables = page.extract_table()

            log_debug("Text Content:")
            log_debug(text if text else "[No text extracted]")
            log_debug("\nTable Content:")
            if tables:
                for t_row in tables:
                    log_debug(f"  {t_row}")
            else:
                log_debug("  [No table extracted]")

            if text:
                extracted_data.append(("text", text))
            if tables:
                extracted_data.append(("table", tables))

    return extracted_data

def parse_pdf_data(extracted_data):
    structured_data = {col: "" for col in COLUMNS}
    line_items = []

    for data_type, content in extracted_data:
        if data_type == "text":
            structured_data.update(parse_text_content(content))
            line_items.extend(parse_line_items_from_text(content))

    return structured_data, line_items

def reset_output_file():
    if os.path.exists(OUTPUT_CSV):
        os.remove(OUTPUT_CSV)

# Update write_to_csv to handle multiple files and prepend file name to each row
def write_to_csv(all_data):
    reset_output_file()

    with open(OUTPUT_CSV, mode='w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=COLUMNS)
        writer.writeheader()
        for data in all_data:
            for item in data['line_items']:
                row = {col: "" for col in COLUMNS}
                row.update(data['structured_data'])
                row.update(item)
                row["Source File"] = data['source_file']
                writer.writerow(row)
                log_debug(f"CSV Row Written: {row}")

# Main batch processor
if __name__ == "__main__":
    if DEBUG_MODE and os.path.exists(DEBUG_LOG):
        os.remove(DEBUG_LOG)

    all_results = []
    for filename in os.listdir(PDF_FOLDER):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(PDF_FOLDER, filename)
            extracted_data = extract_pdf_text(pdf_path)
            structured_data, line_items = parse_pdf_data(extracted_data)
            all_results.append({
                "source_file": filename,
                "structured_data": structured_data,
                "line_items": line_items
            })
    write_to_csv(all_results)
    print("Extraction complete! Data written to output.csv")