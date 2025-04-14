import pdfplumber
import csv
import re
import os

# Define input PDF file and output CSV file
PDF_PATH = "Contracts/Fully Executed Contract_86615424C00017 (1).pdf"
OUTPUT_CSV = "output.csv"
DEBUG_LOG = "debug_log.txt"
DEBUG_MODE = True

COLUMNS = [
    "PR Number", "PR Title", "CLIN", "SLIN", "Title", "Quantity", "Estimated Unit Price ($)",
    "Unit Of Measure", "Amount ($)", "Amount Committed ($)", "Amount Reserved ($)", "Optional",
    "Not to Exceed", "Description", "POP Start Date", "POP End Date", "Group", "Line Item Type",
    "NSP", "Place of Performance"
]

VALID_UOM = {
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
    return ""

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
# Line Item Field Extractors
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
    cleaned = re.sub(r"\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})?$", "", full_text).strip()
    cleaned = re.sub(r"\b\d{1,4}\b(?=\s*$)", "", cleaned).strip()
    return cleaned

def extract_lineitem_quantity(match, is_slin):
    text_segment = match.group(2)
    tokens = text_segment.split()
    for token in reversed(tokens):
        if re.fullmatch(r"\d{1,4}", token):
            return token
    return ""

def extract_lineitem_unit_price(match, is_slin):
    text_segment = match.group(2)
    tokens = text_segment.split()
    prices = [t for t in tokens if re.fullmatch(r"\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})", t)]
    return prices[-1].replace("$", "") if len(prices) > 1 else ""

def extract_lineitem_amount(match):
    return match.group(3).replace("$", "")

def extract_lineitem_uom(match, is_slin):
    text_segment = match.group(2).split()
    for token in text_segment:
        if token in VALID_UOM:
            return token
    return ""

def extract_lineitem_amount_committed(match):
    return ""

def extract_lineitem_amount_reserved(match):
    return ""

# -----------------------------
# Line Item Extractor from Text
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
                    "Unit Of Measure": extract_lineitem_uom(match, is_slin),
                    "Amount Committed ($)": extract_lineitem_amount_committed(match),
                    "Amount Reserved ($)": extract_lineitem_amount_reserved(match),
                }
                line_items.append(item)
                log_debug("Line Item Extracted from Text:")
                for key, val in item.items():
                    log_debug(f"  {key:30}: {val}")
                log_debug("----------------------------------------")

    return line_items

# -----------------------------
# PDF Extraction and Orchestration
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

def write_to_csv(structured_data, line_items):
    reset_output_file()

    with open(OUTPUT_CSV, mode='w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=COLUMNS)
        writer.writeheader()

        if line_items:
            for item in line_items:
                row_data = structured_data.copy()
                row_data.update(item)
                writer.writerow(row_data)
                log_debug(f"CSV Row Written: {row_data}")
        else:
            writer.writerow(structured_data)
            log_debug(f"CSV Row Written (no line items): {structured_data}")

if __name__ == "__main__":
    if DEBUG_MODE and os.path.exists(DEBUG_LOG):
        os.remove(DEBUG_LOG)

    extracted_data = extract_pdf_text(PDF_PATH)
    structured_data, line_items = parse_pdf_data(extracted_data)
    write_to_csv(structured_data, line_items)
    print("Extraction complete! Data written to output.csv")