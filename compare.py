import os
import re
import xml.etree.ElementTree as ET
import xlsxwriter

# ===================================================
# RULES
# ===================================================

RULES = {
    "InvoiceNumber": ("BGM", "Numéro facture"),
    "InvoiceIssueDate": ("DTM+137", "Date facture"),
}

# ===================================================
# UTILS
# ===================================================

def local_name(tag):
    return tag.split("}")[-1]

def iso_to_yyyymmdd(s):
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", str(s))
    return f"{m.group(1)}{m.group(2)}{m.group(3)}" if m else s

def get_available_filename(base="compare.xlsx"):
    name, ext = os.path.splitext(base)
    i = 1
    candidate = base

    while os.path.exists(candidate):
        candidate = f"{name}_{i}{ext}"
        i += 1

    return candidate

# ===================================================
# XML
# ===================================================

def extract_xml_pairs(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    pairs = []

    for elem in root.iter():
        tag = local_name(elem.tag)

        if tag in RULES and elem.text:
            pairs.append((tag, elem.text))

    return pairs

# ===================================================
# EDI
# ===================================================

def extract_edi_segments(edi_path):
    with open(edi_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    segments = content.split("'")
    return [seg.strip() for seg in segments if seg.strip()]

# ===================================================
# MATCH
# ===================================================

def find_edi_segment(edi_segments, edi_key, value):
    code = edi_key.split("+")[0]

    for seg in edi_segments:
        if seg.startswith(code) and value in seg:
            return seg

    for seg in edi_segments:
        if seg.startswith(code):
            return seg

    return ""

# ===================================================
# EXCEL
# ===================================================

def write_excel(rows, edi_segments):
    output = get_available_filename("compare.xlsx")

    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "Champ")
    worksheet.write(0, 1, "Valeur XML")
    worksheet.write(0, 2, "Segment EDI")

    row_index = 1

    for tag, value in rows:
        edi_key, label = RULES.get(tag, ("", tag))

        edi_segment = find_edi_segment(edi_segments, edi_key, value)

        worksheet.write(row_index, 0, label)
        worksheet.write(row_index, 1, value)
        worksheet.write(row_index, 2, edi_segment)

        row_index += 1

    workbook.close()

    return output
