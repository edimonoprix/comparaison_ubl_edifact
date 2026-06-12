import os
import re
import xml.etree.ElementTree as ET
import xlsxwriter

# ================================================================
# MAPPING XML → EDI
# ================================================================

RULES = {
    "InvoiceNumber": ("BGM", "Numéro de facture"),
    "InvoiceIssueDate": ("DTM+137", "Date émission"),
    "ActualShipDate": ("DTM+11", "Date livraison"),
    "PurchaseOrderDate": ("DTM+171", "Date commande"),
    "BuyerOrderNumber": ("RFF+ON", "Num commande"),
    "RefNum": ("RFF+DQ", "Référence"),
    "InvoiceDueDate": ("DTM+13", "Date échéance"),
}

# ================================================================
# UTILS
# ================================================================

def local_name(tag):
    return tag.split("}")[-1]

def iso_to_yyyymmdd(s):
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", str(s))
    return f"{m.group(1)}{m.group(2)}{m.group(3)}" if m else s

def normalized_for_edi(tag, value):
    if "Date" in tag:
        return iso_to_yyyymmdd(value)
    return str(value)

def get_available_filename(base="compare.xlsx"):
    name, ext = os.path.splitext(base)
    i = 1
    candidate = base

    while os.path.exists(candidate):
        candidate = f"{name}_{i}{ext}"
        i += 1

    return candidate

# ================================================================
# EXTRACTION XML (corrigée ✅)
# ================================================================

def extract_xml_pairs(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    pairs = []
    seen = set()

    for elem in root.iter():
        tag = local_name(elem.tag)
        value = elem.text

        if value and value.strip():
            clean_value = value.strip()

            key = (tag, clean_value)

            # éviter doublons
            if key not in seen:
                seen.add(key)
                pairs.append(key)

    return pairs

# ================================================================
# EXTRACTION EDI
# ================================================================

def extract_edi_segments(edi_path):
    with open(edi_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    return [seg.strip() for seg in content.split("'") if seg.strip()]

# ================================================================
# MATCH XML → EDI
# ================================================================

def find_edi_segment(edi_segments, edi_key, value):
    code = edi_key.split("+")[0]

    # 1. priorité sur valeur
    for seg in edi_segments:
        if seg.startswith(code) and value in seg:
            return seg

    # 2. fallback
    for seg in edi_segments:
        if seg.startswith(code):
            return seg

    return ""

# ================================================================
# EXPORT EXCEL
# ================================================================

def write_excel(rows, edi_segments):
    output = get_available_filename("compare.xlsx")

    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Comparaison")

    bold = workbook.add_format({"bold": True})
    red = workbook.add_format({"font_color": "red"})

    # Headers
    worksheet.write(0, 0, "Champ XML", bold)
    worksheet.write(0, 1, "Valeur XML", bold)
    worksheet.write(0, 2, "Segment EDI", bold)

    row_index = 1

    for tag, value in rows:

        # mapping si existe
        if tag in RULES:
            edi_key, label = RULES[tag]
        else:
            edi_key, label = ("", tag)

        clean_value = normalized_for_edi(tag, value)

        if edi_key:
            edi_segment = find_edi_segment(edi_segments, edi_key, clean_value)
        else:
            edi_segment = ""

        worksheet.write(row_index, 0, label)
        worksheet.write(row_index, 1, value)
        worksheet.write(row_index, 2, edi_segment)

        row_index += 1

    workbook.close()

    return output

# ================================================================
# MAIN (test local)
# ================================================================

def main():
    xml_file = "input.xml"
    edi_file = "input.edi"

    edi_segments = extract_edi_segments(edi_file)
    rows = extract_xml_pairs(xml_file)

    print(f"Nombre de lignes XML : {len(rows)}")

    output = write_excel(rows, edi_segments)

    print(f"Fichier généré : {output}")

if __name__ == "__main__":
    main()
