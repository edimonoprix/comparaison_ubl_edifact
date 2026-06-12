import os
import re
import xml.etree.ElementTree as ET
import xlsxwriter

# ================================================================
# MAPPING XML → EDI (FR)
# ================================================================

RULES = {
    "InvoiceNumber": ("BGM", "Numéro de facture"),
    "InvoiceIssueDate": ("DTM+137", "Date d’émission"),
    "ActualShipDate": ("DTM+11", "Date de livraison"),
    "PurchaseOrderDate": ("DTM+171", "Date de commande"),
    "BuyerOrderNumber": ("RFF+ON", "Numéro de commande"),
    "RefNum": ("RFF+DQ", "Référence"),
    "InvoiceDueDate": ("DTM+13", "Date d’échéance"),
}

# ================================================================
# HELPERS (repris de ton code)
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
    candidate = base
    idx = 1

    while os.path.exists(candidate):
        candidate = f"{name}_{idx}{ext}"
        idx += 1

    return candidate

# ================================================================
# 1) EXTRACTION XML (corrigée mais fidèle)
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

            if key not in seen:
                seen.add(key)
                pairs.append((tag, clean_value))

    return pairs

# ================================================================
# 2) EXTRACTION EDI (identique logique d’origine)
# ================================================================

def extract_edi_segments(edi_path):
    with open(edi_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    return [seg.strip() + "'" for seg in content.split("'") if seg.strip()]

# ================================================================
# 3) MATCH XML → EDI (repris de ta logique)
# ================================================================

def find_edi_segment(edi_segments, edi_key, value):
    if not edi_key:
        return ""

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
# 4) ÉCRITURE EXCEL (TA VERSION + AMÉLIORATIONS)
# ================================================================

def write_excel(rows, edi_segments):
    output = get_available_filename("compare.xlsx")

    workbook = xlsxwriter.Workbook(output)
    ws = workbook.add_worksheet("Comparaison")

    bold = workbook.add_format({"bold": True})
    red = workbook.add_format({"font_color": "red"})

    # headers
    ws.write(0, 0, "Ligne XML", bold)
    ws.write(0, 1, "Segment EDI correspondant", bold)
    ws.write(0, 2, "Fonction (FR)", bold)

    row_index = 1

    for tag, value in rows:

        # === XML affichage ===
        xml_display = f"<{tag}>{value}</{tag}>"

        # === mapping ===
        if tag in RULES:
            edi_key, label_fr = RULES[tag]
        else:
            edi_key, label_fr = ("", tag)

        clean_value = normalized_for_edi(tag, value)

        # === recherche EDI ===
        edi_segment = find_edi_segment(edi_segments, edi_key, clean_value)

        # === écriture ===
        ws.write(row_index, 0, xml_display)

        if edi_key and not edi_segment:
            ws.write(row_index, 1, "NON TROUVÉ", red)
        else:
            ws.write(row_index, 1, edi_segment)

        ws.write(row_index, 2, label_fr)

        row_index += 1

    # largeur colonnes (pro)
    ws.set_column(0, 0, 45)
    ws.set_column(1, 1, 50)
    ws.set_column(2, 2, 30)

    workbook.close()

    return output

# ================================================================
# MAIN
# ================================================================

def main():
    xml_file = "input.xml"
    edi_file = "input.edi"

    edi_segments = extract_edi_segments(edi_file)
    rows = extract_xml_pairs(xml_file)

    print(f"{len(rows)} lignes XML extraites")

    output = write_excel(rows, edi_segments)

    print(f"Fichier généré : {output}")

if __name__ == "__main__":
    main()
