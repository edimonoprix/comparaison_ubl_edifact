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

TRANSLATION_FR = {
    "Name1": "Nom",
    "Street": "Adresse",
    "City": "Ville",
    "PostalCode": "Code postal",
    "CountryCoded": "Pays",
    "CurrencyCoded": "Devise",
    "BuyerLineItemNum": "Ligne acheteur",
    "ProductIdentifier": "Code produit",
    "ItemDescription": "Description article",
    "QuantityValue": "Quantité",
    "MonetaryAmount": "Montant",
    "UnitPriceValue": "Prix unitaire",
}

# ================================================================
# UTILS
# ================================================================

def local_name(tag):
    return tag.split("}")[-1]

def clean_text(text):
    if not text:
        return ""
    return text.strip().replace("\n", " ").replace("\r", "")

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

def get_french_label(tag):
    if tag in RULES:
        return RULES[tag][1]
    if tag in TRANSLATION_FR:
        return TRANSLATION_FR[tag]
    return tag

# ================================================================
# XML EXTRACTION
# ================================================================

def extract_xml_pairs(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    pairs = []
    seen = set()

    for elem in root.iter():
        tag = local_name(elem.tag)
        value = clean_text(elem.text)

        if value:
            key = (tag, value)
            if key not in seen:
                seen.add(key)
                pairs.append(key)

    return pairs

# ================================================================
# EDI PARSING (corrigé)
# ================================================================

def extract_edi_segments(edi_path):
    with open(edi_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    content = content.replace("\n", "")

    return [seg.strip() for seg in content.split("'") if seg.strip()]

# ================================================================
# MATCH INTELLIGENT (corrigé ✅)
# ================================================================

def find_best_edi_match(edi_segments, tag, value):

    clean_value = normalized_for_edi(tag, value)

    if not clean_value:
        return ""

    segment_map = {
        "ProductIdentifier": ["LIN"],
        "QuantityValue": ["QTY"],
        "MonetaryAmount": ["MOA"],
        "UnitPriceValue": ["PRI"],
        "City": ["NAD"],
        "Name1": ["NAD"],
        "Street": ["NAD"],
        "PostalCode": ["NAD"],
    }

    allowed_segments = segment_map.get(tag, None)

    best_match = ""

    for seg in edi_segments:

        # filtre type
        if allowed_segments:
            if not any(seg.startswith(code) for code in allowed_segments):
                continue

        # 🔥 match EXACT uniquement (corrige ton problème)
        if clean_value == re.findall(r"[0-9.]+", seg)[-1] if re.findall(r"[0-9.]+", seg) else False:
            return seg

        # fallback raisonnable
        if clean_value in seg:
            best_match = seg

    return best_match

# ================================================================
# ✅ HIGHLIGHT XML VALUE (ROUGE)
# ================================================================

def write_xml_highlight(ws, row, col, tag, value, fmt_red):
    open_tag = f"<{tag}>"
    close_tag = f"</{tag}>"

    ws.write_rich_string(
        row, col,
        open_tag,
        fmt_red, value,
        close_tag
    )

# ================================================================
# EXPORT EXCEL FINAL
# ================================================================

def write_excel(rows, edi_segments):
    output = get_available_filename("compare.xlsx")

    workbook = xlsxwriter.Workbook(output)
    ws = workbook.add_worksheet("Comparaison")

    bold = workbook.add_format({"bold": True})
    red = workbook.add_format({"font_color": "red"})
    blue = workbook.add_format({"font_color": "blue"})

    # headers
    ws.write(0, 0, "Ligne XML", bold)
    ws.write(0, 1, "Segment EDI", bold)
    ws.write(0, 2, "Fonction (FR)", bold)

    row_index = 1
    used_edi = set()

    # ==================================================
    # 1) XML → EDI
    # ==================================================
    for tag, value in rows:

        label = get_french_label(tag)

        edi_match = find_best_edi_match(edi_segments, tag, value)

        # ✅ XML avec valeur en rouge
        write_xml_highlight(ws, row_index, 0, tag, value, red)

        if edi_match:
            ws.write(row_index, 1, edi_match)
            used_edi.add(edi_match)
        else:
            ws.write(row_index, 1, "NON TROUVÉ", red)

        ws.write(row_index, 2, label)

        row_index += 1

    # ==================================================
    # 2) EDI NON UTILISÉS
    # ==================================================
    for seg in edi_segments:
        if seg not in used_edi:
            ws.write(row_index, 0, "")
            ws.write(row_index, 1, seg, blue)
            ws.write(row_index, 2, "Segment EDI sans correspondance XML")
            row_index += 1

    ws.set_column(0, 0, 50)
    ws.set_column(1, 1, 80)
    ws.set_column(2, 2, 35)

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

    print(f"{len(rows)} lignes XML")
    print(f"{len(edi_segments)} segments EDI")

    output = write_excel(rows, edi_segments)

    print(f"Fichier généré : {output}")

if __name__ == "__main__":
    main()
