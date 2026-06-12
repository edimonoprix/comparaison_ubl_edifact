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
# ✅ PARSING EDI AVANCÉ (SEGMENTS COMPLEXES)
# ================================================================

def extract_edi_segments(edi_path):
    with open(edi_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    # IMPORTANT : corrige les lignes coupées
    content = content.replace("\n", "")

    segments = [seg.strip() for seg in content.split("'") if seg.strip()]

    return segments

# ================================================================
# ✅ MATCH INTELLIGENT (gère segments complexes)
# ================================================================

def find_best_edi_match(edi_segments, tag, value):

    clean_value = normalized_for_edi(tag, value)

    matches = []

    for seg in edi_segments:

        # match exact (top priorité)
        if clean_value in seg:
            matches.append((seg, 3))

        # match partiel (texte inclus)
        elif clean_value.lower() in seg.lower():
            matches.append((seg, 2))

        # correspondance type segment (RULES)
        elif tag in RULES:
            code = RULES[tag][0].split("+")[0]
            if seg.startswith(code):
                matches.append((seg, 1))

    if matches:
        # prendre le meilleur score
        matches.sort(key=lambda x: x[1], reverse=True)
        return matches[0][0]

    return ""

# ================================================================
# ✅ EXPORT EXCEL COMPLET
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

    # ============================================
    # 1) XML → EDI
    # ============================================
    for tag, value in rows:

        xml_line = f"<{tag}>{value}</{tag}>"
        label = get_french_label(tag)

        edi_match = find_best_edi_match(edi_segments, tag, value)

        ws.write(row_index, 0, xml_line)

        if edi_match:
            ws.write(row_index, 1, edi_match)
            used_edi.add(edi_match)
        else:
            ws.write(row_index, 1, "NON TROUVÉ", red)

        ws.write(row_index, 2, label)

        row_index += 1

    # ============================================
    # 2) SEGMENTS EDI NON UTILISÉS ✅
    # ============================================
    for seg in edi_segments:
        if seg not in used_edi:
            ws.write(row_index, 0, "")
            ws.write(row_index, 1, seg, blue)
            ws.write(row_index, 2, "Segment EDI sans correspondance XML")
            row_index += 1

    # mise en forme
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
