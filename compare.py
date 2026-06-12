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

# Traductions supplémentaires
TRANSLATION_FR = {
    "BuyerName": "Nom acheteur",
    "SellerName": "Nom vendeur",
    "Street": "Adresse",
    "City": "Ville",
    "PostalCode": "Code postal",
    "Country": "Pays",
    "Amount": "Montant",
    "Currency": "Devise",
    "Tax": "Taxe",
    "TotalAmount": "Montant total",
}

# ================================================================
# HELPERS
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
    i = 1

    while os.path.exists(candidate):
        candidate = f"{name}_{i}{ext}"
        i += 1

    return candidate

def get_french_label(tag):
    if tag in RULES:
        return RULES[tag][1]
    elif tag in TRANSLATION_FR:
        return TRANSLATION_FR[tag]
    else:
        return tag

# ================================================================
# EXTRACTION XML (complète + sans doublons)
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
# EXTRACTION EDI (TOUS LES SEGMENTS ✅)
# ================================================================

def extract_edi_segments(edi_path):
    with open(edi_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    return [seg.strip() for seg in content.split("'") if seg.strip()]

# ================================================================
# MATCH XML → EDI (amélioré)
# ================================================================

def find_edi_segment(edi_segments, edi_key, value):
    if not edi_key:
        return ""

    code = edi_key.split("+")[0]

    # segments correspondants
    candidates = [seg for seg in edi_segments if seg.startswith(code)]

    # priorité valeur exacte
    for seg in candidates:
        if value in seg:
            return seg

    # fallback
    if candidates:
        return candidates[0]

    return ""

# ================================================================
# EXPORT EXCEL (COMPLET ✅)
# ================================================================

def write_excel(rows, edi_segments):
    output = get_available_filename("compare.xlsx")

    workbook = xlsxwriter.Workbook(output)
    ws = workbook.add_worksheet("Comparaison")

    bold = workbook.add_format({"bold": True})
    red = workbook.add_format({"font_color": "red"})
    blue = workbook.add_format({"font_color": "blue"})

    # Headers
    ws.write(0, 0, "Ligne XML", bold)
    ws.write(0, 1, "Segment EDI", bold)
    ws.write(0, 2, "Fonction (FR)", bold)

    row_index = 1

    used_edi = set()

    # ====================================================
    # 1) XML → EDI
    # ====================================================
    for tag, value in rows:

        xml_display = f"<{tag}>{value}</{tag}>"
        label_fr = get_french_label(tag)

        edi_key = RULES[tag][0] if tag in RULES else ""
        clean_value = normalized_for_edi(tag, value)

        edi_segment = find_edi_segment(edi_segments, edi_key, clean_value)

        ws.write(row_index, 0, xml_display)

        if edi_segment:
            ws.write(row_index, 1, edi_segment)
            used_edi.add(edi_segment)
        else:
            ws.write(row_index, 1, "NON TROUVÉ", red)

        ws.write(row_index, 2, label_fr)

        row_index += 1

    # ====================================================
    # 2) TOUS LES SEGMENTS EDI NON UTILISÉS ✅
    # ====================================================
    for seg in edi_segments:
        if seg not in used_edi:
            ws.write(row_index, 0, "")
            ws.write(row_index, 1, seg, blue)
            ws.write(row_index, 2, "Segment sans correspondance XML")
            row_index += 1

    # Mise en forme
    ws.set_column(0, 0, 50)
    ws.set_column(1, 1, 70)
    ws.set_column(2, 2, 35)

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

    print(f"{len(rows)} lignes XML extraites")
    print(f"{len(edi_segments)} segments EDI trouvés")

    output = write_excel(rows, edi_segments)

    print(f"Fichier généré : {output}")

if __name__ == "__main__":
    main()
