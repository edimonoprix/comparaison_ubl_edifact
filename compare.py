import os
import re
import xml.etree.ElementTree as ET
import xlsxwriter

# ================================================================
# MAPPING XML → EDI + LIBELLE FR
# ================================================================

RULES = {
    "InvoiceNumber": ("BGM", "Numéro de facture"),
    "InvoiceIssueDate": ("DTM", "Date facture"),
    "InvoiceDueDate": ("DTM", "Date échéance"),
    "BuyerOrderNumber": ("RFF", "Commande"),
    "RefNum": ("RFF", "Référence"),
}

TRANSLATION_FR = {
    "Name1": "Nom",
    "Street": "Adresse",
    "City": "Ville",
    "PostalCode": "Code postal",
    "CountryCoded": "Pays",
    "CurrencyCoded": "Devise",
    "ProductIdentifier": "Produit",
    "QuantityValue": "Quantité",
    "MonetaryAmount": "Montant",
    "UnitPriceValue": "Prix",
    "BuyerLineItemNum": "Ligne",
}

# ================================================================
# UTILS
# ================================================================

def local_name(tag):
    return tag.split("}")[-1]

def clean_text(v):
    return v.strip() if v else ""

def iso_to_yyyymmdd(s):
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", str(s))
    return f"{m.group(1)}{m.group(2)}{m.group(3)}" if m else s

def normalize(tag, value):
    if "Date" in tag:
        return iso_to_yyyymmdd(value)
    return value

def get_label(tag):
    if tag in RULES:
        return RULES[tag][1]
    if tag in TRANSLATION_FR:
        return TRANSLATION_FR[tag]
    return tag

def get_filename(base="compare.xlsx"):
    i = 1
    name, ext = os.path.splitext(base)
    file = base
    while os.path.exists(file):
        file = f"{name}_{i}{ext}"
        i += 1
    return file

# ================================================================
# EXTRACTION XML
# ================================================================

def extract_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    seen = set()
    data = []

    for el in root.iter():
        tag = local_name(el.tag)
        val = clean_text(el.text)

        if val and (tag, val) not in seen:
            seen.add((tag, val))
            data.append((tag, val))

    return data

# ================================================================
# EXTRACTION EDI
# ================================================================

def extract_edi(edi_file):
    with open(edi_file, "r", encoding="utf-8", errors="ignore") as f:
        txt = f.read()

    txt = txt.replace("\n", "")
    return [s.strip() for s in txt.split("'") if s.strip()]

# ================================================================
# ✅ MATCH STRICT CORRIGÉ
# ================================================================

def match_edi(segments, tag, value):

    value = normalize(tag, value)

    if not value:
        return ""

    segment_types = {
        "ProductIdentifier": "LIN",
        "QuantityValue": "QTY",
        "MonetaryAmount": "MOA",
        "UnitPriceValue": "PRI",
        "City": "NAD",
        "Name1": "NAD",
        "Street": "NAD",
        "PostalCode": "NAD",
        "CountryCoded": "NAD",
        "CurrencyCoded": "CUX",
        "BuyerLineItemNum": "LIN",
    }

    seg_type = segment_types.get(tag)

    for seg in segments:

        # 🚫 bloquer segments inutiles
        if seg.startswith(("UNB", "UNT", "UNZ")):
            continue

        # ✅ filtrer type
        if seg_type and not seg.startswith(seg_type):
            continue

        # ✅ récupérer valeurs numériques
        numbers = re.findall(r"\d+\.?\d*", seg)

        # ✅ match strict chiffre
        if value in numbers:
            return seg

        # ✅ match texte (ville, nom…)
        if value.upper() in seg.upper():
            return seg

    return ""

# ================================================================
# XML EN ROUGE
# ================================================================

def write_xml(ws, row, col, tag, val, red):
    ws.write_rich_string(
        row, col,
        f"<{tag}>",
        red, val,
        f"</{tag}>"
    )

# ================================================================
# EXPORT EXCEL
# ================================================================

def write_excel(xml_data, edi_segments):
    file = get_filename()

    wb = xlsxwriter.Workbook(file)
    ws = wb.add_worksheet("Comparaison")

    bold = wb.add_format({"bold": True})
    red = wb.add_format({"font_color": "red"})
    blue = wb.add_format({"font_color": "blue"})

    # Headers
    ws.write(0, 0, "XML", bold)
    ws.write(0, 1, "EDI", bold)
    ws.write(0, 2, "Fonction", bold)

    used = set()
    r = 1

    # ====================================================
    # XML → EDI
    # ====================================================
    for tag, val in xml_data:

        label = get_label(tag)
        edi = match_edi(edi_segments, tag, val)

        write_xml(ws, r, 0, tag, val, red)

        if edi:
            ws.write(r, 1, edi)
            used.add(edi)
        else:
            ws.write(r, 1, "NON TROUVÉ", red)

        ws.write(r, 2, label)
        r += 1

    # ====================================================
    # EDI NON UTILISÉS
    # ====================================================
    for seg in edi_segments:
        if seg not in used:
            ws.write(r, 0, "")
            ws.write(r, 1, seg, blue)
            ws.write(r, 2, "EDI non mappé")
            r += 1

    ws.set_column(0, 0, 50)
    ws.set_column(1, 1, 80)
    ws.set_column(2, 2, 30)

    wb.close()

    return file

# ================================================================
# ✅ COMPATIBILITÉ AVEC TON APP.PY
# ================================================================

# IMPORTANT → évite ton erreur !
extract_xml_pairs = extract_xml
extract_edi_segments = extract_edi
