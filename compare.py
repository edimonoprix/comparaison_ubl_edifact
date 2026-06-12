import os
import re
import xml.etree.ElementTree as ET
import xlsxwriter

# ================================================================
# MAPPING PRO
# ================================================================

RULES = {
    "InvoiceNumber": ("BGM", "Numéro de facture"),
    "InvoiceIssueDate": ("DTM", "Date de facture"),
    "InvoiceDueDate": ("DTM", "Date d’échéance"),
    "BuyerOrderNumber": ("RFF", "Numéro de commande"),
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
    "UnitPriceValue": "Prix unitaire",
    "BuyerLineItemNum": "Ligne",
}

# ================================================================
# ✅ Traduction automatique FR
# ================================================================

def auto_translate(tag):
    words = re.findall('[A-Z][^A-Z]*', tag)

    dictionary = {
        "Name": "Nom",
        "Street": "Adresse",
        "City": "Ville",
        "Postal": "Postal",
        "Code": "Code",
        "Country": "Pays",
        "Currency": "Devise",
        "Amount": "Montant",
        "Quantity": "Quantité",
        "Price": "Prix",
        "Date": "Date",
        "Invoice": "Facture",
        "Buyer": "Acheteur",
        "Order": "Commande",
        "Ref": "Référence",
        "Identifier": "Identifiant",
        "Number": "Numéro",
        "Line": "Ligne",
        "Tax": "Taxe",
        "Type": "Type",
        "Description": "Description",
        "Term": "Condition",
        "Purpose": "Objet",
        "Item": "Article",
    }

    return " ".join([dictionary.get(w, w) for w in words])

def get_label(tag):
    if tag in RULES:
        return RULES[tag][1]
    if tag in TRANSLATION_FR:
        return TRANSLATION_FR[tag]
    return auto_translate(tag)

# ================================================================
# UTILS
# ================================================================

def local_name(tag):
    return tag.split("}")[-1]

def clean_text(v):
    return v.strip() if v else ""

def format_date(value):
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", value)
    return f"{m.group(1)}{m.group(2)}{m.group(3)}" if m else value

def get_filename(base="compare.xlsx"):
    i = 1
    name, ext = os.path.splitext(base)
    file = base
    while os.path.exists(file):
        file = f"{name}_{i}{ext}"
        i += 1
    return file

# ================================================================
# XML
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
# EDI
# ================================================================

def extract_edi(edi_file):
    with open(edi_file, "r", encoding="utf-8", errors="ignore") as f:
        txt = f.read()

    txt = txt.replace("\n", "")
    return [s.strip() for s in txt.split("'") if s.strip()]

# ================================================================
# ✅ MATCH METIER PROPRE
# ================================================================

def match_edi(segments, tag, value):

    if "Date" in tag:
        value = format_date(value)

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

        # ignorer segments techniques
        if seg.startswith(("UNB", "UNT", "UNZ")):
            continue

        # filtrer bon type
        if seg_type and not seg.startswith(seg_type):
            continue

        numbers = re.findall(r"\d+\.?\d*", seg)

        # match chiffres
        if value in numbers:
            return seg

        # match texte
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
# ✅ EXCEL FINAL (SANS OK/KO)
# ================================================================

def write_excel(xml_data, edi_segments):
    file = get_filename()

    wb = xlsxwriter.Workbook(file)
    ws = wb.add_worksheet("Comparaison")

    bold = wb.add_format({"bold": True})
    red = wb.add_format({"font_color": "red"})
    blue = wb.add_format({"font_color": "blue"})

    # headers
    ws.write(0, 0, "XML", bold)
    ws.write(0, 1, "Segment EDI", bold)
    ws.write(0, 2, "Fonction (FR)", bold)

    used = set()
    r = 1

    # XML → EDI
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

    # EDI restants
    for seg in edi_segments:
        if seg not in used:
            ws.write(r, 0, "")
            ws.write(r, 1, seg, blue)
            ws.write(r, 2, "Pas de correspondance dans le fichier xml")
            r += 1

    ws.set_column(0, 0, 50)
    ws.set_column(1, 1, 80)
    ws.set_column(2, 2, 30)

    wb.close()

    return file

# ================================================================
# ✅ COMPATIBILITE STREAMLIT
# ================================================================

extract_xml_pairs = extract_xml
extract_edi_segments = extract_edi
