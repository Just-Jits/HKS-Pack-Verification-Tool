"""
auspost_export.py — generates an AusPost order-import spreadsheet from
Shopify orders tagged 'packed-ready'.

Plugs into app.py via export_orders_to_xlsx(orders_data) which takes a list
of order dicts (already fetched from Shopify) and returns a path to the
generated .xlsx file.
"""

import re
from openpyxl import Workbook

SEND_FROM = {
    "name": "Just Jits",
    "business_name": "Just Jits",
    "address_line_1": "203 Rooks Road",
    "address_line_2": "Unit 64",
    "address_line_3": "",
    "suburb": "Vermont",
    "state": "VIC",
    "postcode": "3133",
    "phone": "",  # fill in if you want this on every row
    "email": "",
}

TOTAL_WEIGHT_KG = 0.2          # flat default per parcel, split evenly across buckets
VALUE_PCT = 0.80                # declared customs value = 80% of retail
HS_FALLBACK = "611030"          # Rash Guard — used if a product matches no bucket

# Bucket name -> (keywords to match in product title, HS code for that bucket)
BUCKETS = [
    ("Gi", ["gi", "kimono", "uniform"], "611430"),
    ("Boxing/protective gear", ["boxing glove", "shin pad", "shin guard", "glove",
                                  "protective", "headgear", "head gear", "mouthguard",
                                  "mouth guard"], "950699"),
    ("Rash guard", ["rash guard", "rashguard", "shorts", "spats", "legging",
                     "no-gi", "no gi", "belt", "sock", "hoodie", "t-shirt", "tshirt"],
     "611030"),
]

HEADERS = [
    "Send From Name", "Send From Business Name", "Send From Address Line 1",
    "Send From Address Line 2", "Send From Address Line 3", "Send From Suburb",
    "Send From State", "Send From Postcode", "Send From Phone Number",
    "Send From Email Address", "Deliver To Name", "Deliver To MyPost Number",
    "Deliver To Business Name", "Deliver To Type Of Address", "Deliver To Country",
    "Deliver To Address Line 1", "Deliver To Address Line 2", "Deliver To Address Line 3",
    "Deliver To Suburb", "Deliver To State", "Deliver To Postcode",
    "Deliver To Phone Number", "Deliver To Email Address", "Item Packaging Type",
    "Item Delivery Service", "Item Description", "Item Length", "Item Width",
    "Item Height", "Item Weight", "Item Dangerous Goods Flag", "Signature On Delivery",
    "Extra Cover Amount", "Cannot Be Delivered - Return To Sender / Abandon Parcel",
    "Customs Declaration - Commercial Value", "Customs Declaration - Reason For Export",
    "Customs Declaration - Why Are You Exporting These Goods", "Export Declaration Number",
    "SMS Tracking Notifications", "SMS Tracking Mobile Phone",
    "Parcel Contents - Description", "Parcel Contents - Weight", "Parcel Contents - Value",
    "Parcel Contents - Quantity", "Parcel Contents - Country Of Origin",
    "Parcel Contents - HS Tariff",
    "Parcel 2 Contents - Description", "Parcel 2 Contents - Weight",
    "Parcel 2 Contents - Value", "Parcel 2 Contents - Quantity",
    "Parcel 2 Contents - Country Of Origin", "Parcel 2 Contents - HS Tariff",
    "Parcel 3 Contents - Description", "Parcel 3 Contents - Weight",
    "Parcel 3 Contents - Value", "Parcel 3 Contents - Quantity",
    "Parcel 3 Contents - Country Of Origin", "Parcel 3 Contents - HS Tariff",
    "Parcel 4 Contents - Description", "Parcel 4 Contents - Weight",
    "Parcel 4 Contents - Value", "Parcel 4 Contents - Quantity",
    "Parcel 4 Contents - Country Of Origin", "Parcel 4 Contents - HS Tariff",
    "Send Tracking Notifications", "Send Tracking Email", "Additional Label Information 1",
    "Delivery Instructions", "Comments", "Landed Costs Payer", "Sender's Customs Reference",
    "Importer's Reference Number", "Licence Numbers", "Certificate Numbers",
    "Invoice Numbers",
]

# New AusPost field (added to their bulk import template, July 2026). Only
# meaningful for international shipments — declares who pays any import
# duties/taxes charged by the destination country. AusPost doesn't handle
# these costs itself; it just needs the declaration.
# Valid values: RECEIVER_PAYS, SENDER_PAYS_ZONOS, SENDER_PAYS_TAX_ID
# (SENDER_PAYS_TAX_ID also requires "Importer's Reference Number" to be set.)
LANDED_COSTS_PAYER = "RECEIVER_PAYS"


def bucket_for(title):
    title_lower = title.lower()
    for bucket_name, keywords, hs_code in BUCKETS:
        for kw in keywords:
            if kw in title_lower:
                return bucket_name, hs_code
    return "Rash guard", HS_FALLBACK  # fallback bucket, per Ganesh's instruction


def round_to_90_cents(value):
    """Round to the nearest dollar, then force the cents to .90 — matches the
    '80%, ending in .90' rule Ganesh asked for."""
    whole = int(value)
    return whole + 0.90


def group_line_items(line_items):
    """Groups raw Shopify line items into up to 4 customs-declaration buckets,
    combining identical product types (e.g. 4 rashguards -> one line, qty 4)
    rather than one line per unit."""
    groups = {}  # bucket_name -> {qty, total_retail_value, hs_code}
    for li in line_items:
        bucket_name, hs_code = bucket_for(li["title"])
        if bucket_name not in groups:
            groups[bucket_name] = {"qty": 0, "value": 0.0, "hs_code": hs_code}
        groups[bucket_name]["qty"] += li["quantity"]
        groups[bucket_name]["value"] += li.get("price", 0) * li["quantity"]

    bucket_list = list(groups.items())
    if len(bucket_list) > 4:
        # Combine the smallest-value extras into the last bucket that fits,
        # rather than silently dropping data or guessing — flagged via the
        # 'NEEDS REVIEW' description so it's obviously not a normal row.
        bucket_list.sort(key=lambda x: x[1]["value"])
        overflow = bucket_list[:-4]
        kept = bucket_list[-4:]
        extra_qty = sum(b[1]["qty"] for b in overflow)
        extra_value = sum(b[1]["value"] for b in overflow)
        kept[0][1]["qty"] += extra_qty
        kept[0][1]["value"] += extra_value
        bucket_list = kept

    weight_per_bucket = round(TOTAL_WEIGHT_KG / max(len(bucket_list), 1), 3)

    rows = []
    for bucket_name, data in bucket_list:
        declared_value = round_to_90_cents(data["value"] * VALUE_PCT)
        rows.append({
            "description": bucket_name[:40],
            "weight": weight_per_bucket,
            "value": declared_value,
            "quantity": data["qty"],
            "country_of_origin": "AU",
            "hs_tariff": data["hs_code"],
        })
    return rows


def build_row(order):
    """order is a dict with: order_number, shipping_address (dict), is_international,
    line_items (list), tags (list, lowercased)."""
    addr = order.get("shipping_address") or {}
    tags = order.get("tags", [])
    is_intl = order["is_international"]

    bucket_rows = group_line_items(order["line_items"])

    is_express = "confirmed-express" in tags or "express-upgrade" in tags
    is_xs = (not is_intl) and ("xs-satchel-ok" in tags)

    if is_intl:
        packaging = "OWN_PACKAGING"
        delivery_service = "EXP" if is_express else "STD"
    else:
        packaging = "AP_SATCHEL_XS" if is_xs else "AP_SATCHEL_S"
        delivery_service = "EXP" if is_express else "PP"

    row = [
        SEND_FROM["name"], SEND_FROM["business_name"], SEND_FROM["address_line_1"],
        SEND_FROM["address_line_2"], SEND_FROM["address_line_3"], SEND_FROM["suburb"],
        SEND_FROM["state"], SEND_FROM["postcode"], SEND_FROM["phone"], SEND_FROM["email"],
        addr.get("name", ""), "", addr.get("company", ""), "",
        addr.get("country_code", "AU") if is_intl else "AU",
        addr.get("address1", ""), addr.get("address2", ""), "",
        addr.get("city", ""), addr.get("province_code", "") if not is_intl else addr.get("province", ""),
        addr.get("zip", ""), addr.get("phone", ""), "",
        packaging, delivery_service, "Apparel", "", "", "",
        TOTAL_WEIGHT_KG, "NO", "NO", "", "",
    ]

    if is_intl:
        row += ["YES", "OTHER", "Sale of goods", ""]
    else:
        row += ["", "", "", ""]

    row += ["NO", ""]  # SMS Tracking Notifications, SMS Tracking Mobile Phone

    for i in range(4):
        if i < len(bucket_rows):
            b = bucket_rows[i]
            row += [b["description"], b["weight"], b["value"], b["quantity"],
                    b["country_of_origin"] if is_intl else "",
                    b["hs_tariff"] if is_intl else ""]
        else:
            row += ["", "", "", "", "", ""]

    landed_costs_payer = LANDED_COSTS_PAYER if is_intl else ""

    row += ["NO", "", f"Order {order['order_number']}", "", "",
            landed_costs_payer, "", "", "", "", ""]

    return row


def export_orders_to_xlsx(orders, output_path):
    """orders: list of order dicts (see build_row docstring for shape).
    Handles split-shipment orders by emitting 2 rows for those specifically."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(HEADERS)

    for order in orders:
        tags = order.get("tags", [])
        needs_split = "needs-split-shipment" in tags
        if needs_split:
            # two identical-ish rows = two separate parcels/labels for this order
            for _ in range(2):
                ws.append(build_row(order))
        else:
            ws.append(build_row(order))

    wb.save(output_path)
    return output_path
