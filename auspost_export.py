"""
auspost_export.py — generates an AusPost order-import spreadsheet from
Shopify orders tagged 'packed-ready'.

Plugs into app.py via export_orders_to_xlsx(orders_data) which takes a list
of order dicts (already fetched from Shopify) and returns a path to the
generated .xlsx file.
"""

import re
import csv

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

# Flat default parcel dimensions in cm, applied to every row regardless of
# packaging type (satchel or own packaging). AusPost started rejecting bulk
# imports without these — adjust these three numbers if the real parcels run
# bigger/smaller than a folded rashguard/gi in a satchel.
ITEM_LENGTH_CM = 25
ITEM_WIDTH_CM = 25
ITEM_HEIGHT_CM = 5

# What AusPost should do if a parcel can't be delivered.
# AusPost only accepts: RETURN or ABANDONED
CANNOT_BE_DELIVERED = "RETURN"

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
    line_items (list), tags (list, lowercased), email (str — top-level order email,
    NOT shipping_address email, since Shopify's shipping_address has no email field),
    shipping_line_title (str — the actual shipping method the customer selected and
    paid for at checkout, e.g. 'Express Shipping' / 'International Express' — NOT
    the same thing as the confirmed-express/express-upgrade tags below)."""
    addr = order.get("shipping_address") or {}
    tags = order.get("tags", [])
    is_intl = order["is_international"]

    bucket_rows = group_line_items(order["line_items"])

    # Three separate ways an order can end up as Express, in priority order:
    # 1. Customer actually selected and paid for an Express-named shipping
    #    method at checkout (title varies by store — "Express Shipping",
    #    "International Express", etc. — so this is a substring match, not
    #    an exact one).
    # 2. Staff manually confirmed an upgrade via the pack-screen weight
    #    warning (confirmed-express tag).
    # 3. Staff or CS manually flagged an order for upgrade another way
    #    (express-upgrade tag).
    shipping_line_title = (order.get("shipping_line_title") or "").lower()
    customer_paid_express = "express" in shipping_line_title
    is_express = customer_paid_express or "confirmed-express" in tags or "express-upgrade" in tags
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
        addr.get("zip", ""), addr.get("phone", ""), order.get("email", ""),
        packaging, delivery_service, "Apparel",
        ITEM_LENGTH_CM, ITEM_WIDTH_CM, ITEM_HEIGHT_CM,
        TOTAL_WEIGHT_KG, "NO", "NO", "", CANNOT_BE_DELIVERED,
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
    Handles split-shipment orders by emitting 2 rows for those specifically.

    Despite the function name (kept for backwards compatibility with app.py's
    existing calls), this now writes a real .csv file — AusPost's bulk upload
    only accepts CSV, not .xlsx, so output_path should end in .csv."""
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)

        for order in orders:
            tags = order.get("tags", [])
            needs_split = "needs-split-shipment" in tags
            if needs_split:
                # two identical-ish rows = two separate parcels/labels for this order
                for _ in range(2):
                    writer.writerow(build_row(order))
            else:
                writer.writerow(build_row(order))

    return output_path
