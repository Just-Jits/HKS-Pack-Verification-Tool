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
    "phone": "61422695001",
    "email": "info@justjits.com",
}

TOTAL_WEIGHT_KG = 0.2          # flat default per parcel, single combined customs line
VALUE_PCT = 0.10                # declared customs value = 10% of retail (a 90% reduction)
HS_FALLBACK = "611030"          # Rash Guard — used if a product matches no bucket

# Three separate dimension profiles now, not one flat default for everyone:
#
# 1. Domestic default (small/XS satchel) — every domestic order unless the
#    single-item override below applies.
DOMESTIC_SATCHEL_LENGTH_CM = 30
DOMESTIC_SATCHEL_WIDTH_CM = 10
DOMESTIC_SATCHEL_HEIGHT_CM = 5
#
# 2. Domestic, confirmed single tiny item (xs-satchel-ok tag) — switches
#    packaging to Own Packaging with these smaller dimensions instead of
#    the AusPost satchel product.
DOMESTIC_XS_OWN_LENGTH_CM = 20
DOMESTIC_XS_OWN_WIDTH_CM = 20
DOMESTIC_XS_OWN_HEIGHT_CM = 2
#
# 3. International — always Own Packaging. Left unchanged from the original
#    flat default (not part of the domestic dimension rework).
INTL_LENGTH_CM = 25
INTL_WIDTH_CM = 25
INTL_HEIGHT_CM = 5

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
#
# IMPORTANT: we deliberately leave this BLANK rather than sending an explicit
# value. Per AusPost's policy (auspost.com.au/disruptions-and-updates/
# international-service-updates/global-customs-imports): "If you do not
# choose an option... it will automatically default to RECEIVER_PAYS for all
# destinations, except for shipments [where the destination country has
# mandated a different payer], [in which case] the appropriate option will
# be automatically pre-selected." Real export errors have shown explicit
# RECEIVER_PAYS gets rejected for the US and Spain specifically — and the
# set of mandated-payer countries is tied to evolving international customs
# rules (EU/Canada/Norway/UK changes are ongoing per the same policy page),
# so hardcoding a country whitelist here would need constant upkeep and will
# go stale. Leaving the field blank lets AusPost apply the correct default
# per destination automatically, with zero risk of rejection either way.
# On top of that, the two Sender-pays options aren't enabled for CSV bulk
# lodgement until 16 July 2026, so there's no valid explicit override we
# could send for the mandated-country cases even if we wanted to right now.
LANDED_COSTS_PAYER = ""


def normalize_au_phone(phone):
    """Converts an AU phone number to '61...' format (country code, no
    leading 0, no +, no spaces/dashes) — e.g. '0416 123 123' -> '61416123123'.

    Why: a number written as '0416123123' gets silently mangled to
    '416123123' the moment anyone opens the CSV in Excel/Sheets to check it
    — those apps auto-detect a leading-zero numeric string as a number, and
    numbers don't have leading zeros. Writing it as '61416123123' instead
    sidesteps the problem entirely: it's unambiguously a long digit string,
    not something a spreadsheet app will "helpfully" reformat.

    Only touches numbers that already look like AU numbers (start with 0 or
    already start with 61) — deliberately leaves other countries' numbers
    untouched, since e.g. a US customer's phone already carries its own
    correct country code and prepending 61 to it would break it."""
    if not phone:
        return phone
    digits = re.sub(r"[^\d]", "", phone)  # strip spaces, dashes, parens, +
    if digits.startswith("61") and len(digits) >= 11:
        return digits  # already in 61... format, just strip any punctuation
    if digits.startswith("0"):
        return "61" + digits[1:]  # 0416... -> 61416...
    return phone  # doesn't look like an AU number — leave as-is


def bucket_for(title):
    title_lower = title.lower()
    for bucket_name, keywords, hs_code in BUCKETS:
        for kw in keywords:
            if kw in title_lower:
                return bucket_name, hs_code
    return "Rash guard", HS_FALLBACK  # fallback bucket, per Ganesh's instruction


def round_to_90_cents(value):
    """Round down to the nearest dollar, then force the cents to .90 — matches
    the '10% of retail, ending in .90' rule Ganesh asked for."""
    whole = int(value)
    return whole + 0.90


def group_line_items(line_items):
    """Combines ALL line items in the order into a SINGLE customs-declaration
    line — one description, one HS code, quantity always 1, one combined
    value — rather than a separate line per product category.

    Splitting into multiple customs lines was the root cause of the
    over-declared-value issue: when entering these by hand, each of the
    (up to 4) lines was getting the full order value repeatedly instead of
    each getting its own share, multiplying the effective declared value.
    One combined line removes that failure mode entirely.

    Quantity is hardcoded to 1 regardless of how many physical units are in
    the parcel — this declares "one item" for customs purposes, per Ganesh's
    explicit instruction, not a literal per-unit count.

    The representative description/HS code is taken from whichever product
    category has the highest total retail value in the order (so a "3
    rashguards + 1 gi" order gets classified/described by the gi if the gi
    is worth more, even though rashguards outnumber it) — this is a
    simplification for a single-line declaration, not a precise per-item
    customs breakdown.
    """
    groups = {}  # bucket_name -> {qty, total_retail_value, hs_code}
    for li in line_items:
        bucket_name, hs_code = bucket_for(li["title"])
        if bucket_name not in groups:
            groups[bucket_name] = {"qty": 0, "value": 0.0, "hs_code": hs_code}
        groups[bucket_name]["qty"] += li["quantity"]
        groups[bucket_name]["value"] += li.get("price", 0) * li["quantity"]

    if not groups:
        return []

    total_retail_value = sum(g["value"] for g in groups.values())

    # Pick the highest-value category as the representative description/HS code
    dominant_bucket_name, dominant = max(groups.items(), key=lambda kv: kv[1]["value"])

    declared_value = round_to_90_cents(total_retail_value * VALUE_PCT)

    return [{
        "description": dominant_bucket_name[:40],
        "weight": TOTAL_WEIGHT_KG,
        "value": declared_value,
        "quantity": 1,
        "country_of_origin": "AU",
        "hs_tariff": dominant["hs_code"],
    }]


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
        item_length, item_width, item_height = INTL_LENGTH_CM, INTL_WIDTH_CM, INTL_HEIGHT_CM
    elif is_xs:
        # Packer confirmed this is a single item under 250g that fits a tiny
        # satchel — switch off the AusPost satchel product entirely and use
        # Own Packaging at the smaller confirmed dimensions.
        packaging = "OWN_PACKAGING"
        delivery_service = "EXP" if is_express else "PP"
        item_length, item_width, item_height = DOMESTIC_XS_OWN_LENGTH_CM, DOMESTIC_XS_OWN_WIDTH_CM, DOMESTIC_XS_OWN_HEIGHT_CM
    else:
        # Domestic default — this is OUR OWN packaging/satchels, not
        # AusPost's own "Small Satchel" product — at the standard dimensions.
        packaging = "OWN_PACKAGING"
        delivery_service = "EXP" if is_express else "PP"
        item_length, item_width, item_height = DOMESTIC_SATCHEL_LENGTH_CM, DOMESTIC_SATCHEL_WIDTH_CM, DOMESTIC_SATCHEL_HEIGHT_CM

    # AusPost rejects "Deliver To Business Name" over 40 characters. Rather
    # than silently truncating and losing info like "Attention: Brigitte",
    # cut the business name to 40 chars for that field but keep the full
    # original text in Delivery Instructions so packers/couriers still see it.
    #
    # NOTE: addr.get("company", "") is NOT enough on its own — Shopify's API
    # returns "company": null explicitly when there's no company name (the
    # key exists, its value is None), so .get()'s default never kicks in.
    # The trailing "or ''" catches that None and turns it into a safe string.
    raw_company = addr.get("company", "") or ""
    if len(raw_company) > 40:
        business_name = raw_company[:40]
        company_overflow_note = raw_company
    else:
        business_name = raw_company
        company_overflow_note = ""

    # Only normalize the AU leading-zero format for domestic orders —
    # international customers' numbers already carry their own correct
    # country code and shouldn't have 61 prepended on top of it.
    deliver_to_phone = normalize_au_phone(addr.get("phone", "")) if not is_intl else addr.get("phone", "")

    row = [
        SEND_FROM["name"], SEND_FROM["business_name"], SEND_FROM["address_line_1"],
        SEND_FROM["address_line_2"], SEND_FROM["address_line_3"], SEND_FROM["suburb"],
        SEND_FROM["state"], SEND_FROM["postcode"], SEND_FROM["phone"], SEND_FROM["email"],
        addr.get("name", ""), "", business_name, "",
        addr.get("country_code", "AU") if is_intl else "AU",
        addr.get("address1", ""), addr.get("address2", ""), "",
        addr.get("city", ""), addr.get("province_code", "") if not is_intl else addr.get("province", ""),
        addr.get("zip", ""), deliver_to_phone, order.get("email", ""),
        packaging, delivery_service, "Apparel",
        item_length, item_width, item_height,
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

    delivery_instructions = f"Business name (full): {company_overflow_note}" if company_overflow_note else ""

    row += ["YES", SEND_FROM["email"], f"Order {order['order_number']}", delivery_instructions, "",
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
