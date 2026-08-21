"""
auspost_export.py — generates an AusPost order-import spreadsheet from
Shopify orders tagged 'packed-ready'.

Plugs into app.py via export_orders_to_xlsx(orders_data) which takes a list
of order dicts (already fetched from Shopify) and returns a path to the
generated .xlsx file.
"""

import re
import csv

try:
    import pykakasi
    _kks = pykakasi.kakasi()
    HAVE_PYKAKASI = True
except ImportError:
    HAVE_PYKAKASI = False

try:
    from unidecode import unidecode
    HAVE_UNIDECODE = True
except ImportError:
    HAVE_UNIDECODE = False

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

# Three separate dimension profiles now, not one flat default for everyone:
#
# 1. Domestic default — AusPost Small Satchel (their own preset-size
#    product, not our own packaging) unless the single-item override below
#    applies. Dimensions are set below but likely ignored by AusPost for a
#    named satchel product (it has its own fixed preset size in their
#    system) — left in place regardless since it's harmless either way and
#    some downstream tooling/reports may still read these columns.
DOMESTIC_SATCHEL_PACKAGING_CODE = "AP_SATCHEL_S"
DOMESTIC_SATCHEL_LENGTH_CM = 30
DOMESTIC_SATCHEL_WIDTH_CM = 10
DOMESTIC_SATCHEL_HEIGHT_CM = 5
#
# 2. Domestic, confirmed single tiny item (xs-satchel-ok tag) — switches
#    packaging to AusPost's Extra Small Satchel product instead of the
#    default Small Satchel.
DOMESTIC_XS_SATCHEL_PACKAGING_CODE = "AP_SATCHEL_XS"
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

# HS code lookup — keyed on the FIRST 6 DIGITS of whatever's actually
# assigned to the product in Shopify (Product > Inventory > HS Code), NOT
# guessed from the product title. Confirmed with Ganesh 2026-08 after
# finding Shopify's own HS codes were mostly missing or defaulted to the
# gi code across the whole catalog — this table only supplies the correct
# DESCRIPTION and US 10-digit HTSUS extension for a code Shopify already
# has recorded; it never invents a code that isn't there.
#
# us_hts10: the US import HTSUS code (administered by USITC) — required in
# full for any parcel where the destination country is the US. Different
# from a generic "10-digit HS code": it's specifically the US import
# extension, not Schedule B (US export) or any other country's extension.
HS_CODE_TABLE = {
    "620322": {"description": "Cotton martial arts uniform", "us_hts10": "6203221000"},
    "611430": {"description": "Cotton martial arts uniform", "us_hts10": "6203221000"},  # confirmed same as 620322 (Ganesh, 2026-08)
    "611030": {"description": "Synthetic knitted rash guard", "us_hts10": "6110303053"},
    "420321": {"description": "Leather sports gloves", "us_hts10": "4203218060"},
    "950699": {"description": "Synthetic sports/protective equipment", "us_hts10": "9506996080"},
    "611231": {"description": "Synthetic training shorts, men's", "us_hts10": "6112310010"},
    "611241": {"description": "Synthetic training shorts, women's", "us_hts10": "6112410010"},
    "640220": {"description": "Footwear thongs/sandals", "us_hts10": "6402200000"},

    # Added 2026-08 after the master-store lookup fix exposed several
    # product types (belts, bags, spats/tights, t-shirts) that had been
    # bulk-defaulted to the gi code (620322) in Shopify rather than left
    # blank — meaning they were resolving "successfully" but with the
    # WRONG description, not flagging for review. Descriptions and the
    # spats/tights HS6 (men's cut, applied universally per Ganesh's
    # instruction) confirmed with Ganesh 2026-08. US HTS-10 values below
    # are best-effort generic "Other" bucket codes, NOT broker-confirmed —
    # flagged for follow-up verification, same caveat applies to the
    # Belt/Bag entries' classification itself (genuinely ambiguous headings).
    "621710": {"description": "Cotton martial arts belt", "us_hts10": "6217109550"},
    "420292": {"description": "Duffle bag", "us_hts10": "4202923131"},
    "610343": {"description": "Polyester leggings and tights", "us_hts10": "6103431550"},
    "610910": {"description": "Cotton t-shirt", "us_hts10": "6109100027"},
    "610990": {"description": "Synthetic t-shirt", "us_hts10": "6109901007"},

    # Added 2026-08 while testing order #1488 (True Illusion, Executioner
    # Muay Thai Shorts). Ganesh gave this code directly (6203.43 - men's
    # synthetic trousers/breeches/shorts). Note: this is the SAME heading
    # family as 611231/611241 above, which are labelled "training shorts"
    # but are actually the 6112.3x/6112.4x swimwear subheadings, not
    # shorts - those two may be misclassified and worth a second look.
    "620343": {"description": "Synthetic athletic shorts, men's", "us_hts10": "6203439030"},
}


def hs6_key(raw_code):
    """Normalizes whatever's stored in Shopify's HS code field (may come
    with or without dots, may be 6 or 8+ digits depending on how it was
    entered) down to the plain 6-digit HS root for table lookup."""
    digits = re.sub(r"\D", "", raw_code or "")
    return digits[:6]


def resolve_hs_entry(raw_code, title):
    """Looks up the real Shopify-assigned HS code against our table. If the
    code isn't in the table (new product type we haven't catalogued yet, or
    Shopify has no code set at all), this does NOT silently guess — it
    returns a description that visibly flags the row for manual review, so
    a bad customs declaration can't slip out quietly the way the old
    title-keyword guesser did."""
    hs6 = hs6_key(raw_code)
    entry = HS_CODE_TABLE.get(hs6)
    if entry:
        return {
            "hs6": hs6,
            "description": entry["description"],
            "us_hts10": entry["us_hts10"],
        }
    flagged_desc = f"REVIEW HS CODE - {title}"[:40]
    return {
        "hs6": hs6 or "MISSING",
        "description": flagged_desc,
        "us_hts10": None,
    }

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


# Unicode blocks used to detect non-Latin script. Deliberately narrow —
# only the scripts Ganesh actually flagged (Japanese, Hebrew) trigger this,
# so e.g. accented European names pass through untouched.
JAPANESE_RANGES = [
    (0x3040, 0x309F),  # Hiragana
    (0x30A0, 0x30FF),  # Katakana
    (0x4E00, 0x9FFF),  # Kanji (CJK Unified Ideographs)
]
HEBREW_RANGE = (0x0590, 0x05FF)


def _contains_range(text, ranges):
    return any(any(lo <= ord(ch) <= hi for lo, hi in ranges) for ch in text)


def _is_japanese(text):
    return _contains_range(text, JAPANESE_RANGES)


def _is_hebrew(text):
    return _contains_range(text, [HEBREW_RANGE])


def transliterate_if_needed(text):
    """AusPost requires at least one Latin character in name/suburb/city
    fields. Rather than a placeholder character, this does a real
    best-effort ROMANIZATION (sounds-like, not translation) and prepends
    it ahead of the original script, e.g. '田中' -> 'Tanaka 田中'.

    Only touches text containing Japanese or Hebrew characters — anything
    else (including already-Latin text) passes through unchanged. Only
    ever called on name/city/suburb fields, per Ganesh's instruction —
    NEVER on street address lines, which stay exactly as the customer
    entered them to avoid any risk of mangling a real address."""
    if not text:
        return text

    if _is_japanese(text):
        if HAVE_PYKAKASI:
            try:
                romanized = "".join(item["hepburn"] for item in _kks.convert(text))
                romanized = romanized.strip().title()
                if romanized:
                    return f"{romanized} {text}"
            except Exception:
                pass  # fall through to unidecode/unchanged below
        if HAVE_UNIDECODE:
            romanized = unidecode(text).strip()
            if romanized:
                return f"{romanized} {text}"
        return text  # neither library available — leave unchanged rather than guess

    if _is_hebrew(text):
        if HAVE_UNIDECODE:
            romanized = unidecode(text).strip()
            if romanized:
                return f"{romanized} {text}"
        return text

    return text


def split_address_line1(address1, existing_line2=""):
    """AusPost hard-caps 'Deliver To Address Line 1' at 40 characters. If
    Shopify's address1 is longer, cut it at the last space at or before
    character 40 (never mid-word) and push the overflow onto the FRONT of
    address line 2 — ahead of whatever was already there (e.g. a unit
    number the customer put in address2), so nothing gets silently dropped.

    Returns (line1, line2). If line2 itself still exceeds 40 chars after
    combining (rare — needs a very long address1 AND a long existing
    line2), the excess is trimmed and returned separately as `overflow`
    so the caller can surface it in Delivery Instructions instead of
    losing it outright."""
    address1 = (address1 or "").strip()
    existing_line2 = (existing_line2 or "").strip()

    if len(address1) <= 40:
        return address1, existing_line2, ""

    cut = address1.rfind(" ", 0, 40)
    if cut <= 0:  # no space in the first 40 chars (one long word) — hard cut
        cut = 40
    line1 = address1[:cut].rstrip()
    carried_over = address1[cut:].strip()

    line2 = f"{carried_over} {existing_line2}".strip() if existing_line2 else carried_over

    if len(line2) <= 40:
        return line1, line2, ""

    # Still too long even combined — keep line2 at 40 and surface the rest
    # separately rather than silently truncating it away.
    return line1, line2[:40].rstrip(), line2[40:].strip()


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
    groups = {}  # hs6 (or 'MISSING') -> {qty, value, description, us_hts10}
    for li in line_items:
        entry = resolve_hs_entry(li.get("hs_code"), li["title"])
        key = entry["hs6"]
        if key not in groups:
            groups[key] = {
                "qty": 0, "value": 0.0,
                "description": entry["description"],
                "us_hts10": entry["us_hts10"],
            }
        groups[key]["qty"] += li["quantity"]
        groups[key]["value"] += li.get("price", 0) * li["quantity"]

    if not groups:
        return []

    total_retail_value = sum(g["value"] for g in groups.values())

    # Pick the highest-value HS code as the representative description/code
    # for this single combined customs line.
    dominant_key, dominant = max(groups.items(), key=lambda kv: kv[1]["value"])

    declared_value = round_to_90_cents(total_retail_value * VALUE_PCT)

    return [{
        "description": dominant["description"][:40],
        "weight": TOTAL_WEIGHT_KG,
        "value": declared_value,
        "quantity": 1,
        "country_of_origin": "AU",
        "hs_tariff": dominant_key,        # international default: 6-digit HS
        "us_hts10": dominant["us_hts10"], # used instead, only when destination is US
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
        # satchel — use AusPost's Extra Small Satchel product instead of
        # the default Small Satchel.
        packaging = DOMESTIC_XS_SATCHEL_PACKAGING_CODE
        delivery_service = "EXP" if is_express else "PP"
        item_length, item_width, item_height = DOMESTIC_XS_OWN_LENGTH_CM, DOMESTIC_XS_OWN_WIDTH_CM, DOMESTIC_XS_OWN_HEIGHT_CM
    else:
        # Domestic default — AusPost's own Small Satchel product.
        packaging = DOMESTIC_SATCHEL_PACKAGING_CODE
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

    # 40-char address line 1 cap — split cleanly rather than truncating.
    address_line_1, address_line_2, address_overflow = split_address_line1(
        addr.get("address1", ""), addr.get("address2", "")
    )

    # Non-Latin name/city — AusPost requires at least one Latin character.
    # Only name and city are touched (never street address lines).
    deliver_to_name = transliterate_if_needed(addr.get("name", ""))
    deliver_to_city = transliterate_if_needed(addr.get("city", ""))

    destination_country_code = addr.get("country_code", "AU") if is_intl else "AU"

    row = [
        SEND_FROM["name"], SEND_FROM["business_name"], SEND_FROM["address_line_1"],
        SEND_FROM["address_line_2"], SEND_FROM["address_line_3"], SEND_FROM["suburb"],
        SEND_FROM["state"], SEND_FROM["postcode"], SEND_FROM["phone"], SEND_FROM["email"],
        deliver_to_name, "", business_name, "",
        destination_country_code,
        address_line_1, address_line_2, "",
        deliver_to_city, addr.get("province_code", "") if not is_intl else addr.get("province", ""),
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
            # US shipments need the full 10-digit HTSUS code, not the
            # 6-digit international HS root. Falls back to the 6-digit
            # code if we don't have a US 10-digit match on file for this
            # HS code (also true for anything flagged MISSING/REVIEW).
            if is_intl and destination_country_code == "US" and b.get("us_hts10"):
                hs_tariff_out = b["us_hts10"]
            else:
                hs_tariff_out = b["hs_tariff"] if is_intl else ""
            row += [b["description"], b["weight"], b["value"], b["quantity"],
                    b["country_of_origin"] if is_intl else "",
                    hs_tariff_out]
        else:
            row += ["", "", "", "", "", ""]

    landed_costs_payer = LANDED_COSTS_PAYER if is_intl else ""

    instruction_parts = []
    if company_overflow_note:
        instruction_parts.append(f"Business name (full): {company_overflow_note}")
    if address_overflow:
        instruction_parts.append(f"Address (cont.): {address_overflow}")
    delivery_instructions = " | ".join(instruction_parts)

    # Personal customs code / tax ID — South Korea's Personal Customs
    # Code, Brazil's CPF/CNPJ, China's shipping credential, etc. Shopify
    # Markets decides which destination countries require this and
    # collects it at checkout automatically (order's SHIPPING-purpose
    # localizationExtensions); we just relay whatever value app.py's
    # attach_customs_codes() found. Only meaningful for international.
    importer_reference = order.get("importer_reference", "") if is_intl else ""

    row += ["YES", SEND_FROM["email"], f"Order {order['order_number']}", delivery_instructions, "",
            landed_costs_payer, "", importer_reference, "", "", ""]

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
