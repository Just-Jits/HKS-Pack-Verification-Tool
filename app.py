import os
import re
import json
import time
import hmac
import hashlib
import tempfile
import requests
import pdfplumber
from flask import Flask, request, redirect, render_template, jsonify, session, send_file
from urllib.parse import urlencode
from auspost_export import export_orders_to_xlsx

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me-in-railway-env")

CLIENT_ID = os.environ.get("SHOPIFY_CLIENT_ID")  # fallback/default app
CLIENT_SECRET = os.environ.get("SHOPIFY_CLIENT_SECRET")  # fallback/default app
APP_URL = os.environ["APP_URL"]  # e.g. https://pack-verify-tool-production.up.railway.app
SCOPES = ",".join([
    "read_orders", "write_orders",
    "read_products", "write_products",
    "read_merchant_managed_fulfillment_orders", "write_merchant_managed_fulfillment_orders",
    "read_locations",
    "read_draft_orders", "write_draft_orders",
    "read_inventory", "write_inventory",
])
API_VERSION = "2026-04"

# Shopify only allows ONE custom app to cover multiple stores if they're in
# the same Plus organization. Since these are 5 separate standalone stores,
# each one needs its OWN custom app (own Client ID + Secret). Rather than
# hardcoding one pair, store a JSON mapping of shop -> {client_id, client_secret}
# in the SHOPIFY_APPS_JSON env var, e.g.:
#   {"hooks-jiu-jitsu.myshopify.com": {"client_id": "...", "client_secret": "..."},
#    "justjits.myshopify.com": {"client_id": "...", "client_secret": "..."}}
try:
    SHOPIFY_APPS = json.loads(os.environ.get("SHOPIFY_APPS_JSON", "{}"))
except json.JSONDecodeError:
    SHOPIFY_APPS = {}


def get_app_credentials(shop):
    """Returns (client_id, client_secret) for a given shop — checks the
    per-shop mapping first, falls back to the single default app credentials
    for backward compatibility with the original Hooks setup."""
    if shop in SHOPIFY_APPS:
        return SHOPIFY_APPS[shop]["client_id"], SHOPIFY_APPS[shop]["client_secret"]
    return CLIENT_ID, CLIENT_SECRET

TOKENS_FILE = "/data/tokens.json"  # mount a Railway volume at /data for this to persist


LOGS_FILE = "/data/error_log.json"


def log_error(context, detail):
    try:
        logs = []
        if os.path.exists(LOGS_FILE):
            with open(LOGS_FILE, "r") as f:
                logs = json.load(f)
        logs.append({
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "context": context,
            "detail": str(detail)[:2000],  # cap length so one giant error doesn't bloat the file
        })
        logs = logs[-200:]  # keep last 200 entries only
        os.makedirs(os.path.dirname(LOGS_FILE), exist_ok=True)
        with open(LOGS_FILE, "w") as f:
            json.dump(logs, f, indent=2)
    except Exception:
        pass  # logging itself should never crash the app


from datetime import timedelta
from functools import wraps

app.permanent_session_lifetime = timedelta(minutes=30)

try:
    PACK_USERS = json.loads(os.environ.get("PACK_TOOL_USERS_JSON", "{}"))
except json.JSONDecodeError:
    PACK_USERS = {}


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in_user"):
            return redirect(f"/login?next={request.path}")
        session.permanent = True  # refresh the 30-min sliding timeout on activity
        return f(*args, **kwargs)
    return wrapper


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        password = request.form.get("password", "")
        if PACK_USERS.get(name) == password:
            session.permanent = True
            session["logged_in_user"] = name
            return redirect(request.args.get("next", "/scan"))
        error = "Wrong name or password"
    return f"""
    <html><body style="background:#111;color:#fff;font-family:sans-serif;
    display:flex;align-items:center;justify-content:center;height:100vh;">
    <form method="post" style="background:#1d1d1d;padding:30px;border-radius:12px;width:280px;">
      <h2>📦 Pack Verify Login</h2>
      {f'<p style="color:#ff6b6b;">{error}</p>' if error else ''}
      <input name="name" placeholder="Name" style="width:100%;padding:10px;margin-bottom:10px;
        background:#333;color:#fff;border:none;border-radius:6px;" autofocus>
      <input name="password" type="password" placeholder="Password" style="width:100%;padding:10px;
        margin-bottom:10px;background:#333;color:#fff;border:none;border-radius:6px;">
      <button type="submit" style="width:100%;padding:10px;background:#2ecc71;color:#000;
        border:none;border-radius:6px;font-weight:bold;">Log in</button>
    </form>
    </body></html>
    """


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# ---------- token storage ----------

def load_tokens():
    if not os.path.exists(TOKENS_FILE):
        return {}
    with open(TOKENS_FILE, "r") as f:
        return json.load(f)


def save_tokens(tokens):
    os.makedirs(os.path.dirname(TOKENS_FILE), exist_ok=True)
    with open(TOKENS_FILE, "w") as f:
        json.dump(tokens, f, indent=2)


def get_token_for_shop(shop):
    tokens = load_tokens()
    return tokens.get(shop, {}).get("access_token")


def all_connected_shops():
    return list(load_tokens().keys())


# ---------- OAuth: install flow ----------

@app.route("/install")
def install():
    shop = request.args.get("shop")
    if not shop:
        return "Missing ?shop=your-store.myshopify.com in the URL", 400

    client_id, _ = get_app_credentials(shop)
    if not client_id:
        return f"No app configured for {shop} — add it to SHOPIFY_APPS_JSON in Railway", 400

    state = hashlib.sha256(os.urandom(16)).hexdigest()
    session["state"] = state
    session["shop"] = shop

    params = {
        "client_id": client_id,
        "scope": SCOPES,
        "redirect_uri": f"{APP_URL}/auth/callback",
        "state": state,
    }
    auth_url = f"https://{shop}/admin/oauth/authorize?{urlencode(params)}"
    return redirect(auth_url)


@app.route("/auth/callback")
def auth_callback():
    shop = request.args.get("shop")
    code = request.args.get("code")
    state = request.args.get("state")

    if state != session.get("state"):
        return "Invalid state — please restart install", 400

    client_id, client_secret = get_app_credentials(shop)
    if not client_id:
        return f"No app configured for {shop}", 400

    # exchange code for permanent-ish access token (Authorization Code Grant ->
    # this returns a long-lived token for live stores, unlike client_credentials)
    resp = requests.post(
        f"https://{shop}/admin/oauth/access_token",
        json={
            "client_id": client_id,
            "client_secret": client_secret,
            "code": code,
        },
    )
    if resp.status_code != 200:
        return f"Token exchange failed: {resp.text}", 400

    data = resp.json()
    tokens = load_tokens()
    tokens[shop] = {
        "access_token": data["access_token"],
        "scope": data.get("scope"),
        "installed_at": int(time.time()),
    }
    save_tokens(tokens)

    return f"""
    <h2>✅ Installed on {shop}</h2>
    <p>You can close this tab. Go to <a href="{APP_URL}/scan">the scan screen</a> to start packing.</p>
    """


# ---------- Shopify Admin API helper ----------

def shopify_get(shop, path, params=None):
    token = get_token_for_shop(shop)
    if not token:
        return None
    url = f"https://{shop}/admin/api/{API_VERSION}/{path}"
    r = requests.get(url, headers={"X-Shopify-Access-Token": token}, params=params or {})
    if r.status_code != 200:
        return None
    return r.json()


def shopify_get_all_pages(shop, path, params, results_key):
    """Like shopify_get, but follows Shopify's Link-header cursor pagination
    to fetch every page instead of just the first. Needed because relying
    on a single request silently truncates results once a store has more
    orders than fit on one page — and unlike the tag-search approach this
    replaces, page_info cursor pagination has no indexing delay, so it's
    safe to trust immediately after a tag was just applied."""
    token = get_token_for_shop(shop)
    if not token:
        return []
    all_results = []
    url = f"https://{shop}/admin/api/{API_VERSION}/{path}"
    next_params = dict(params or {})
    while url:
        r = requests.get(url, headers={"X-Shopify-Access-Token": token}, params=next_params)
        if r.status_code != 200:
            break
        data = r.json()
        all_results.extend(data.get(results_key, []))
        # Parse the Link header for a rel="next" cursor URL — this is how
        # Shopify's REST API paginates now (since_id/page params are gone).
        next_params = None  # cursor URL already carries all query params
        url = None
        link_header = r.headers.get("Link", "")
        for part in link_header.split(","):
            if 'rel="next"' in part:
                url = part[part.find("<") + 1:part.find(">")]
                break
    return all_results


def shopify_post(shop, path, payload):
    token = get_token_for_shop(shop)
    url = f"https://{shop}/admin/api/{API_VERSION}/{path}"
    r = requests.post(url, headers={"X-Shopify-Access-Token": token}, json=payload)
    return r.status_code, (r.json() if r.text else {})


def shopify_put(shop, path, payload):
    token = get_token_for_shop(shop)
    url = f"https://{shop}/admin/api/{API_VERSION}/{path}"
    r = requests.put(url, headers={"X-Shopify-Access-Token": token}, json=payload)
    return r.status_code, (r.json() if r.text else {})


def shopify_graphql(shop, query, variables=None):
    """Generic GraphQL POST helper — same auth pattern as the REST helpers
    above, just hitting graphql.json instead. Returns the parsed JSON body
    (which may itself contain a top-level "errors" array on failure — check
    for that separately from the HTTP status), or None on a transport-level
    failure (bad token, network error, non-200)."""
    token = get_token_for_shop(shop)
    if not token:
        return None
    url = f"https://{shop}/admin/api/{API_VERSION}/graphql.json"
    try:
        r = requests.post(
            url,
            headers={"X-Shopify-Access-Token": token, "Content-Type": "application/json"},
            json={"query": query, "variables": variables or {}},
            timeout=20,
        )
    except requests.exceptions.RequestException as e:
        log_error("shopify_graphql", f"shop={shop} | {e}")
        return None
    if r.status_code != 200:
        log_error("shopify_graphql", f"shop={shop} | HTTP {r.status_code}: {r.text[:500]}")
        return None
    return r.json()


WEIGHT_TABLE = [
    # (keywords to match in product title, lowercase), weight_kg
    (["gi"], 1.7),
    (["rashguard", "rash guard"], 0.3),
    (["shorts"], 0.3),
    (["hoodie"], 1.0),
    (["gear bag", "gearbag"], 1.0),
    (["belt"], 0.3),
]
DEFAULT_ITEM_WEIGHT_KG = 0.3  # fallback for anything unmatched


def estimate_item_weight(title):
    title_lower = title.lower()
    for keywords, weight in WEIGHT_TABLE:
        if any(kw in title_lower for kw in keywords):
            return weight
    return DEFAULT_ITEM_WEIGHT_KG


# Dedicated, more precise weight table just for the XS-satchel check below.
# Deliberately separate from WEIGHT_TABLE above — that one is a coarse
# shipping-weight estimate used for the Express/split-shipment checks, which
# can tolerate rough numbers. This one needs an actual garment-weight cutoff,
# since it's specifically deciding "does this fit under 250g".
#
# Only the per-unit weight is stored — max quantity under 250g is COMPUTED
# from it (250 // weight_g), not hand-typed. Hand-typing it previously led to
# every category being capped at 1 "to be safe", even featherweight items
# like key chains (20g, actually fits 12) and mouthguards (30g, fits 8) —
# computing it removes that whole class of silently-wrong number.
XS_SATCHEL_LIMIT_G = 250
XS_SATCHEL_KIDS_KEYWORDS = ["kids", "kid ", "youth"]
XS_SATCHEL_KIDS_ITEM_KEYWORDS = ["rashguard", "rash guard", "shorts", "spats",
                                  "legging", "compression"]
# (keywords, weight_grams) — checked in order, first match wins
XS_SATCHEL_ITEM_TABLE = [
    (["finger tape"], 120),
    (["rashguard", "rash guard", "shorts", "spats", "legging", "compression",
      "muay thai shorts", "mma shorts"], 180),
    (["belt"], 180),
    (["hand wrap"], 80),
    (["mouthguard", "mouth guard"], 30),
    (["lace converter"], 20),
    (["key chain", "keychain", "key ring", "keyring"], 20),
    (["air freshener"], 30),
]


def xs_satchel_item_lookup(title):
    """Returns (weight_grams, max_qty_under_250g) for a product title, or
    None if it doesn't match any XS-satchel-eligible category at all.
    max_qty is always computed from weight, never hand-typed."""
    title_lower = title.lower()
    # Kids'/youth rash guards & shorts are noticeably lighter than adult —
    # check this first so "Kids Rash Guard" doesn't fall through to the
    # heavier generic "rash guard" entry below.
    is_kids = any(kw in title_lower for kw in XS_SATCHEL_KIDS_KEYWORDS)
    if is_kids and any(kw in title_lower for kw in XS_SATCHEL_KIDS_ITEM_KEYWORDS):
        weight_g = 120
        return (weight_g, XS_SATCHEL_LIMIT_G // weight_g)
    for keywords, weight_g in XS_SATCHEL_ITEM_TABLE:
        if any(kw in title_lower for kw in keywords):
            return (weight_g, XS_SATCHEL_LIMIT_G // weight_g)
    return None


def estimate_order_weight(line_items):
    total = 0.0
    for li in line_items:
        total += estimate_item_weight(li["title"]) * li["quantity"]
    return round(total, 2)


def lookup_live_barcode(shop, sku):
    """REST's /variants.json?sku= filter is unreliable (doesn't actually
    filter), so use GraphQL's productVariants(query: "sku:...") instead,
    which properly filters server-side."""
    if not sku:
        return None
    token = get_token_for_shop(shop)
    if not token:
        return None
    url = f"https://{shop}/admin/api/{API_VERSION}/graphql.json"
    query = """
    query getVariantBySku($q: String!) {
      productVariants(first: 1, query: $q) {
        edges { node { barcode sku } }
      }
    }
    """
    # Shopify search syntax requires exact-match quoting for SKUs that may
    # contain special characters like hyphens.
    safe_sku = sku.replace('"', '\\"')
    variables = {"q": f'sku:"{safe_sku}"'}
    try:
        r = requests.post(
            url,
            headers={"X-Shopify-Access-Token": token, "Content-Type": "application/json"},
            json={"query": query, "variables": variables},
            timeout=15,
        )
        if r.status_code != 200:
            return None
        edges = r.json().get("data", {}).get("productVariants", {}).get("edges", [])
        if edges:
            return edges[0]["node"].get("barcode")
    except requests.exceptions.RequestException:
        pass
    return None


# ---------- core: find order across all connected stores ----------

def find_order_by_number(order_number):
    """order_number = plain digits scanned off the packing slip barcode, e.g. '9511'"""
    for shop in all_connected_shops():
        data = shopify_get(shop, "orders.json", params={"name": f"#{order_number}", "status": "any"})
        if data and data.get("orders"):
            order = data["orders"][0]
            return shop, order
    return None, None


# ---------- fulfill-from-PDF: extract tracking numbers, push to Shopify ----------

# Domestic labels (Parcel Post / Express Post) print the trackable number as
# plain text right next to "AP Article Id:" — but that text runs vertically
# up the side of the label, so pdfplumber extracts it as "AP\nArticle\nId:\n
# <digits>" (newlines, not spaces) rather than one line. \s* covers both.
TRACKING_DOMESTIC_RE = re.compile(r"AP\s*Article\s*Id:\s*(\d{15,30})")

# International CN23 labels print the tracking number in UPU format: two
# letters, 9 digits (sometimes grouped in 3s with spaces), then "AU" — e.g.
# "EJ 352 962 815 AU" or "LK279520986AU". Spacing is inconsistent between
# label variants, so the regex tolerates optional spaces between groups and
# the groups get joined back together with no spaces afterward.
TRACKING_INTL_RE = re.compile(r"\b([A-Z]{2})\s?(\d{3})\s?(\d{3})\s?(\d{3})\s?(AU)\b")


def find_tracking_matches(text):
    """Returns [(start_index, tracking_number), ...] for every tracking
    number on a page, in document order. A page-wide single .search() only
    ever finds the FIRST match — but several AusPost export layouts tile
    2-4 separate parcels' labels onto one physical page (Express Post pairs,
    Parcel Post 2x2 grids), so a page can legitimately contain multiple
    distinct tracking numbers that all need capturing, not just the first."""
    matches = []
    for m in TRACKING_DOMESTIC_RE.finditer(text):
        matches.append((m.start(), m.group(1)))
    for m in TRACKING_INTL_RE.finditer(text):
        matches.append((m.start(), "".join(m.groups())))
    matches.sort(key=lambda item: item[0])
    return matches

# Every label variant (domestic + international) prints the Shopify order
# number as plain text somewhere on the page — usually "Order #9452" or
# "Ref: Order #9547". Some AusPost export templates instead only carry it
# under the "Invoice No(s)" field as "Invoice No(s) #9452" with no separate
# "Order" text anywhere on the label — ORDER_NUMBER_RE catches the common
# case, ORDER_NUMBER_INVOICE_RE is the fallback for that template variant.
ORDER_NUMBER_RE = re.compile(r"Order[\s#-]*#?\s*(\d+)")
ORDER_NUMBER_INVOICE_RE = re.compile(r"Invoice No\(s\)\s*#\s*(\d+)")


def find_order_number_in_text(text):
    match = ORDER_NUMBER_RE.search(text)
    if match:
        return match.group(1)
    match = ORDER_NUMBER_INVOICE_RE.search(text)
    if match:
        return match.group(1)
    return None


def find_order_number_matches(text):
    """Like find_tracking_matches, but for order numbers: returns every
    occurrence with its position, since a page can hold more than one
    order's worth of text (multi-parcel domestic pages)."""
    matches = []
    for m in ORDER_NUMBER_RE.finditer(text):
        matches.append((m.start(), m.group(1)))
    for m in ORDER_NUMBER_INVOICE_RE.finditer(text):
        matches.append((m.start(), m.group(1)))
    matches.sort(key=lambda item: item[0])
    return matches


def extract_tracking_pairs_from_pdf(file_stream):
    """Parses an AusPost label-batch PDF and returns a list of
    {order_number, tracking_number, carrier} dicts — one per unique parcel.

    Handles two layout quirks seen in real exports:
    1. International CN23 shipments print 2-4 near-identical pages per
       order (Attach To Item / Customs Copy / Sender's Copy) — deduping by
       tracking_number (not order_number) handles that safely, since a
       single order is never legitimately split across two different
       tracking numbers within one PDF export.
    2. Domestic Express/Parcel Post layouts tile 2-4 SEPARATE parcels onto
       one physical page, and — critically — the order number appears
       AFTER the tracking number in reading order on those labels, while
       on CN23 labels the order number appears BEFORE the tracking number.
       A single "look forward from tracking to next tracking" block
       assumption only covers one of those two directions, so instead
       every order-number occurrence on the page is matched to its
       CLOSEST tracking number by text position — this works regardless
       of which side of the tracking number the order text falls on, and
       naturally isolates each parcel's own pair even when several are
       tiled on one page."""
    pairs = {}
    with pdfplumber.open(file_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            tracking_matches = find_tracking_matches(text)
            order_matches = find_order_number_matches(text)
            if not tracking_matches or not order_matches:
                continue

            for pos, tracking_number in tracking_matches:
                nearest_order = min(order_matches, key=lambda om: abs(om[0] - pos))
                order_number = nearest_order[1]

                pairs[tracking_number] = {
                    "order_number": order_number,
                    "tracking_number": tracking_number,
                    "carrier": "Australia Post",
                }
    return list(pairs.values())


FULFILLMENT_ORDERS_QUERY = """
query getFulfillmentOrders($orderId: ID!) {
  order(id: $orderId) {
    id
    name
    fulfillmentOrders(first: 10) {
      edges {
        node {
          id
          status
        }
      }
    }
  }
}
"""

FULFILLMENT_CREATE_MUTATION = """
mutation fulfillmentCreateV2($fulfillment: FulfillmentV2Input!) {
  fulfillmentCreateV2(fulfillment: $fulfillment) {
    fulfillment {
      id
      status
    }
    userErrors {
      field
      message
    }
  }
}
"""


def fulfill_order_with_tracking(shop, rest_order_id, tracking_number, carrier="Australia Post"):
    """Fulfills every currently-open fulfillment order on this Shopify order
    and attaches the given tracking number. Returns (ok: bool, message: str).

    NOTE: this fulfills ALL open fulfillment orders on the order, not a
    specific subset of line items. That's correct for the normal case (one
    order = one parcel = one label = one tracking number), but if an order
    is ever legitimately split across two separate shipments (e.g. partial
    stock from two locations), this would attach the same tracking number
    to both halves. Worth a manual check for any order flagged with a
    split-shipment tag before trusting this blindly on those."""
    gid = f"gid://shopify/Order/{rest_order_id}"
    result = shopify_graphql(shop, FULFILLMENT_ORDERS_QUERY, {"orderId": gid})
    if not result or result.get("errors"):
        return False, f"Could not fetch fulfillment orders: {result.get('errors') if result else 'no response — see /logs'}"

    order_data = (result.get("data") or {}).get("order")
    if not order_data:
        return False, "Order not found via GraphQL — token/scope issue or wrong store"

    fo_edges = order_data.get("fulfillmentOrders", {}).get("edges", [])
    open_fo_ids = [
        edge["node"]["id"] for edge in fo_edges
        if edge["node"]["status"] in ("OPEN", "IN_PROGRESS", "SCHEDULED")
    ]
    if not open_fo_ids:
        return False, "No open fulfillment orders — likely already fulfilled, cancelled, or on hold"

    fulfillment_input = {
        "lineItemsByFulfillmentOrder": [{"fulfillmentOrderId": fo_id} for fo_id in open_fo_ids],
        "trackingInfo": {"number": tracking_number, "company": carrier},
        "notifyCustomer": True,
    }
    result = shopify_graphql(shop, FULFILLMENT_CREATE_MUTATION, {"fulfillment": fulfillment_input})
    if not result or result.get("errors"):
        return False, f"GraphQL error: {result.get('errors') if result else 'no response — see /logs'}"

    payload = (result.get("data") or {}).get("fulfillmentCreateV2", {})
    user_errors = payload.get("userErrors", [])
    if user_errors:
        return False, "; ".join(e["message"] for e in user_errors)

    fulfillment = payload.get("fulfillment")
    if not fulfillment:
        return False, "No fulfillment returned and no error — unexpected response, check /logs"

    return True, f"Fulfilled (status: {fulfillment.get('status')})"


@app.route("/")
def home():
    shops = all_connected_shops()
    return render_template("home.html", shops=shops, app_url=APP_URL)


def extract_export_line_items(order):
    """Filters one order's raw Shopify line_items down to the {title,
    quantity, price} shape the export needs, skipping removed/swapped
    items via current_quantity. Pulled out of process_order_for_export so
    it can also be called per-order when combining a merged pair."""
    line_items = []
    for li in order.get("line_items", []):
        # Skip items removed/swapped via a Shopify order edit — quantity
        # still reflects the original order, current_quantity reflects
        # what's actually left after the edit. Without this, a removed
        # item still ends up on the customs declaration / weight calc.
        # Explicitly check for None (Shopify sending current_quantity:
        # null rather than omitting the key) rather than using "or",
        # since a legitimate current_quantity of 0 is falsy too and
        # "or" would wrongly fall back to the original quantity for
        # exactly the removed-item case this is meant to catch.
        current_qty = li.get("current_quantity", li["quantity"])
        if current_qty is None:
            current_qty = li["quantity"]
        if current_qty == 0:
            continue
        line_items.append({
            "title": li["title"],
            "quantity": current_qty,
            "price": float(li.get("price", 0)),
        })
    return line_items


def process_order_for_export(shop, order, merge_partner_order=None):
    """Converts one raw Shopify order dict into the format the AusPost
    export needs. Pulled out as its own function so both the normal
    packed-ready flow and the re-export-by-exception flow can share it
    rather than duplicating this logic.

    merge_partner_order: if this order is shipping together with another
    as ONE physical parcel (combined shipment), pass the partner's raw
    order dict here — their line items get combined into a single export
    row (one weight, one combined declared value via the existing
    highest-value-item logic in group_line_items), and both order IDs get
    returned so the caller can mark BOTH as exported/fulfilled together."""
    shipping_address = order.get("shipping_address") or {}
    is_international = shipping_address.get("country_code", "AU") not in ("AU", "")
    tags = [t.strip().lower() for t in (order.get("tags") or "").split(",")]

    line_items = extract_export_line_items(order)
    order_number_label = order["name"]
    extra_order_ids = []

    if merge_partner_order:
        line_items += extract_export_line_items(merge_partner_order)
        order_number_label = f"{order['name']} + {merge_partner_order['name']}"
        extra_order_ids = [merge_partner_order["id"]]

    shipping_lines = order.get("shipping_lines") or []
    shipping_line_title = shipping_lines[0].get("title", "") if shipping_lines else ""

    return {
        "order_id": order["id"],
        "extra_order_ids": extra_order_ids,
        "order_number": order_number_label,
        "shop": shop,
        "shipping_address": shipping_address,
        "is_international": is_international,
        "line_items": line_items,
        "tags": tags,
        "email": order.get("email", ""),
        "shipping_line_title": shipping_line_title,
    }


def fetch_packed_ready_orders(shop, single_order_id=None):
    """Pulls full order data needed for the AusPost export — either every
    packed-ready order not yet exported, or just one specific order.

    IMPORTANT: this does NOT search Shopify for tag:packed-ready. Shopify's
    tag filtering (REST's undocumented tag= param, and GraphQL's tag: search)
    both run against a search index that lags behind real-time by anywhere
    from seconds to a couple of minutes after a tag is applied. Packing a
    batch of orders and immediately exporting hits that lag directly — the
    orders tagged in the last minute or two before export are invisible to
    a tag search even though the tag is genuinely on the order, which is
    exactly what caused only ~15 of 50 packed orders to show up in one run.

    Instead: pull orders by fulfillment status (an indexed field with no
    search-lag behaviour) and read each order's actual `tags` field
    directly — Shopify populates that in real time with no propagation
    delay, since it's just reading the record rather than searching for it.
    Paginated via shopify_get_all_pages so nothing silently caps out either.
    """
    results = []
    if single_order_id:
        data = shopify_get(shop, f"orders/{single_order_id}.json")
        raw_orders = [data["order"]] if data else []
    else:
        # Packed-ready orders are, by definition, not fully fulfilled yet —
        # covers both "nothing fulfilled" and the split-fulfillment case
        # (e.g. one line already shipped from the US warehouse, rest still
        # needs to go from Vermont) without ever touching tag search.
        unfulfilled = shopify_get_all_pages(shop, "orders.json", {
            "status": "any", "fulfillment_status": "unfulfilled", "limit": 250,
        }, "orders")
        partial = shopify_get_all_pages(shop, "orders.json", {
            "status": "any", "fulfillment_status": "partial", "limit": 250,
        }, "orders")
        seen_ids = set()
        raw_orders = []
        for order in unfulfilled + partial:
            if order["id"] not in seen_ids:
                seen_ids.add(order["id"])
                raw_orders.append(order)

    processed_order_ids = set()
    for order in raw_orders:
        if order["id"] in processed_order_ids:
            continue  # already included as the partner half of a merged pair
        tags = [t.strip().lower() for t in (order.get("tags") or "").split(",")]
        if single_order_id is None:
            if "packed-ready" not in tags or "auspost-exported" in tags:
                continue

        merge_partner_number = find_merge_partner_number(tags)
        merge_partner_order = None
        if merge_partner_number:
            merge_partner_order = find_order_by_number_in_shop(shop, merge_partner_number)

        results.append(process_order_for_export(shop, order, merge_partner_order=merge_partner_order))
        processed_order_ids.add(order["id"])
        if merge_partner_order:
            processed_order_ids.add(merge_partner_order["id"])
    return results


def find_order_by_number_any_status(order_number):
    """Like find_order_by_number, but doesn't care whether the order is
    packed-ready or already auspost-exported — used for re-export by
    exception, where the order has usually already been exported once
    (that's the whole reason it needs re-exporting)."""
    for shop in all_connected_shops():
        data = shopify_get(shop, "orders.json", params={"name": f"#{order_number}", "status": "any"})
        if data and data.get("orders"):
            return shop, data["orders"][0]
    return None, None


def add_order_tag(shop, order_id, tag):
    """Generic single-tag add, used by the merge-shipment feature. Distinct
    from mark_orders_exported (which is specifically for auspost-exported)
    since merge tags need to be added to two DIFFERENT orders with two
    DIFFERENT tag values (each order gets a tag pointing at the OTHER
    order's number), not the same tag applied in bulk."""
    current = shopify_get(shop, f"orders/{order_id}.json")
    if not current:
        return False
    existing_tags = [t.strip() for t in (current["order"].get("tags") or "").split(",") if t.strip()]
    if not any(t.lower() == tag.lower() for t in existing_tags):
        existing_tags.append(tag)
        shopify_put(shop, f"orders/{order_id}.json", {
            "order": {"id": order_id, "tags": ", ".join(existing_tags)}
        })
    return True


MERGE_TAG_RE = re.compile(r"^merged-with-(\d+)$")


def find_merge_partner_number(tags):
    """tags: lowercased tag list off an order. Returns the partner order
    number (str) if this order carries a 'merged-with-XXXX' tag, else None.
    Reading the sibling's number straight out of the tag itself — rather
    than searching Shopify for 'which order has a matching group tag' —
    deliberately avoids depending on Shopify's tag SEARCH index at all,
    which is the exact search-index-lag class of bug fixed earlier this
    session for the packed-ready fetch. This only ever needs a direct
    order-number lookup, which is immediate and not search-based."""
    for t in tags:
        m = MERGE_TAG_RE.match(t)
        if m:
            return m.group(1)
    return None


def find_order_by_number_in_shop(shop, order_number):
    """Like find_order_by_number, but scoped to one specific shop rather
    than searching every connected store — merges should only ever happen
    between two orders in the SAME store."""
    data = shopify_get(shop, "orders.json", params={"name": f"#{order_number}", "status": "any"})
    if data and data.get("orders"):
        return data["orders"][0]
    return None


def unmark_order_exported(shop, order_id):
    """Removes the auspost-exported tag from one order, if present. Opposite
    of mark_orders_exported — used so a re-exported order doesn't keep its
    stale 'already exported' tag pointing at the broken CSV it was on before."""
    current = shopify_get(shop, f"orders/{order_id}.json")
    if not current:
        return
    existing_tags = [t.strip() for t in (current["order"].get("tags") or "").split(",") if t.strip()]
    if any(t.lower() == "auspost-exported" for t in existing_tags):
        existing_tags = [t for t in existing_tags if t.lower() != "auspost-exported"]
        shopify_put(shop, f"orders/{order_id}.json", {
            "order": {"id": order_id, "tags": ", ".join(existing_tags)}
        })


def mark_orders_exported(shop, order_ids):
    for order_id in order_ids:
        current = shopify_get(shop, f"orders/{order_id}.json")
        if not current:
            continue
        existing_tags = [t.strip() for t in (current["order"].get("tags") or "").split(",") if t.strip()]
        if "auspost-exported" not in existing_tags:
            existing_tags.append("auspost-exported")
            shopify_put(shop, f"orders/{order_id}.json", {
                "order": {"id": order_id, "tags": ", ".join(existing_tags)}
            })


# Tracks which orders were included in the most recent /api/export_batch run,
# so "re-export last batch" can redo exactly that set without anyone needing
# to type out order numbers by hand. Written to a temp file rather than kept
# in memory, so it survives across requests even if Railway runs multiple
# workers or restarts the process between the export and the re-export.
LAST_BATCH_EXPORT_PATH = os.path.join(tempfile.gettempdir(), "last_batch_export.json")


def save_last_batch_export(orders):
    """orders: the same list of dicts fetch_packed_ready_orders returns —
    just need shop/order_id/order_number out of each."""
    record = {
        "exported_at": time.time(),
        "orders": [
            {"shop": o["shop"], "order_id": o["order_id"], "order_number": o["order_number"]}
            for o in orders
        ],
    }
    with open(LAST_BATCH_EXPORT_PATH, "w") as f:
        json.dump(record, f)


def load_last_batch_export():
    """Returns the saved record, or None if no batch has been exported yet
    (or the temp file got cleared, e.g. by a Railway restart)."""
    if not os.path.exists(LAST_BATCH_EXPORT_PATH):
        return None
    try:
        with open(LAST_BATCH_EXPORT_PATH) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


@app.route("/api/export_batch")
@login_required
def api_export_batch():
    all_orders = []
    for shop in all_connected_shops():
        all_orders.extend(fetch_packed_ready_orders(shop))

    if not all_orders:
        return jsonify({"error": "No packed-ready orders found to export"}), 404

    tmp_path = os.path.join(tempfile.gettempdir(), f"auspost_export_{int(time.time())}.csv")
    export_orders_to_xlsx(all_orders, tmp_path)

    by_shop = {}
    for o in all_orders:
        by_shop.setdefault(o["shop"], []).extend([o["order_id"]] + o.get("extra_order_ids", []))
    for shop, order_ids in by_shop.items():
        mark_orders_exported(shop, order_ids)

    save_last_batch_export(all_orders)

    return send_file(tmp_path, as_attachment=True, download_name="auspost_export.csv", mimetype="text/csv")


@app.route("/api/export_single")
@login_required
def api_export_single():
    shop = request.args.get("shop")
    order_id = request.args.get("order_id")
    if not shop or not order_id:
        return jsonify({"error": "missing shop or order_id"}), 400

    orders = fetch_packed_ready_orders(shop, single_order_id=order_id)
    if not orders:
        return jsonify({"error": "Could not load that order"}), 404

    tmp_path = os.path.join(tempfile.gettempdir(), f"auspost_single_{order_id}_{int(time.time())}.csv")
    export_orders_to_xlsx(orders, tmp_path)
    mark_orders_exported(shop, [order_id])

    return send_file(tmp_path, as_attachment=True, download_name=f"auspost_order_{orders[0]['order_number'].replace('#','')}.csv", mimetype="text/csv")


@app.route("/api/reexport_orders")
@login_required
def api_reexport_orders():
    """Re-export specific orders BY EXCEPTION — e.g. a handful of orders
    that errored on the AusPost side and need fixed data resent. Unlike
    /api/export_batch, this does NOT touch every packed-ready order — only
    the exact order numbers passed in. For each one: strips its stale
    'auspost-exported' tag (if present), regenerates its row with whatever
    the current code produces, bundles them into one CSV, then re-tags them
    exported again afterward.

    Usage: /api/reexport_orders?orders=16341,16342,3812
    (accepts '#' prefixes too, e.g. #16341 — they're stripped automatically)
    """
    raw = request.args.get("orders", "").strip()
    if not raw:
        return jsonify({"error": "no order numbers provided — use ?orders=16341,16342,..."}), 400

    order_numbers = [o.strip().lstrip("#") for o in raw.split(",") if o.strip()]
    if not order_numbers:
        return jsonify({"error": "no valid order numbers found in that list"}), 400

    all_orders = []
    not_found = []
    for order_number in order_numbers:
        shop, order = find_order_by_number_any_status(order_number)
        if not order:
            not_found.append(order_number)
            continue
        unmark_order_exported(shop, order["id"])
        all_orders.append(process_order_for_export(shop, order))

    if not all_orders:
        return jsonify({"error": "none of the given order numbers were found in any connected store",
                         "not_found": not_found}), 404

    tmp_path = os.path.join(tempfile.gettempdir(), f"auspost_reexport_{int(time.time())}.csv")
    export_orders_to_xlsx(all_orders, tmp_path)

    by_shop = {}
    for o in all_orders:
        by_shop.setdefault(o["shop"], []).extend([o["order_id"]] + o.get("extra_order_ids", []))
    for shop, order_ids in by_shop.items():
        mark_orders_exported(shop, order_ids)

    response = send_file(tmp_path, as_attachment=True, download_name="auspost_reexport.csv", mimetype="text/csv")
    if not_found:
        # Surfaced as a response header rather than failing the whole
        # request — the orders that WERE found still export successfully,
        # and the front-end can read this header to warn about the rest.
        response.headers["X-Not-Found-Orders"] = ",".join(not_found)
    return response


@app.route("/api/reexport_last_batch")
@login_required
def api_reexport_last_batch():
    """Re-exports exactly the same set of orders that the most recent
    /api/export_batch run sent — no need to type out order numbers by hand
    when a whole batch needs redoing (e.g. it went out with errors and got
    fixed since). Reuses the same strip-tag / regenerate / re-tag flow as
    /api/reexport_orders, just sourced from the saved last-batch record
    instead of a typed-in list."""
    record = load_last_batch_export()
    if not record or not record.get("orders"):
        return jsonify({"error": "No previous batch export found to redo. "
                                  "This resets if the server restarts, or if "
                                  "no batch export has been run yet this session."}), 404

    all_orders = []
    not_found = []
    for entry in record["orders"]:
        shop, order_id = entry["shop"], entry["order_id"]
        data = shopify_get(shop, f"orders/{order_id}.json")
        if not data or not data.get("order"):
            not_found.append(entry.get("order_number", str(order_id)))
            continue
        unmark_order_exported(shop, order_id)
        all_orders.append(process_order_for_export(shop, data["order"]))

    if not all_orders:
        return jsonify({"error": "None of the orders from the last batch could be found anymore",
                         "not_found": not_found}), 404

    tmp_path = os.path.join(tempfile.gettempdir(), f"auspost_reexport_last_batch_{int(time.time())}.csv")
    export_orders_to_xlsx(all_orders, tmp_path)

    by_shop = {}
    for o in all_orders:
        by_shop.setdefault(o["shop"], []).extend([o["order_id"]] + o.get("extra_order_ids", []))
    for shop, order_ids in by_shop.items():
        mark_orders_exported(shop, order_ids)

    save_last_batch_export(all_orders)  # this re-export IS now the new "last batch"

    response = send_file(tmp_path, as_attachment=True, download_name="auspost_reexport_last_batch.csv", mimetype="text/csv")
    if not_found:
        response.headers["X-Not-Found-Orders"] = ",".join(not_found)
    return response


@app.route("/scan")
@login_required
def scan():
    return render_template("scan.html", current_user=session.get("logged_in_user"))


@app.route("/fulfill-pdf", methods=["GET", "POST"])
@login_required
def fulfill_pdf():
    if request.method == "GET":
        return """
        <html><body style="background:#111;color:#fff;font-family:sans-serif;padding:30px;max-width:600px;">
        <h2>📦 Fulfill from AusPost Label PDF</h2>
        <p style="color:#aaa;">Upload the label PDF you got back from AusPost after printing.
        This extracts every order number + tracking number pair on it and fulfills each order
        in Shopify with tracking attached — across whichever connected store the order belongs to.</p>
        <form method="post" enctype="multipart/form-data">
          <input type="file" name="pdf" accept="application/pdf" required
            style="margin-bottom:15px;color:#fff;display:block;">
          <button type="submit" style="padding:10px 20px;background:#2ecc71;color:#000;
            border:none;border-radius:6px;font-weight:bold;cursor:pointer;">Extract &amp; Fulfill</button>
        </form>
        <p style="margin-top:20px;"><a href="/scan" style="color:#888;">&larr; Back to scan</a></p>
        </body></html>
        """

    file = request.files.get("pdf")
    if not file or file.filename == "":
        return jsonify({"error": "no PDF uploaded"}), 400

    try:
        pairs = extract_tracking_pairs_from_pdf(file.stream)
    except Exception as e:
        log_error("fulfill_pdf_parse", str(e))
        return jsonify({"error": f"Could not parse PDF: {e}"}), 500

    if not pairs:
        return jsonify({"error": "No order/tracking pairs found in this PDF — "
                                  "check it's an AusPost label export, not a different document"}), 400

    results = []
    for pair in pairs:
        order_number = pair["order_number"]
        tracking_number = pair["tracking_number"]
        shop, order = find_order_by_number(order_number)
        if not order:
            results.append({
                "order_number": order_number, "tracking_number": tracking_number,
                "shop": None, "ok": False, "message": "Order not found in any connected store",
            })
            continue

        ok, message = fulfill_order_with_tracking(shop, order["id"], tracking_number, pair["carrier"])
        results.append({
            "order_number": order_number, "tracking_number": tracking_number,
            "shop": shop, "ok": ok, "message": message,
        })
        if not ok:
            log_error("fulfill_pdf", f"order=#{order_number} shop={shop} tracking={tracking_number} | {message}")

        # If this order shipped together with another as one combined
        # parcel, the same tracking number belongs on both — the label
        # only shows ONE order number, but the physical parcel covers two
        # Shopify orders, so the second one needs fulfilling here too.
        if ok:
            tags = [t.strip().lower() for t in (order.get("tags") or "").split(",")]
            partner_number = find_merge_partner_number(tags)
            if partner_number:
                partner_order = find_order_by_number_in_shop(shop, partner_number)
                if partner_order:
                    partner_ok, partner_message = fulfill_order_with_tracking(
                        shop, partner_order["id"], tracking_number, pair["carrier"])
                    results.append({
                        "order_number": partner_order["name"], "tracking_number": tracking_number,
                        "shop": shop, "ok": partner_ok,
                        "message": f"(merged with #{order_number}) {partner_message}",
                    })
                    if not partner_ok:
                        log_error("fulfill_pdf", f"order=#{partner_number} shop={shop} tracking={tracking_number} "
                                                  f"| merge-partner fulfillment failed: {partner_message}")

    succeeded = sum(1 for r in results if r["ok"])
    failed = len(results) - succeeded

    row_parts = []
    for r in results:
        row_bg = "#1a3d1a" if r["ok"] else "#3d1a1a"
        status_icon = "✅" if r["ok"] else "❌"
        store_label = r["shop"] or "—"
        row_parts.append(
            f"<tr style='background:{row_bg};'>"
            f"<td style='padding:8px;'>#{r['order_number']}</td>"
            f"<td style='padding:8px;'>{r['tracking_number']}</td>"
            f"<td style='padding:8px;'>{store_label}</td>"
            f"<td style='padding:8px;'>{status_icon}</td>"
            f"<td style='padding:8px;'>{r['message']}</td></tr>"
        )
    rows = "".join(row_parts)
    return f"""
    <html><head><title>Fulfill Results</title></head>
    <body style="background:#111;color:#fff;font-family:sans-serif;padding:30px;">
    <h2>📦 Fulfillment Results</h2>
    <p>{succeeded} succeeded, {failed} failed — {len(results)} parcels parsed from the PDF.</p>
    <table style="width:100%;border-collapse:collapse;">
      <tr style="text-align:left;color:#888;">
        <th style="padding:8px;">Order</th><th style="padding:8px;">Tracking</th>
        <th style="padding:8px;">Store</th><th style="padding:8px;">Status</th><th style="padding:8px;">Detail</th>
      </tr>
      {rows}
    </table>
    <p style="margin-top:20px;">
      <a href="/fulfill-pdf" style="color:#2ecc71;">Upload another PDF</a> &nbsp;|&nbsp;
      <a href="/scan" style="color:#888;">Back to scan</a>
    </p>
    </body></html>
    """


AUTO_CONFIRM_KEYWORDS = ["package protection", "free return", "free returns"]


def build_pack_screen_line_items(shop, order, from_order_number=None):
    """Builds the pack-screen line item list for one order. Pulled out as
    its own function so a merged shipment can call this once per order and
    concatenate the results into a single combined checklist, rather than
    duplicating this whole block. from_order_number tags each item with
    which of the two merged orders it came from, purely for display —
    None for the normal, non-merged case."""
    line_items = []
    for li in order["line_items"]:
        # current_quantity accounts for order edits/removals — quantity
        # alone still shows the ORIGINAL ordered amount even after an
        # item has been fully swapped out or removed, so filtering on
        # quantity would keep showing removed items as needing packing.
        # Explicitly check for None rather than using "or" — a
        # legitimate current_quantity of 0 (the removed-item case we're
        # trying to catch) is falsy too, so "or" would wrongly undo the
        # filter for exactly the case it's meant to handle.
        current_qty = li.get("current_quantity", li["quantity"])
        if current_qty is None:
            current_qty = li["quantity"]
        if current_qty == 0:
            continue
        barcode = li.get("barcode")
        if not barcode:
            barcode = lookup_live_barcode(shop, li.get("sku"))
        title_lower = li["title"].lower()
        auto_confirm = any(kw in title_lower for kw in AUTO_CONFIRM_KEYWORDS)
        item = {
            "id": li["id"],
            "title": li["title"],
            "sku": li.get("sku"),
            "quantity": current_qty,
            "barcode": barcode,
            "auto_confirm": auto_confirm,
        }
        if from_order_number:
            item["from_order"] = from_order_number
        line_items.append(item)
    return line_items


@app.route("/api/lookup_order")
def api_lookup_order():
    order_number = request.args.get("order_number", "").strip().lstrip("#")
    if not order_number:
        return jsonify({"error": "no order number provided"}), 400

    try:
        shop, order = find_order_by_number(order_number)
        if not order:
            return jsonify({"error": f"Order #{order_number} not found in any connected store"}), 404

        line_items = build_pack_screen_line_items(shop, order)

        tags_raw = order.get("tags") or ""
        existing_tag_list = [t.strip().lower() for t in tags_raw.split(",")]

        # If this order is merged with another, pull the partner order's
        # items in too so the pack screen shows ONE combined checklist for
        # both — reading the partner's number straight off the tag rather
        # than searching for it (see find_merge_partner_number's docstring
        # for why that matters).
        merged_with_number = find_merge_partner_number(existing_tag_list)
        merged_order_name = None
        if merged_with_number:
            partner_order = find_order_by_number_in_shop(shop, merged_with_number)
            if partner_order:
                merged_order_name = partner_order["name"]
                line_items = (
                    [dict(li, from_order=order["name"]) for li in line_items]
                    + build_pack_screen_line_items(shop, partner_order, from_order_number=partner_order["name"])
                )

        shipping_address = order.get("shipping_address") or {}
        shipping_country = shipping_address.get("country_code", "")
        is_international = shipping_country not in ("AU", "")
        estimated_weight = estimate_order_weight(line_items)
        needs_express_tag = "express-upgrade" in existing_tag_list

        # small, single-product-type domestic orders that might qualify for
        # the cheaper XS satchel. Quantity allowed depends on actual item
        # weight — e.g. 2 kids' rashguards or 2 finger tape rolls can still
        # land under 250g together, even though a single adult rashguard
        # already exceeds it alone. See xs_satchel_item_lookup() above.
        distinct_titles = set(li["title"].lower() for li in line_items)
        total_qty = sum(li["quantity"] for li in line_items)
        is_small_item_order = False
        if len(distinct_titles) == 1:
            only_title = next(iter(distinct_titles))
            lookup = xs_satchel_item_lookup(only_title)
            if lookup:
                _weight_g, max_qty = lookup
                is_small_item_order = total_qty <= max_qty

        checks = {
            # BUG FIX: this used to also require estimated_weight <= 0.25,
            # but the weight-estimate table rates a single rashguard or pair
            # of shorts at 0.3kg — ABOVE that 0.25 threshold — so the check
            # could never fire for the exact two product types it exists to
            # catch. It's been mathematically unreachable since it was built.
            # Dropping that gate: is_small_item_order already restricts this
            # to a single qualifying product, and the whole point of the
            # banner is to ask the packer to physically verify the actual
            # weight — pre-filtering on a rough, already-known-inaccurate
            # estimate defeats that purpose.
            "ask_xs_satchel": (not is_international) and is_small_item_order,
            "ask_express_upgrade": is_international and estimated_weight > 2,
            "ask_split_shipment": (not is_international) and estimated_weight > 5,
        }
        xs_satchel_confirmed = "xs-satchel-ok" in existing_tag_list
        split_shipment_confirmed = "needs-split-shipment" in existing_tag_list
        express_confirmed_tag = "confirmed-express" in existing_tag_list

        return jsonify({
            "shop": shop,
            "order_id": order["id"],
            "order_number": order["name"],
            "customer": shipping_address.get("name", ""),
            "is_international": is_international,
            "estimated_weight_kg": estimated_weight,
            "checks": checks,
            "xs_satchel_confirmed": xs_satchel_confirmed,
            "split_shipment_confirmed": split_shipment_confirmed,
            "express_confirmed_tag": express_confirmed_tag,
            "needs_express_tag": needs_express_tag,
            "line_items": line_items,
            "merged_with": merged_order_name,
        })
    except Exception as e:
        log_error("lookup_order", f"order_number={order_number} | {e}")
        return jsonify({"error": f"Internal error looking up order — see /logs for detail"}), 500


@app.route("/api/find_order_for_merge")
@login_required
def api_find_order_for_merge():
    """Preview step before merging — looks up the candidate partner order
    (same shop only) and returns its name/address so the packer can
    visually confirm it's genuinely the same customer before committing.
    Doesn't tag anything yet; that only happens in /api/merge_orders."""
    shop = request.args.get("shop", "")
    order_number = request.args.get("order_number", "").strip().lstrip("#")
    if not shop or not order_number:
        return jsonify({"error": "missing shop or order_number"}), 400

    order = find_order_by_number_in_shop(shop, order_number)
    if not order:
        return jsonify({"error": f"Order #{order_number} not found in {shop}"}), 404

    addr = order.get("shipping_address") or {}
    return jsonify({
        "order_id": order["id"],
        "order_number": order["name"],
        "name": addr.get("name", ""),
        "address1": addr.get("address1", ""),
        "city": addr.get("city", ""),
        "zip": addr.get("zip", ""),
    })


@app.route("/api/merge_orders", methods=["POST"])
@login_required
def api_merge_orders():
    """Applies the merge tags: each order gets 'merged-with-<other order's
    number>'. Both tags point directly at each other's order number rather
    than a shared group ID — see find_merge_partner_number's docstring for
    why (avoids any dependency on Shopify's tag search index)."""
    body = request.get_json(force=True)
    shop = body.get("shop")
    order_id_a = body.get("order_id_a")
    order_number_a = str(body.get("order_number_a", "")).lstrip("#")
    order_id_b = body.get("order_id_b")
    order_number_b = str(body.get("order_number_b", "")).lstrip("#")

    if not all([shop, order_id_a, order_number_a, order_id_b, order_number_b]):
        return jsonify({"error": "missing required fields", "ok": False}), 400

    ok_a = add_order_tag(shop, order_id_a, f"merged-with-{order_number_b}")
    ok_b = add_order_tag(shop, order_id_b, f"merged-with-{order_number_a}")

    if not (ok_a and ok_b):
        log_error("merge_orders", f"shop={shop} a=#{order_number_a} b=#{order_number_b} | one or both tag writes failed")
        return jsonify({"error": "Failed to tag one or both orders — see /logs", "ok": False}), 500

    return jsonify({"ok": True})


@app.route("/api/confirm_check", methods=["POST"])
@login_required
def api_confirm_check():
    try:
        body = request.get_json()
        shop = body["shop"]
        order_id = body["order_id"]
        check_tag = body["tag"]  # e.g. "xs-satchel-ok", "confirmed-express", "needs-split-shipment"

        current = shopify_get(shop, f"orders/{order_id}.json")
        if not current:
            return jsonify({"error": "Could not fetch order — see /logs", "ok": False}), 500

        existing_tags = (current["order"].get("tags") or "")
        tag_list = [t.strip() for t in existing_tags.split(",") if t.strip()]
        if check_tag not in tag_list:
            tag_list.append(check_tag)

        payload = {"order": {"id": order_id, "tags": ", ".join(tag_list)}}
        status_code, resp = shopify_put(shop, f"orders/{order_id}.json", payload)
        if status_code not in (200, 201):
            log_error("confirm_check", f"shop={shop} order_id={order_id} | {status_code}: {resp}")
            return jsonify({"error": "failed to tag order", "ok": False}), 500
        return jsonify({"ok": True})
    except Exception as e:
        log_error("confirm_check", str(e))
        return jsonify({"error": str(e), "ok": False}), 500


@app.route("/api/mark_order", methods=["POST"])
@login_required
def api_mark_order():
    try:
        body = request.get_json()
        shop = body["shop"]
        order_id = body["order_id"]
        status = body["status"]
        missing_items = body.get("missing_items", [])
        packed_by = session.get("logged_in_user", "unknown")

        current = shopify_get(shop, f"orders/{order_id}.json")
        if not current:
            log_error("mark_order", f"shop={shop} order_id={order_id} | failed to fetch current order — check token/scopes")
            return jsonify({"error": "Could not fetch order from Shopify — see /logs", "ok": False}), 500

        # Local pickup orders have no shipping address at all in Shopify —
        # these should never be eligible for AusPost export, so they get a
        # different tag entirely rather than "packed-ready", which
        # fetch_packed_ready_orders uses as its trigger for pulling orders
        # into the export batch.
        is_pickup = not current["order"].get("shipping_address")

        if status == "ready":
            tag = "packed-pickup" if is_pickup else "packed-ready"
        else:
            tag = "packed-incomplete"

        existing_tags = (current["order"].get("tags") or "")
        tag_list = [t.strip() for t in existing_tags.split(",") if t.strip()]
        if tag not in tag_list:
            tag_list.append(tag)

        note_addition = f"\n[Pack Verify] Packed by: {packed_by} ({status})"
        if status == "incomplete" and missing_items:
            names = ", ".join(i["title"] for i in missing_items)
            note_addition += f" — Missing: {names}"

        payload = {
            "order": {
                "id": order_id,
                "tags": ", ".join(tag_list),
            }
        }
        if note_addition:
            current_note = (current["order"].get("note") or "")
            payload["order"]["note"] = current_note + note_addition

        status_code, resp = shopify_put(shop, f"orders/{order_id}.json", payload)
        if status_code not in (200, 201):
            log_error("mark_order", f"shop={shop} order_id={order_id} | Shopify returned {status_code}: {resp}")
            return jsonify({"error": "failed to tag order — see /logs", "detail": resp, "ok": False}), 500

        # If this order is merged with another (shipping together as one
        # physical parcel), apply the same status tag to the partner order
        # too — they need to move through packing as one unit, since only
        # one of them will actually get exported/fulfilled with tracking.
        partner_number = find_merge_partner_number([t.lower() for t in tag_list])
        if partner_number:
            partner_order = find_order_by_number_in_shop(shop, partner_number)
            if partner_order:
                partner_tags_raw = (partner_order.get("tags") or "")
                partner_tag_list = [t.strip() for t in partner_tags_raw.split(",") if t.strip()]
                if tag not in partner_tag_list:
                    partner_tag_list.append(tag)
                    shopify_put(shop, f"orders/{partner_order['id']}.json", {
                        "order": {"id": partner_order["id"], "tags": ", ".join(partner_tag_list)}
                    })

        return jsonify({"ok": True})
    except Exception as e:
        log_error("mark_order", str(e))
        return jsonify({"error": f"Internal error — see /logs", "ok": False}), 500


@app.route("/admin/reveal_token")
def reveal_token():
    secret = request.args.get("secret")
    if secret != os.environ.get("ADMIN_SECRET", "change-me"):
        return "Forbidden", 403
    shop = request.args.get("shop")
    token = get_token_for_shop(shop)
    if not token:
        return f"No token stored for {shop}", 404
    return f"Token for {shop}:\n{token}\n\nScopes: {load_tokens().get(shop, {}).get('scope')}"


@app.route("/logs")
@login_required
def view_logs():
    logs = []
    if os.path.exists(LOGS_FILE):
        with open(LOGS_FILE, "r") as f:
            logs = json.load(f)
    logs = list(reversed(logs))
    rows = "".join(
        f"<tr><td>{l['timestamp']}</td><td>{l['context']}</td>"
        f"<td><pre style='white-space:pre-wrap;margin:0;'>{l['detail']}</pre></td></tr>"
        for l in logs
    )
    return f"""
    <html><head><title>Pack Verify — Error Log</title>
    <style>
      body {{ font-family: monospace; background:#111; color:#eee; padding:20px; }}
      table {{ width:100%; border-collapse: collapse; }}
      td {{ border-bottom:1px solid #333; padding:8px; vertical-align:top; }}
      td:first-child {{ white-space:nowrap; color:#888; }}
      td:nth-child(2) {{ white-space:nowrap; color:#e7a23c; font-weight:bold; }}
    </style>
    </head><body>
    <h2>Error Log (most recent first, last 200 kept)</h2>
    <p>Refresh this page any time something looks wrong on the scan screen.</p>
    <table>{rows if rows else "<tr><td colspan=3>No errors logged yet 🎉</td></tr>"}</table>
    </body></html>
    """


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
