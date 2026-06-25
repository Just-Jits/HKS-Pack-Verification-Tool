import os
import json
import time
import hmac
import hashlib
import requests
from flask import Flask, request, redirect, render_template, jsonify, session
from urllib.parse import urlencode

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me-in-railway-env")

CLIENT_ID = os.environ.get("SHOPIFY_CLIENT_ID")  # fallback/default app
CLIENT_SECRET = os.environ.get("SHOPIFY_CLIENT_SECRET")  # fallback/default app
APP_URL = os.environ["APP_URL"]  # e.g. https://pack-verify-tool-production.up.railway.app
SCOPES = "read_orders,write_orders,read_products,write_products"
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


def estimate_order_weight(line_items):
    total = 0.0
    for li in line_items:
        total += estimate_item_weight(li["title"]) * li["quantity"]
    return round(total, 2)


# ---------- core: find order across all connected stores ----------

def find_order_by_number(order_number):
    """order_number = plain digits scanned off the packing slip barcode, e.g. '9511'"""
    for shop in all_connected_shops():
        data = shopify_get(shop, "orders.json", params={"name": f"#{order_number}", "status": "any"})
        if data and data.get("orders"):
            order = data["orders"][0]
            return shop, order
    return None, None


@app.route("/")
def home():
    shops = all_connected_shops()
    return render_template("home.html", shops=shops, app_url=APP_URL)


@app.route("/scan")
def scan():
    return render_template("scan.html")


@app.route("/api/lookup_order")
def api_lookup_order():
    order_number = request.args.get("order_number", "").strip().lstrip("#")
    if not order_number:
        return jsonify({"error": "no order number provided"}), 400

    try:
        shop, order = find_order_by_number(order_number)
        if not order:
            return jsonify({"error": f"Order #{order_number} not found in any connected store"}), 404

        line_items = []
        for li in order["line_items"]:
            line_items.append({
                "id": li["id"],
                "title": li["title"],
                "sku": li.get("sku"),
                "quantity": li["quantity"],
                "barcode": li.get("barcode"),
            })

        shipping_address = order.get("shipping_address") or {}
        shipping_country = shipping_address.get("country_code", "")
        is_international = shipping_country not in ("AU", "")
        estimated_weight = estimate_order_weight(line_items)
        tags_raw = order.get("tags") or ""
        needs_express_tag = "express-upgrade" in [t.strip().lower() for t in tags_raw.split(",")]

        weight_warning = None
        if is_international and estimated_weight > 2:
            weight_warning = "International parcel est. over 2kg — check if Express upgrade is needed"
        elif not is_international and estimated_weight > 5:
            weight_warning = "Domestic parcel est. over 5kg — check if it needs splitting into 2 satchels"

        return jsonify({
            "shop": shop,
            "order_id": order["id"],
            "order_number": order["name"],
            "customer": shipping_address.get("name", ""),
            "is_international": is_international,
            "estimated_weight_kg": estimated_weight,
            "weight_warning": weight_warning,
            "needs_express_tag": needs_express_tag,
            "line_items": line_items,
        })
    except Exception as e:
        log_error("lookup_order", f"order_number={order_number} | {e}")
        return jsonify({"error": f"Internal error looking up order — see /logs for detail"}), 500


@app.route("/api/mark_order", methods=["POST"])
def api_mark_order():
    try:
        body = request.get_json()
        shop = body["shop"]
        order_id = body["order_id"]
        status = body["status"]
        missing_items = body.get("missing_items", [])

        tag = "packed-ready" if status == "ready" else "packed-incomplete"

        current = shopify_get(shop, f"orders/{order_id}.json")
        if not current:
            log_error("mark_order", f"shop={shop} order_id={order_id} | failed to fetch current order — check token/scopes")
            return jsonify({"error": "Could not fetch order from Shopify — see /logs", "ok": False}), 500

        existing_tags = (current["order"].get("tags") or "")
        tag_list = [t.strip() for t in existing_tags.split(",") if t.strip()]
        if tag not in tag_list:
            tag_list.append(tag)

        note_addition = ""
        if status == "incomplete" and missing_items:
            names = ", ".join(i["title"] for i in missing_items)
            note_addition = f"\n[Pack Verify] Missing at pack time: {names}"

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

# v2
