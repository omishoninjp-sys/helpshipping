"""
客人集運預報系統
GOYOUTATI x OMISHONIN 雲倉
"""

from flask import Flask, request, jsonify, render_template, make_response, send_file
from datetime import datetime
import requests
import json
import os
import sqlite3
import csv
import io
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)

# ============ 設定區（從環境變數讀取）============
JPD_BASE_URL = "https://biz.cloudwh.jp"
JPD_EMAIL = os.environ.get("JPD_EMAIL", "")
JPD_PASSWORD = os.environ.get("JPD_PASSWORD", "")
JPD_WAREHOUSE_ID = int(os.environ.get("JPD_WAREHOUSE_ID", "1"))

SHOPIFY_STORE = os.environ.get("SHOPIFY_STORE", "")
SHOPIFY_ACCESS_TOKEN = os.environ.get("SHOPIFY_ACCESS_TOKEN", "")

# 預設運費（台幣/kg），0 表示未設定
DEFAULT_SHIPPING_RATE = int(os.environ.get("DEFAULT_SHIPPING_RATE", "0"))

# 台幣 → 日圓匯率（可透過環境變數調整）
TWD_TO_JPY_RATE = float(os.environ.get("TWD_TO_JPY_RATE", "5.0"))

DB_PATH = os.environ.get("DB_PATH", "packages.db")
# ================================


# ============ SQLite 初始化 ============

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS packages (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            logis_num   TEXT,
            product_name TEXT   DEFAULT '',
            weight      TEXT    DEFAULT '',
            status      TEXT    DEFAULT '已到貨',
            note        TEXT    DEFAULT '',
            in_date     TEXT,
            created_at  TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS shipment_requests (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            customer_name TEXT  DEFAULT '',
            package_ids TEXT    NOT NULL,
            package_summary TEXT DEFAULT '',
            status      TEXT    DEFAULT '待處理',
            note        TEXT    DEFAULT '',
            admin_note  TEXT    DEFAULT '',
            created_at  TEXT    NOT NULL,
            updated_at  TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS forecasts (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            customer_name TEXT  DEFAULT '',
            items_json  TEXT    NOT NULL,
            status      TEXT    DEFAULT '待處理',
            note        TEXT    DEFAULT '',
            created_at  TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS admin_settings (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS addresses (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            g_code      TEXT    NOT NULL,
            label       TEXT    DEFAULT '',
            recipient   TEXT    NOT NULL,
            phone       TEXT    NOT NULL,
            zipcode     TEXT    DEFAULT '',
            address     TEXT    NOT NULL,
            is_default  INTEGER DEFAULT 0,
            created_at  TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS announcements (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            title       TEXT    NOT NULL,
            content     TEXT    NOT NULL,
            is_active   INTEGER DEFAULT 1,
            created_at  TEXT    NOT NULL
        )
    """)
    # 帳單欄位遷移（已存在的表加欄位）
    for col, col_type, default in [
        ("admin_note", "TEXT", "''"),
        ("updated_at", "TEXT", "NULL"),
        ("billed_weight", "REAL", "0"),
        ("rate_per_kg", "REAL", "0"),
        ("shipping_fee", "REAL", "0"),
        ("handling_fee", "REAL", "0"),
        ("total_fee", "REAL", "0"),
        ("payment_last5", "TEXT", "''"),
        ("payment_at", "TEXT", "''"),
        ("tracking_num", "TEXT", "''"),
        ("extra_services", "TEXT", "''"),
        ("ship_recipient", "TEXT", "''"),
        ("ship_phone", "TEXT", "''"),
        ("ship_address", "TEXT", "''"),
    ]:
        try:
            conn.execute(f"ALTER TABLE shipment_requests ADD COLUMN {col} {col_type} DEFAULT {default}")
        except:
            pass
    conn.commit()
    conn.close()


init_db()

# ============ 工具函數 ============

def normalize_phone(phone_raw):
    phone = phone_raw.replace(" ", "").replace("-", "")
    if phone.startswith("+886"):
        phone = "0" + phone[4:]
    elif phone.startswith("+81"):
        phone = "0" + phone[3:]
    return phone


def twd_to_jpy(twd_rate):
    """台幣運費 → 日圓運費（四捨五入至整數）"""
    return round(twd_rate * TWD_TO_JPY_RATE)


def jpd_request(operation, data):
    url = f"{JPD_BASE_URL}/api/json.php?Service=SDC&Operation={operation}"
    payload = {
        "login_email": JPD_EMAIL,
        "login_password": JPD_PASSWORD,
        "data": data
    }
    print(f"\n{'='*50}")
    print(f"📤 JPD API 請求: {operation}")
    try:
        response = requests.post(url, json=payload, timeout=30)
        result = response.json()
        return result
    except Exception as e:
        print(f"❌ 錯誤: {e}")
        return {"error": str(e)}


def shopify_graphql(query, variables=None):
    graphql_url = f"https://{SHOPIFY_STORE}/admin/api/2026-01/graphql.json"
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    payload = {"query": query}
    if variables:
        payload["variables"] = variables
    try:
        response = requests.post(graphql_url, headers=headers, json=payload, timeout=30)
        return response.json()
    except Exception as e:
        print(f"❌ GraphQL 錯誤: {e}")
        return {"error": str(e)}


def shopify_request(endpoint, method="GET", data=None):
    url = f"https://{SHOPIFY_STORE}/admin/api/2026-01/{endpoint}"
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    try:
        if method == "GET":
            response = requests.get(url, headers=headers, timeout=30)
        elif method == "POST":
            response = requests.post(url, headers=headers, json=data, timeout=30)
        return response.json()
    except Exception as e:
        return {"error": str(e)}


# 會員快取（避免每次登入都打 Shopify API）
_customers_cache = {"data": None, "time": 0}
CACHE_TTL = 300  # 5 分鐘


def get_all_goyoutati_customers(force_refresh=False):
    global _customers_cache
    now = time.time()
    if not force_refresh and _customers_cache["data"] is not None and (now - _customers_cache["time"]) < CACHE_TTL:
        return _customers_cache["data"]

    customers = _fetch_customers_from_shopify()
    _customers_cache = {"data": customers, "time": now}
    return customers


def _fetch_customers_from_shopify():
    customers = []
    cursor = None
    has_next = True

    while has_next:
        after_clause = f', after: "{cursor}"' if cursor else ''
        graphql_query = """
        {
            metafieldDefinitions(first: 1, ownerType: CUSTOMER, namespace: "custom", key: "goyoutati_id") {
                edges {
                    node {
                        id
                        metafields(first: 100%s) {
                            edges {
                                node {
                                    value
                                    owner {
                                        ... on Customer {
                                            id
                                            firstName
                                            lastName
                                            email
                                            phone
                                            defaultAddress {
                                                phone
                                                province
                                                city
                                                address1
                                                address2
                                            }
                                            createdAt
                                            shippingRate: metafield(namespace: "custom", key: "shipping_rate") {
                                                value
                                            }
                                        }
                                    }
                                }
                                cursor
                            }
                            pageInfo {
                                hasNextPage
                            }
                        }
                    }
                }
            }
        }
        """ % after_clause

        result = shopify_graphql(graphql_query)
        has_next = False

        if "data" not in result:
            break

        definitions = result["data"].get("metafieldDefinitions", {}).get("edges", [])
        if not definitions:
            break

        metafields_data = definitions[0]["node"].get("metafields", {})
        edges = metafields_data.get("edges", [])
        page_info = metafields_data.get("pageInfo", {})
        has_next = page_info.get("hasNextPage", False)

        for mf in edges:
            node = mf["node"]
            cursor = mf.get("cursor")
            g_code = node.get("value", "")
            owner = node.get("owner", {})
            if not g_code or not owner:
                continue
            gid = owner.get("id", "")
            customer_id = gid.split("/")[-1] if "/" in gid else gid
            customer_name = f"{owner.get('lastName', '')}{owner.get('firstName', '')}".strip()
            if not customer_name:
                customer_name = owner.get("email", "")
            default_address = owner.get("defaultAddress") or {}
            phone_raw = default_address.get("phone") or owner.get("phone") or ""
            phone = normalize_phone(phone_raw)
            address = " ".join(filter(None, [
                default_address.get("province", ""),
                default_address.get("city", ""),
                default_address.get("address1", ""),
                default_address.get("address2", "")
            ])).strip()
            rate_mf = owner.get("shippingRate")
            # shipping_rate 現在儲存台幣值
            shipping_rate_twd = rate_mf["value"] if rate_mf and rate_mf.get("value") else ""
            customers.append({
                "g_code": g_code,
                "customer_id": customer_id,
                "gid": gid,
                "name": customer_name,
                "email": owner.get("email", ""),
                "address": address,
                "phone": phone,
                "phone_raw": phone_raw,
                "shipping_rate": shipping_rate_twd,  # 台幣
                "created_at": owner.get("createdAt", "")
            })
    return customers


# ============ 路由 ============

@app.route("/admin")
def admin_page():
    return render_template("admin.html")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config")
def get_config():
    """回傳前端所需設定（匯率等）"""
    return jsonify({
        "twd_to_jpy_rate": TWD_TO_JPY_RATE
    })


def get_admin_password():
    """取得管理員密碼：DB 優先，否則環境變數"""
    conn = get_db()
    row = conn.execute("SELECT value FROM admin_settings WHERE key='admin_password'").fetchone()
    conn.close()
    if row:
        return row["value"]
    return os.environ.get("ADMIN_PASSWORD", "admin123")


@app.route("/api/admin/verify", methods=["POST"])
def admin_verify():
    data = request.json
    password = data.get("password", "")
    if password == get_admin_password():
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "密碼錯誤"})


@app.route("/api/admin/change_password", methods=["POST"])
def admin_change_password():
    data = request.json
    current = data.get("current", "")
    new_pwd = data.get("new_password", "").strip()
    confirm = data.get("confirm", "").strip()

    if current != get_admin_password():
        return jsonify({"success": False, "error": "目前密碼錯誤"})
    if not new_pwd or len(new_pwd) < 4:
        return jsonify({"success": False, "error": "新密碼至少 4 個字元"})
    if new_pwd != confirm:
        return jsonify({"success": False, "error": "兩次密碼不一致"})

    conn = get_db()
    conn.execute(
        "INSERT OR REPLACE INTO admin_settings (key, value) VALUES ('admin_password', ?)",
        (new_pwd,)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "密碼已更新"})


@app.route("/api/admin/members", methods=["GET"])
def get_all_members():
    try:
        force = request.args.get("refresh") == "1"
        members = get_all_goyoutati_customers(force_refresh=force)
        members.sort(key=lambda x: x["g_code"])
        used_numbers = set()
        for m in members:
            if m["g_code"].startswith("G"):
                try:
                    used_numbers.add(int(m["g_code"][1:]))
                except:
                    pass
        max_number = max(used_numbers) if used_numbers else 0
        next_number = 1
        while next_number in used_numbers:
            next_number += 1
        next_g_code = f"G{next_number:04d}"
        return jsonify({
            "success": True,
            "members": members,
            "total": len(members),
            "max_number": max_number,
            "next_g_code": next_g_code,
            "default_shipping_rate": DEFAULT_SHIPPING_RATE,  # 台幣
            "twd_to_jpy_rate": TWD_TO_JPY_RATE
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/admin/shipping_rate", methods=["POST"])
def set_shipping_rate():
    data = request.json
    customer_gid = data.get("customer_gid", "")
    shipping_rate = data.get("shipping_rate", "")  # 台幣
    if not customer_gid:
        return jsonify({"success": False, "error": "缺少客戶 ID"})
    if shipping_rate == "" or shipping_rate is None:
        return jsonify({"success": False, "error": "請輸入運費"})
    try:
        rate_val = int(shipping_rate)
        if rate_val < 0:
            return jsonify({"success": False, "error": "運費不能為負數"})
    except ValueError:
        return jsonify({"success": False, "error": "運費必須為整數"})

    mutation = """
    mutation metafieldsSet($metafields: [MetafieldsSetInput!]!) {
        metafieldsSet(metafields: $metafields) {
            metafields { key value }
            userErrors { field message }
        }
    }
    """
    variables = {
        "metafields": [{
            "ownerId": customer_gid,
            "namespace": "custom",
            "key": "shipping_rate",
            "type": "single_line_text_field",
            "value": str(rate_val)  # 儲存台幣值
        }]
    }
    try:
        result = shopify_graphql(mutation, variables)
        if "data" in result:
            mutation_result = result["data"].get("metafieldsSet", {})
            user_errors = mutation_result.get("userErrors", [])
            if user_errors:
                return jsonify({"success": False, "error": "; ".join([e["message"] for e in user_errors])})
            if mutation_result.get("metafields"):
                return jsonify({
                    "success": True,
                    "shipping_rate_twd": rate_val,
                    "shipping_rate_jpy": twd_to_jpy(rate_val)
                })
        if "errors" in result:
            return jsonify({"success": False, "error": str(result["errors"])})
        return jsonify({"success": False, "error": "設定失敗，請重試"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


# ============ 管理員：到貨包裹管理 ============

@app.route("/api/admin/packages", methods=["GET"])
def admin_list_packages():
    g_code = request.args.get("g_code", "")
    conn = get_db()
    if g_code:
        rows = conn.execute(
            "SELECT * FROM packages WHERE g_code=? ORDER BY id DESC", (g_code.upper(),)
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM packages ORDER BY id DESC"
        ).fetchall()
    conn.close()
    return jsonify({"success": True, "packages": [dict(r) for r in rows]})


@app.route("/api/admin/packages", methods=["POST"])
def admin_add_package():
    data = request.json
    g_code      = (data.get("g_code") or "").strip().upper()
    logis_num   = (data.get("logis_num") or "").strip()
    product_name= (data.get("product_name") or "").strip()
    weight      = (data.get("weight") or "").strip()
    note        = (data.get("note") or "").strip()
    status      = data.get("status", "已到貨")

    if not g_code:
        return jsonify({"success": False, "error": "請輸入客戶編號"})
    if not g_code.startswith("G"):
        g_code = "G" + g_code

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today = datetime.now().strftime("%Y-%m-%d")

    conn = get_db()
    cur = conn.execute(
        """INSERT INTO packages (g_code, logis_num, product_name, weight, status, note, in_date, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (g_code, logis_num, product_name, weight, status, note, today, now)
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return jsonify({"success": True, "id": new_id})


@app.route("/api/admin/packages/<int:pkg_id>", methods=["PUT"])
def admin_update_package(pkg_id):
    data = request.json
    fields = []
    values = []
    for key in ["g_code", "logis_num", "product_name", "weight", "status", "note", "in_date"]:
        if key in data:
            val = data[key]
            if key == "g_code":
                val = val.strip().upper()
                if not val.startswith("G"):
                    val = "G" + val
            fields.append(f"{key}=?")
            values.append(val)
    if not fields:
        return jsonify({"success": False, "error": "沒有要更新的欄位"})
    values.append(pkg_id)
    conn = get_db()
    conn.execute(f"UPDATE packages SET {', '.join(fields)} WHERE id=?", values)
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/packages/bulk_ship", methods=["POST"])
def admin_bulk_ship():
    data = request.json
    ids = data.get("ids", [])
    if not ids:
        return jsonify({"success": False, "error": "沒有選取任何包裹"})
    conn = get_db()
    conn.execute(
        f"UPDATE packages SET status='已出貨' WHERE id IN ({','.join(['?']*len(ids))})",
        ids
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "updated": len(ids)})


@app.route("/api/admin/packages/<int:pkg_id>", methods=["DELETE"])
def admin_delete_package(pkg_id):
    conn = get_db()
    conn.execute("DELETE FROM packages WHERE id=?", (pkg_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ============ 客戶端 API ============

@app.route("/api/verify_customer", methods=["POST"])
def verify_customer():
    data = request.json
    g_code = data.get("customer_id", "").strip().upper()
    password = data.get("password", "").strip()
    if not g_code:
        return jsonify({"success": False, "error": "請輸入會員編號"})
    if not password:
        return jsonify({"success": False, "error": "請輸入密碼"})
    if not g_code.startswith("G"):
        g_code = "G" + g_code
    password_clean = normalize_phone(password)

    try:
        customers = get_all_goyoutati_customers()
        for c in customers:
            if c["g_code"] == g_code:
                if c["phone"] and c["phone"] == password_clean:
                    # shipping_rate 現為台幣
                    try:
                        rate_twd = int(c["shipping_rate"]) if c["shipping_rate"] else DEFAULT_SHIPPING_RATE
                    except (ValueError, TypeError):
                        rate_twd = DEFAULT_SHIPPING_RATE
                    rate_jpy = twd_to_jpy(rate_twd) if rate_twd else 0
                    return jsonify({
                        "success": True,
                        "customer": {
                            "id": c["customer_id"],
                            "g_code": g_code,
                            "name": c["name"] or "會員",
                            "email": c["email"],
                            "phone": c["phone"],
                            "phone_raw": c["phone_raw"],
                            "address": c.get("address", ""),
                            "shipping_rate_twd": rate_twd,
                            "shipping_rate_jpy": rate_jpy
                        }
                    })
                else:
                    return jsonify({"success": False, "error": "密碼錯誤，請輸入您的手機號碼"})
        return jsonify({"success": False, "error": "找不到此會員編號，請確認後重試"})
    except Exception as e:
        return jsonify({"success": False, "error": f"查詢失敗: {str(e)}"})


@app.route("/api/forecast", methods=["POST"])
def create_forecast():
    data = request.json
    customer_id = data.get("customer_id")
    g_code = data.get("g_code", "")
    packages = data.get("packages", [])

    if not customer_id:
        return jsonify({"success": False, "error": "缺少客戶編號"})
    if not packages:
        return jsonify({"success": False, "error": "請至少填寫一個包裹"})

    results = []
    for idx, pkg in enumerate(packages):
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        local_logis_num = f"{g_code}-{timestamp}-{idx+1}"
        declare_list = []
        for item in pkg.get("items", []):
            declare_list.append({
                "product_name": item.get("name", "商品"),
                "product_name_local": item.get("name", "商品"),
                "product_num": int(item.get("quantity", 1)),
                "product_price": int(float(item.get("price", 0))),
                "product_url": item.get("url", "")
            })
        total_num = sum(int(item.get("quantity", 1)) for item in pkg.get("items", []))
        total_price = sum(int(float(item.get("price", 0))) * int(item.get("quantity", 1)) for item in pkg.get("items", []))
        forecast_data = {
            "packages": [{
                "local_logis_num": local_logis_num,
                "client_cid": g_code,
                "client_pid": pkg.get("client_pid") or local_logis_num,
                "client_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "warehouse_id": JPD_WAREHOUSE_ID,
                "product_name": declare_list[0]["product_name"] if declare_list else "商品",
                "product_num": total_num,
                "product_price": total_price,
                "declare_list": declare_list
            }]
        }
        result = jpd_request("TForecastPackage", forecast_data)
        if "OperationResult" in result:
            op_result = result["OperationResult"]
            if op_result["Request"]["IsValid"] == "True":
                result_data = op_result.get("Result", {})
                if result_data.get("Result") == "SUCCESS":
                    pkg_data = result_data.get("Data", [{}])[0]
                    results.append({
                        "success": True,
                        "local_logis_num": local_logis_num,
                        "package_id": pkg_data.get("package_id"),
                        "message": pkg_data.get("msg", "預報成功")
                    })
                    continue
        results.append({"success": False, "local_logis_num": local_logis_num, "error": "預報失敗"})

    return jsonify({"success": all(r["success"] for r in results), "results": results})


@app.route("/api/packages", methods=["GET"])
def get_packages():
    g_code = request.args.get("g_code") or request.args.get("customer_id")
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})

    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM packages WHERE g_code=? ORDER BY id DESC",
        (g_code.upper(),)
    ).fetchall()
    conn.close()

    packages = []
    for row in rows:
        r = dict(row)
        packages.append({
            "id":           r["id"],
            "logis_num":    r["logis_num"] or "-",
            "product_name": r["product_name"] or "-",
            "weight":       r["weight"] or "",
            "status":       r["status"],
            "note":         r["note"] or "",
            "in_date":      r["in_date"] or "",
            "created_at":   r["created_at"],
        })
    return jsonify({"success": True, "packages": packages})


@app.route("/api/orders", methods=["GET"])
def get_orders():
    g_code = request.args.get("g_code") or request.args.get("customer_id")
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    result = jpd_request("TSearchOrders", {
        "client_cid": g_code,
        "warehouse_id": JPD_WAREHOUSE_ID
    })
    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            orders = op_result.get("Result", {}).get("Data", [])
            formatted = []
            for order in orders:
                formatted.append({
                    "order_id": order.get("order_id"),
                    "customer_order_id": order.get("customer_order_id"),
                    "logis_num": order.get("logis_num"),
                    "status": order.get("status_name"),
                    "recipient": order.get("recipient"),
                    "create_date": order.get("create_date"),
                    "weight": order.get("weight"),
                    "deliv_fee": order.get("deliv_fee")
                })
            return jsonify({"success": True, "orders": formatted})
    return jsonify({"success": False, "error": "查詢失敗"})


# ============ 統計 API ============

@app.route("/api/admin/stats/monthly/detail", methods=["GET"])
def admin_monthly_detail():
    """取得指定月份的出貨明細"""
    month = request.args.get("month", "")
    if not month:
        return jsonify({"success": False, "error": "缺少月份"})
    try:
        conn = get_db()
        rows = conn.execute("""
            SELECT * FROM shipment_requests
            WHERE status='已出貨' AND total_fee > 0
            ORDER BY updated_at ASC
        """).fetchall()
        conn.close()
        details = []
        for r in rows:
            rd = dict(r)
            date_str = rd.get("updated_at") or rd.get("created_at") or ""
            if date_str[:7] != month:
                continue
            extras = []
            try:
                extras = json.loads(rd.get("extra_services") or "[]")
            except:
                pass
            details.append({
                "date": date_str[:10],
                "g_code": rd.get("g_code", ""),
                "customer_name": rd.get("customer_name", ""),
                "ship_recipient": rd.get("ship_recipient", ""),
                "ship_phone": rd.get("ship_phone", ""),
                "ship_address": rd.get("ship_address", ""),
                "billed_weight": float(rd.get("billed_weight") or 0),
                "rate_per_kg": float(rd.get("rate_per_kg") or 0),
                "shipping_fee": float(rd.get("shipping_fee") or 0),
                "handling_fee": float(rd.get("handling_fee") or 0),
                "extra_services": extras,
                "total_fee": float(rd.get("total_fee") or 0),
            })
        return jsonify({"success": True, "details": details})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/admin/stats/monthly/excel", methods=["GET"])
def admin_monthly_excel():
    """下載指定月份的出貨明細 Excel"""
    month = request.args.get("month", "")  # e.g. "2026-04"
    if not month:
        return jsonify({"success": False, "error": "缺少月份參數"})
    try:
        conn = get_db()
        rows = conn.execute("""
            SELECT * FROM shipment_requests
            WHERE status='已出貨' AND total_fee > 0
            ORDER BY updated_at ASC
        """).fetchall()
        conn.close()

        # 篩選指定月份
        filtered = []
        for r in rows:
            rd = dict(r)
            date_str = rd.get("updated_at") or rd.get("created_at") or ""
            if date_str[:7] == month:
                filtered.append(rd)

        wb = Workbook()
        ws = wb.active
        ws.title = f"{month} 出貨明細"

        # 標題樣式
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )

        headers = ["出貨日期", "客戶編號", "客戶姓名", "寄送地址", "計費重量(kg)",
                    "運費單價", "運費小計", "理貨費", "加值服務明細", "加值服務小計", "合計(台幣)"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        total_kg = 0
        total_shipping = 0
        total_handling = 0
        total_extra = 0
        total_all = 0

        for i, r in enumerate(filtered, 2):
            date_str = r["updated_at"] or r["created_at"] or ""
            bw = float(r["billed_weight"] or 0)
            rate = float(r["rate_per_kg"] or 0)
            sf = float(r["shipping_fee"] or 0)
            hf = float(r["handling_fee"] or 0)
            tf = float(r["total_fee"] or 0)

            # 加值服務
            extra_desc = ""
            extra_total = 0
            try:
                extras = json.loads(r["extra_services"] or "[]")
                parts = []
                for e in extras:
                    qty = int(e.get("qty", 1))
                    price = int(e.get("price", 0))
                    sub = int(e.get("subtotal", qty * price))
                    parts.append(f"{e.get('name','')} ×{qty} = NT${sub}")
                    extra_total += sub
                extra_desc = " / ".join(parts)
            except:
                pass

            ship_addr = " ".join(filter(None, [str(r.get("ship_recipient") or ""), str(r.get("ship_phone") or ""), str(r.get("ship_address") or "")]))

            ws.cell(row=i, column=1, value=date_str[:10]).border = thin_border
            ws.cell(row=i, column=2, value=r["g_code"]).border = thin_border
            ws.cell(row=i, column=3, value=str(r.get("customer_name") or "")).border = thin_border
            ws.cell(row=i, column=4, value=ship_addr).border = thin_border
            ws.cell(row=i, column=5, value=bw).border = thin_border
            ws.cell(row=i, column=6, value=rate).border = thin_border
            ws.cell(row=i, column=7, value=sf).border = thin_border
            ws.cell(row=i, column=8, value=hf).border = thin_border
            ws.cell(row=i, column=9, value=extra_desc).border = thin_border
            ws.cell(row=i, column=10, value=extra_total).border = thin_border
            ws.cell(row=i, column=11, value=tf).border = thin_border

            total_kg += bw
            total_shipping += sf
            total_handling += hf
            total_extra += extra_total
            total_all += tf

        # 合計列
        sum_row = len(filtered) + 2
        sum_font = Font(bold=True, size=11)
        sum_fill = PatternFill(start_color="F39C12", end_color="F39C12", fill_type="solid")
        ws.cell(row=sum_row, column=1, value="合計").font = sum_font
        ws.cell(row=sum_row, column=1).fill = sum_fill
        ws.cell(row=sum_row, column=1).border = thin_border
        for c in range(2, 12):
            ws.cell(row=sum_row, column=c).border = thin_border
            ws.cell(row=sum_row, column=c).font = sum_font
        ws.cell(row=sum_row, column=2, value=f"{len(filtered)} 筆")
        ws.cell(row=sum_row, column=5, value=total_kg)
        ws.cell(row=sum_row, column=7, value=total_shipping)
        ws.cell(row=sum_row, column=8, value=total_handling)
        ws.cell(row=sum_row, column=10, value=total_extra)
        ws.cell(row=sum_row, column=11, value=total_all)

        # 欄寬
        widths = {'A':12, 'B':10, 'C':12, 'D':30, 'E':12, 'F':10, 'G':12, 'H':10, 'I':30, 'J':12, 'K':12}
        for col_letter, w in widths.items():
            ws.column_dimensions[col_letter].width = w

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        filename = f"GOYOUTATI_{month}_出貨明細.xlsx"
        return send_file(buf, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/admin/stats/monthly", methods=["GET"])
def admin_monthly_stats():
    """月報統計：每月出貨公斤數、運費、理貨費、加值服務、總收入"""
    try:
        conn = get_db()
        rows = conn.execute("""
            SELECT * FROM shipment_requests
            WHERE status='已出貨' AND total_fee > 0
            ORDER BY updated_at DESC
        """).fetchall()
        conn.close()

        monthly = {}
        for r in rows:
            date_str = r["updated_at"] or r["created_at"] or ""
            if not date_str:
                continue
            month_key = date_str[:7]  # "2026-04"

            if month_key not in monthly:
                monthly[month_key] = {
                    "month": month_key,
                    "shipments": 0,
                    "total_kg": 0,
                    "shipping_fee": 0,
                    "handling_fee": 0,
                    "extra_fee": 0,
                    "total_revenue": 0,
                    "customers": set()
                }

            m = monthly[month_key]
            m["shipments"] += 1
            m["total_kg"] += float(r["billed_weight"] or 0)
            m["shipping_fee"] += float(r["shipping_fee"] or 0)
            m["handling_fee"] += float(r["handling_fee"] or 0)
            m["total_revenue"] += float(r["total_fee"] or 0)
            m["customers"].add(r["g_code"])

            # 加值服務小計
            try:
                extras = json.loads(r["extra_services"] or "[]")
                for e in extras:
                    m["extra_fee"] += int(e.get("subtotal") or e.get("qty", 1) * e.get("price", 0) or 0)
            except:
                pass

        # set 轉 count
        result = []
        for key in sorted(monthly.keys(), reverse=True):
            m = monthly[key]
            m["customer_count"] = len(m["customers"])
            del m["customers"]
            result.append(m)

        return jsonify({"success": True, "monthly": result})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


# ============ 公告 API ============

@app.route("/api/announcements", methods=["GET"])
def get_announcements():
    """取得啟用中的公告"""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM announcements WHERE is_active=1 ORDER BY id DESC"
    ).fetchall()
    conn.close()
    return jsonify({"success": True, "announcements": [dict(r) for r in rows]})


@app.route("/api/admin/announcements", methods=["GET"])
def admin_get_announcements():
    """管理員取得所有公告"""
    conn = get_db()
    rows = conn.execute("SELECT * FROM announcements ORDER BY id DESC").fetchall()
    conn.close()
    return jsonify({"success": True, "announcements": [dict(r) for r in rows]})


@app.route("/api/admin/announcements", methods=["POST"])
def admin_create_announcement():
    """管理員新增公告"""
    data = request.json
    title = (data.get("title") or "").strip()
    content = (data.get("content") or "").strip()
    if not title or not content:
        return jsonify({"success": False, "error": "標題和內容為必填"})
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    cur = conn.execute(
        "INSERT INTO announcements (title, content, is_active, created_at) VALUES (?, ?, 1, ?)",
        (title, content, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "id": cur.lastrowid})


@app.route("/api/admin/announcements/<int:ann_id>", methods=["PUT"])
def admin_update_announcement(ann_id):
    """管理員更新公告"""
    data = request.json
    conn = get_db()
    fields = {}
    for key in ["title", "content"]:
        if key in data:
            fields[key] = (data[key] or "").strip()
    if "is_active" in data:
        fields["is_active"] = 1 if data["is_active"] else 0
    if fields:
        sets = ", ".join(f"{k}=?" for k in fields)
        vals = list(fields.values()) + [ann_id]
        conn.execute(f"UPDATE announcements SET {sets} WHERE id=?", vals)
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/announcements/<int:ann_id>", methods=["DELETE"])
def admin_delete_announcement(ann_id):
    """管理員刪除公告"""
    conn = get_db()
    conn.execute("DELETE FROM announcements WHERE id=?", (ann_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ============ 地址簿 API ============

@app.route("/api/addresses", methods=["GET"])
def get_addresses():
    """取得客戶地址簿"""
    g_code = request.args.get("g_code", "").upper()
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM addresses WHERE g_code=? ORDER BY is_default DESC, id DESC", (g_code,)
    ).fetchall()
    conn.close()
    return jsonify({"success": True, "addresses": [dict(r) for r in rows]})


@app.route("/api/addresses", methods=["POST"])
def add_address():
    """新增地址"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    recipient = (data.get("recipient") or "").strip()
    phone = (data.get("phone") or "").strip()
    address = (data.get("address") or "").strip()
    label = (data.get("label") or "").strip()
    zipcode = (data.get("zipcode") or "").strip()
    is_default = 1 if data.get("is_default") else 0

    if not g_code or not recipient or not phone or not address:
        return jsonify({"success": False, "error": "收件人、電話、地址為必填"})

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    # 如果設為預設，先清除其他預設
    if is_default:
        conn.execute("UPDATE addresses SET is_default=0 WHERE g_code=?", (g_code,))
    # 如果是第一筆，自動設為預設
    count = conn.execute("SELECT COUNT(*) as c FROM addresses WHERE g_code=?", (g_code,)).fetchone()["c"]
    if count == 0:
        is_default = 1

    conn.execute(
        """INSERT INTO addresses (g_code, label, recipient, phone, zipcode, address, is_default, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (g_code, label, recipient, phone, zipcode, address, is_default, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "地址已新增"})


@app.route("/api/addresses/<int:addr_id>", methods=["PUT"])
def update_address(addr_id):
    """更新地址"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    conn = get_db()
    # 驗證是本人的
    row = conn.execute("SELECT g_code FROM addresses WHERE id=?", (addr_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該地址"})

    fields = {}
    for key in ["label", "recipient", "phone", "zipcode", "address"]:
        if key in data:
            fields[key] = (data[key] or "").strip()
    if "is_default" in data and data["is_default"]:
        conn.execute("UPDATE addresses SET is_default=0 WHERE g_code=?", (g_code,))
        fields["is_default"] = 1

    if fields:
        sets = ", ".join(f"{k}=?" for k in fields)
        vals = list(fields.values()) + [addr_id]
        conn.execute(f"UPDATE addresses SET {sets} WHERE id=?", vals)
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/addresses/<int:addr_id>", methods=["DELETE"])
def delete_address(addr_id):
    """刪除地址"""
    data = request.json or {}
    g_code = (data.get("g_code") or request.args.get("g_code", "")).strip().upper()
    conn = get_db()
    row = conn.execute("SELECT g_code, is_default FROM addresses WHERE id=?", (addr_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該地址"})
    conn.execute("DELETE FROM addresses WHERE id=?", (addr_id,))
    # 如果刪的是預設，把第一筆設為預設
    if row["is_default"]:
        first = conn.execute("SELECT id FROM addresses WHERE g_code=? ORDER BY id LIMIT 1", (g_code,)).fetchone()
        if first:
            conn.execute("UPDATE addresses SET is_default=1 WHERE id=?", (first["id"],))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/addresses/<int:addr_id>/default", methods=["POST"])
def set_default_address(addr_id):
    """設為預設地址"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    conn = get_db()
    row = conn.execute("SELECT g_code FROM addresses WHERE id=?", (addr_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該地址"})
    conn.execute("UPDATE addresses SET is_default=0 WHERE g_code=?", (g_code,))
    conn.execute("UPDATE addresses SET is_default=1 WHERE id=?", (addr_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ============ 出貨申請 API ============

@app.route("/api/shipment_request", methods=["POST"])
def create_shipment_request():
    """客戶申請出貨"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    customer_name = data.get("customer_name", "")
    package_ids = data.get("package_ids", [])
    note = (data.get("note") or "").strip()
    # 收件地址
    ship_recipient = (data.get("ship_recipient") or "").strip()
    ship_phone = (data.get("ship_phone") or "").strip()
    ship_address = (data.get("ship_address") or "").strip()

    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    if not package_ids:
        return jsonify({"success": False, "error": "請選擇要出貨的包裹"})
    if not ship_recipient or not ship_phone or not ship_address:
        return jsonify({"success": False, "error": "請選擇寄送地址"})

    # 組合包裹摘要
    conn = get_db()
    placeholders = ",".join(["?"] * len(package_ids))
    rows = conn.execute(
        f"SELECT id, logis_num, product_name, weight FROM packages WHERE id IN ({placeholders}) AND g_code=?",
        package_ids + [g_code]
    ).fetchall()

    if not rows:
        conn.close()
        return jsonify({"success": False, "error": "找不到對應的包裹"})

    summary_parts = []
    total_weight = 0
    for idx, r in enumerate(rows, 1):
        r = dict(r)
        name = r["product_name"] or "商品"
        logis = r["logis_num"] or ""
        w = r["weight"] or ""
        line = f"{idx}. {name}"
        if w:
            line += f" / {w} kg"
        if logis and logis != "-":
            line += f" / {logis}"
        summary_parts.append(line)
        try:
            total_weight += float(w) if w else 0
        except:
            pass
    summary = "\n".join(summary_parts)
    if total_weight > 0:
        summary += f"\n合計約 {total_weight:.1f} kg"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ids_str = ",".join(str(i) for i in package_ids)

    conn.execute(
        """INSERT INTO shipment_requests (g_code, customer_name, package_ids, package_summary, status, note, ship_recipient, ship_phone, ship_address, created_at)
           VALUES (?, ?, ?, ?, '待處理', ?, ?, ?, ?, ?)""",
        (g_code, customer_name, ids_str, summary, note, ship_recipient, ship_phone, ship_address, now)
    )
    conn.commit()
    conn.close()

    return jsonify({"success": True, "message": "出貨申請已送出，管理員會盡快處理！"})


@app.route("/api/shipment_requests", methods=["GET"])
def get_my_shipment_requests():
    """客戶查看自己的出貨申請"""
    g_code = request.args.get("g_code", "").upper()
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM shipment_requests WHERE g_code=? ORDER BY id DESC", (g_code,)
    ).fetchall()
    conn.close()
    return jsonify({"success": True, "requests": [dict(r) for r in rows]})


@app.route("/api/shipment_requests/<int:req_id>/payment", methods=["POST"])
def submit_payment_info(req_id):
    """客戶回報匯款後五碼"""
    data = request.json
    last5 = (data.get("last5") or "").strip()
    g_code = (data.get("g_code") or "").strip().upper()

    if not last5 or len(last5) != 5:
        return jsonify({"success": False, "error": "請輸入帳號後五碼（5位數字）"})
    if not last5.isdigit():
        return jsonify({"success": False, "error": "請輸入數字"})

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    # 確認是該客戶的申請
    row = conn.execute("SELECT g_code FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
    if not row or row["g_code"] != g_code:
        conn.close()
        return jsonify({"success": False, "error": "找不到該申請"})

    conn.execute(
        "UPDATE shipment_requests SET payment_last5=?, payment_at=? WHERE id=?",
        (last5, now, req_id)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "匯款回報成功！"})


@app.route("/api/admin/shipment_requests", methods=["GET"])
def admin_get_shipment_requests():
    """管理員查看所有出貨申請"""
    status = request.args.get("status", "")
    try:
        conn = get_db()
        if status:
            rows = conn.execute(
                "SELECT * FROM shipment_requests WHERE status=? ORDER BY id DESC", (status,)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM shipment_requests ORDER BY id DESC"
            ).fetchall()
        conn.close()
        return jsonify({"success": True, "requests": [dict(r) for r in rows]})
    except Exception as e:
        return jsonify({"success": False, "error": str(e), "requests": []})


@app.route("/api/admin/shipment_requests/<int:req_id>", methods=["PUT"])
def admin_update_shipment_request(req_id):
    """管理員更新出貨申請狀態（含帳單資訊）"""
    data = request.json
    status = data.get("status", "")
    admin_note = data.get("admin_note", "")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db()

    # 帳單欄位（出貨時填寫）
    billed_weight = data.get("billed_weight", 0)
    rate_per_kg = data.get("rate_per_kg", 0)
    shipping_fee = data.get("shipping_fee", 0)
    handling_fee = data.get("handling_fee", 0)
    total_fee = data.get("total_fee", 0)
    tracking_num = data.get("tracking_num", "")
    extra_services = json.dumps(data.get("extra_services", []), ensure_ascii=False)

    if status == "已出貨" and billed_weight:
        conn.execute(
            """UPDATE shipment_requests 
               SET status=?, admin_note=?, updated_at=?,
                   billed_weight=?, rate_per_kg=?, shipping_fee=?, handling_fee=?, total_fee=?,
                   tracking_num=?, extra_services=?
               WHERE id=?""",
            (status, admin_note, now, billed_weight, rate_per_kg, shipping_fee, handling_fee, total_fee, tracking_num, extra_services, req_id)
        )
    else:
        conn.execute(
            "UPDATE shipment_requests SET status=?, admin_note=?, updated_at=? WHERE id=?",
            (status, admin_note, now, req_id)
        )

    # 如果管理員標記為「已出貨」，同步更新包裹狀態
    if status == "已出貨":
        req = conn.execute("SELECT package_ids FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
        if req:
            pkg_ids = [int(x.strip()) for x in req["package_ids"].split(",") if x.strip()]
            if pkg_ids:
                placeholders = ",".join(["?"] * len(pkg_ids))
                conn.execute(
                    f"UPDATE packages SET status='已出貨' WHERE id IN ({placeholders})", pkg_ids
                )

    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/shipment_requests/<int:req_id>/revert", methods=["POST"])
def admin_revert_shipment_request(req_id):
    """還原出貨申請：狀態回到待處理，包裹回到已到貨，清空帳單"""
    conn = get_db()
    req = conn.execute("SELECT * FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
    if not req:
        conn.close()
        return jsonify({"success": False, "error": "找不到該申請"})

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        """UPDATE shipment_requests 
           SET status='待處理', updated_at=?,
               billed_weight=0, rate_per_kg=0, shipping_fee=0, handling_fee=0, total_fee=0,
               tracking_num='', payment_last5='', payment_at='', extra_services=''
           WHERE id=?""",
        (now, req_id)
    )

    # 包裹狀態還原為「已到貨」
    pkg_ids_str = req["package_ids"]
    if pkg_ids_str:
        pkg_ids = [int(x.strip()) for x in pkg_ids_str.split(",") if x.strip()]
        if pkg_ids:
            placeholders = ",".join(["?"] * len(pkg_ids))
            conn.execute(
                f"UPDATE packages SET status='已到貨' WHERE id IN ({placeholders})", pkg_ids
            )

    conn.commit()
    conn.close()
    return jsonify({"success": True})


# ============ 預報包裹 API（本地存檔，不連 JPD）============

@app.route("/api/forecast_simple", methods=["POST"])
def create_forecast_simple():
    """客戶提交預報（存到本地 DB）"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    customer_name = data.get("customer_name", "")
    items = data.get("items", [])
    note = (data.get("note") or "").strip()

    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    if not items:
        return jsonify({"success": False, "error": "請至少填寫一個商品"})

    # 過濾空的
    valid_items = [i for i in items if (i.get("name") or "").strip()]
    if not valid_items:
        return jsonify({"success": False, "error": "請至少填寫一個商品名稱"})

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    conn.execute(
        """INSERT INTO forecasts (g_code, customer_name, items_json, status, note, created_at)
           VALUES (?, ?, ?, '待處理', ?, ?)""",
        (g_code, customer_name, json.dumps(valid_items, ensure_ascii=False), note, now)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "預報已送出！我們收到後會盡快處理。"})


@app.route("/api/my_forecasts", methods=["GET"])
def get_my_forecasts():
    """客戶查看自己的預報"""
    g_code = request.args.get("g_code", "").upper()
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM forecasts WHERE g_code=? ORDER BY id DESC LIMIT 20", (g_code,)
    ).fetchall()
    conn.close()
    results = []
    for r in rows:
        row = dict(r)
        try:
            row["items"] = json.loads(row.get("items_json") or "[]")
        except:
            row["items"] = []
        results.append(row)
    return jsonify({"success": True, "forecasts": results})


@app.route("/api/admin/forecasts", methods=["GET"])
def admin_get_forecasts():
    """管理員查看所有預報"""
    status = request.args.get("status", "")
    conn = get_db()
    if status:
        rows = conn.execute(
            "SELECT * FROM forecasts WHERE status=? ORDER BY id DESC", (status,)
        ).fetchall()
    else:
        rows = conn.execute("SELECT * FROM forecasts ORDER BY id DESC").fetchall()
    conn.close()
    results = []
    for r in rows:
        row = dict(r)
        try:
            row["items"] = json.loads(row.get("items_json") or "[]")
        except:
            row["items"] = []
        results.append(row)
    return jsonify({"success": True, "forecasts": results})


@app.route("/api/admin/forecasts/<int:fc_id>", methods=["PUT"])
def admin_update_forecast(fc_id):
    """管理員更新預報狀態"""
    data = request.json
    status = data.get("status", "")
    conn = get_db()
    conn.execute("UPDATE forecasts SET status=? WHERE id=?", (status, fc_id))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/forecasts/<int:fc_id>/excel")
def admin_download_forecast_excel(fc_id):
    """下載單筆預報的 JPD Excel"""
    conn = get_db()
    row = conn.execute("SELECT * FROM forecasts WHERE id=?", (fc_id,)).fetchone()
    conn.close()
    if not row:
        return "Not found", 404
    row = dict(row)
    try:
        items = json.loads(row.get("items_json") or "[]")
    except:
        items = []

    g_code = row["g_code"]
    today_str = datetime.now().strftime("%m%d")
    customer_order_id = f"{g_code}-{today_str}"

    wb = Workbook()
    ws = wb.active
    ws.title = "預報資料"

    # 標頭
    headers = [
        "客戶運單號", "JpD包裹ID", "運單ID", "包裹特殊服務",
        "收件人", "收件人身份證ID", "收件人詳細地址", "收件人电话号码",
        "備註", "特殊服务", "渠道ID",
        "申報人", "申報人身份證ID", "申報人詳細地址", "申報人电话号码",
        "品名", "数量", "金额", "材質", "產地", "URL/JanCode"
    ]
    hfill = PatternFill("solid", fgColor="1F4E79")
    hfont = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = hfill
        cell.font = hfont
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin

    # 資料
    for row_idx, item in enumerate(items, 2):
        data_row = [
            customer_order_id, "", "", "",
            "", "", "", "",
            row.get("note", ""), "", "40",
            "", "", "", "",
            item.get("name", ""),
            item.get("quantity", 1),
            item.get("price", 0),
            "", "Japan",
            item.get("url", "")
        ]
        for col_idx, val in enumerate(data_row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.border = thin
            cell.alignment = Alignment(vertical="center")

    # 欄寬
    col_widths = {1:16, 2:14, 5:12, 7:20, 8:16, 9:12, 11:8, 16:20, 17:8, 18:10, 21:30}
    for col, w in col_widths.items():
        ws.column_dimensions[chr(64+col) if col<=26 else 'A'].width = w
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"{g_code}_{today_str}_forecast.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print(f"""
    ╔═══════════════════════════════════════════════════════════╗
    ║       客人集運預報系統                                      ║
    ║       御用達 × JPD 雲倉                                     ║
    ╚═══════════════════════════════════════════════════════════╝
    🌐 服務啟動於 Port: {port}
    💱 TWD → JPY 匯率: {TWD_TO_JPY_RATE}
    """)
    app.run(host="0.0.0.0", port=port, debug=debug)
