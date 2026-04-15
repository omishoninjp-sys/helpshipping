"""
客人集運預報系統
GOYOUTATI x OMISHONIN 雲倉
"""

from flask import Flask, request, jsonify, render_template, make_response
from datetime import datetime
import requests
import json
import os
import sqlite3
import csv
import io

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
    # 帳單欄位遷移（已存在的表加欄位）
    for col, default in [
        ("billed_weight", "0"), ("rate_per_kg", "0"),
        ("shipping_fee", "0"), ("handling_fee", "0"), ("total_fee", "0"),
        ("payment_last5", "''"), ("payment_at", "''")
    ]:
        try:
            conn.execute(f"ALTER TABLE shipment_requests ADD COLUMN {col} REAL DEFAULT {default}")
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


def get_all_goyoutati_customers():
    graphql_query = """
    {
        metafieldDefinitions(first: 1, ownerType: CUSTOMER, namespace: "custom", key: "goyoutati_id") {
            edges {
                node {
                    id
                    name
                    metafieldsCount
                    metafields(first: 100) {
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
                                        }
                                        createdAt
                                        shippingRate: metafield(namespace: "custom", key: "shipping_rate") {
                                            value
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    """
    result = shopify_graphql(graphql_query)
    customers = []

    if "data" in result:
        definitions = result["data"].get("metafieldDefinitions", {}).get("edges", [])
        if definitions:
            metafields = definitions[0]["node"].get("metafields", {}).get("edges", [])
            for mf in metafields:
                node = mf["node"]
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
                rate_mf = owner.get("shippingRate")
                # shipping_rate 現在儲存台幣值
                shipping_rate_twd = rate_mf["value"] if rate_mf and rate_mf.get("value") else ""
                customers.append({
                    "g_code": g_code,
                    "customer_id": customer_id,
                    "gid": gid,
                    "name": customer_name,
                    "email": owner.get("email", ""),
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
        members = get_all_goyoutati_customers()
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


# ============ 出貨申請 API ============

@app.route("/api/shipment_request", methods=["POST"])
def create_shipment_request():
    """客戶申請出貨"""
    data = request.json
    g_code = (data.get("g_code") or "").strip().upper()
    customer_name = data.get("customer_name", "")
    package_ids = data.get("package_ids", [])
    note = (data.get("note") or "").strip()

    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    if not package_ids:
        return jsonify({"success": False, "error": "請選擇要出貨的包裹"})

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
        """INSERT INTO shipment_requests (g_code, customer_name, package_ids, package_summary, status, note, created_at)
           VALUES (?, ?, ?, ?, '待處理', ?, ?)""",
        (g_code, customer_name, ids_str, summary, note, now)
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

    if status == "已出貨" and billed_weight:
        conn.execute(
            """UPDATE shipment_requests 
               SET status=?, admin_note=?, updated_at=?,
                   billed_weight=?, rate_per_kg=?, shipping_fee=?, handling_fee=?, total_fee=?
               WHERE id=?""",
            (status, admin_note, now, billed_weight, rate_per_kg, shipping_fee, handling_fee, total_fee, req_id)
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


@app.route("/api/admin/forecasts/<int:fc_id>/csv")
def admin_download_forecast_csv(fc_id):
    """下載單筆預報的 JPD CSV"""
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

    output = io.StringIO()
    writer = csv.writer(output)
    # JPD CSV 標頭
    writer.writerow([
        "客戶運單號", "JpD包裹ID", "運單ID", "包裹特殊服務",
        "收件人", "收件人身份證ID", "收件人詳細地址", "收件人电话号码",
        "備註", "特殊服务", "渠道ID",
        "申報人", "申報人身份證ID", "申報人詳細地址", "申報人电话号码",
        "品名", "数量", "金额", "材質", "產地", "URL/JanCode"
    ])
    for item in items:
        writer.writerow([
            customer_order_id,  # 客戶運單號
            "",  # JpD包裹ID（手動填）
            "",  # 運單ID
            "",  # 包裹特殊服務
            "",  # 收件人（手動填）
            "",  # 身份證
            "",  # 地址（手動填）
            "",  # 電話（手動填）
            row.get("note", ""),  # 備註
            "",  # 特殊服务
            "40",  # 渠道ID
            "",  # 申報人
            "",  # 申報人身份證
            "",  # 申報人地址
            "",  # 申報人電話
            item.get("name", ""),  # 品名
            item.get("quantity", 1),  # 数量
            item.get("price", 0),  # 金额
            "",  # 材質
            "Japan",  # 產地
            item.get("url", ""),  # URL
        ])

    resp = make_response(output.getvalue())
    resp.headers["Content-Type"] = "text/csv; charset=utf-8-sig"
    resp.headers["Content-Disposition"] = f"attachment; filename={g_code}_{today_str}_forecast.csv"
    return resp


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
