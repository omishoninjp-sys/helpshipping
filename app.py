"""
客人集運預報系統
GOYOUTATI x OMISHONIN 雲倉
"""

from flask import Flask, request, jsonify, render_template, make_response, send_file, session
from datetime import datetime, timedelta
import requests
import json
import os
import sqlite3
import csv
import io
import time
import re
import secrets
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
# Session 設定（環境變數 SESSION_SECRET 沒設就用隨機值，每次重啟會失效但不會暴露 fallback）
app.secret_key = os.environ.get("SESSION_SECRET") or secrets.token_hex(32)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)

# ============ 設定區（從環境變數讀取）============
JPD_BASE_URL = "https://biz.cloudwh.jp"
JPD_EMAIL = os.environ.get("JPD_EMAIL", "omishoninjp@gmail.com")
JPD_PASSWORD = os.environ.get("JPD_PASSWORD", "omi0131")
JPD_WAREHOUSE_ID = int(os.environ.get("JPD_WAREHOUSE_ID", "1"))
JPD_DELIV_ID = int(os.environ.get("JPD_DELIV_ID", "40"))  # 台灣空運線

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
    conn.execute("""
        CREATE TABLE IF NOT EXISTS admin_users (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            username    TEXT    UNIQUE NOT NULL,
            password    TEXT    NOT NULL,
            role        TEXT    DEFAULT 'admin',
            created_at  TEXT    NOT NULL
        )
    """)
    # ===== 代理帳號表（Phase 1）=====
    conn.execute("""
        CREATE TABLE IF NOT EXISTS agents (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            username        TEXT UNIQUE NOT NULL,
            password        TEXT NOT NULL,
            prefix          TEXT UNIQUE NOT NULL,
            name            TEXT NOT NULL,
            min_rate        REAL DEFAULT 180,
            contact_phone   TEXT DEFAULT '',
            contact_email   TEXT DEFAULT '',
            status          TEXT DEFAULT 'active',
            note            TEXT DEFAULT '',
            created_at      TEXT NOT NULL
        )
    """)
    # ===== 會員表（代理建的客戶，存本地；你自己的客戶仍走 Shopify）=====
    conn.execute("""
        CREATE TABLE IF NOT EXISTS members (
            g_code          TEXT PRIMARY KEY,
            agent_id        INTEGER NOT NULL,
            name            TEXT NOT NULL,
            password        TEXT DEFAULT '',
            phone           TEXT DEFAULT '',
            address         TEXT DEFAULT '',
            line_id         TEXT DEFAULT '',
            email           TEXT DEFAULT '',
            note            TEXT DEFAULT '',
            status          TEXT DEFAULT 'active',
            created_at      TEXT NOT NULL
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_members_agent ON members(agent_id)")
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
        ("consolidation_fee", "REAL", "0"),
    ]:
        try:
            conn.execute(f"ALTER TABLE shipment_requests ADD COLUMN {col} {col_type} DEFAULT {default}")
        except:
            pass

    # ===== Phase 2: 加 agent_id 欄位（既有資料預設 0 = 主管理員的）=====
    for table in ["packages", "forecasts", "shipment_requests", "announcements"]:
        try:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN agent_id INTEGER DEFAULT 0")
            print(f"[migrate] 已加 {table}.agent_id 欄位", flush=True)
        except:
            pass
    # 索引加速 agent 過濾
    for table in ["packages", "forecasts", "shipment_requests"]:
        try:
            conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table}_agent ON {table}(agent_id)")
        except:
            pass

    # ===== Phase 3+: 加 members.shipping_rate 欄位（代理為每個客戶設定獨立費率）=====
    # 預設 0 = 沿用該代理的 min_rate；>0 = 該會員的專屬費率
    try:
        conn.execute("ALTER TABLE members ADD COLUMN shipping_rate REAL DEFAULT 0")
        print("[migrate] 已加 members.shipping_rate 欄位", flush=True)
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
        response = requests.post(graphql_url, headers=headers, json=payload, timeout=15)
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
_customers_cache = {"data": None, "time": 0, "loading": False}
CACHE_TTL = 600  # 10 分鐘


def get_all_goyoutati_customers(force_refresh=False):
    global _customers_cache
    now = time.time()
    # 快取有效就直接回
    if not force_refresh and _customers_cache["data"] is not None and (now - _customers_cache["time"]) < CACHE_TTL:
        return _customers_cache["data"]

    # 防止重複拉取
    if _customers_cache["loading"]:
        return _customers_cache["data"] or []
    
    _customers_cache["loading"] = True
    try:
        print("[Shopify] 🔄 開始拉取會員資料...", flush=True)
        customers = _fetch_customers_from_shopify()
        if customers:
            _customers_cache = {"data": customers, "time": time.time(), "loading": False}
            print(f"[Shopify] ✅ 拉取完成，共 {len(customers)} 位會員", flush=True)
        else:
            _customers_cache["loading"] = False
            if _customers_cache["data"] is not None:
                return _customers_cache["data"]
        return customers
    except Exception as e:
        print(f"[Shopify] ❌ 拉取失敗: {e}", flush=True)
        _customers_cache["loading"] = False
        return _customers_cache["data"] or []


def _fetch_customers_from_shopify():
    customers = []
    cursor = None
    has_next = True
    page = 0

    while has_next and page < 10:  # 最多 10 頁 = 1000 會員
        page += 1
        after_arg = f', after: "{cursor}"' if cursor else ''
        graphql_query = '{metafieldDefinitions(first:1,ownerType:CUSTOMER,namespace:"custom",key:"goyoutati_id"){edges{node{id metafields(first:100' + after_arg + '){edges{node{value owner{...on Customer{id firstName lastName email phone defaultAddress{phone province city address1 address2} createdAt shippingRate:metafield(namespace:"custom",key:"shipping_rate"){value}}}} cursor} pageInfo{hasNextPage}}}}}}'

        print(f"[Shopify] Fetching customers page {page}, cursor={cursor}", flush=True)
        result = shopify_graphql(graphql_query)
        has_next = False

        if "data" not in result:
            print(f"[Shopify] Error: {result}", flush=True)
            break

        definitions = result["data"].get("metafieldDefinitions", {}).get("edges", [])
        if not definitions:
            print("[Shopify] No metafieldDefinitions found", flush=True)
            break

        metafields_data = definitions[0]["node"].get("metafields", {})
        edges = metafields_data.get("edges", [])
        page_info = metafields_data.get("pageInfo", {})
        has_next = page_info.get("hasNextPage", False)
        print(f"[Shopify] Page {page}: got {len(edges)} metafields, hasNextPage={has_next}", flush=True)

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
    """取得管理員密碼：環境變數優先，否則 DB，最後預設"""
    env_pw = os.environ.get("ADMIN_PASSWORD", "")
    if env_pw:
        return env_pw
    conn = get_db()
    row = conn.execute("SELECT value FROM admin_settings WHERE key='admin_password'").fetchone()
    conn.close()
    if row:
        return row["value"]
    return "admin123"


def _ensure_super_admin():
    """確保至少有一個超級管理員"""
    try:
        conn = get_db()
        count = conn.execute("SELECT COUNT(*) as c FROM admin_users").fetchone()["c"]
        if count == 0:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            pwd = get_admin_password()
            conn.execute(
                "INSERT INTO admin_users (username, password, role, created_at) VALUES (?, ?, 'super', ?)",
                ("admin", pwd, now)
            )
            conn.commit()
            print(f"[Admin] ✅ 已建立超級管理員帳號: admin / {pwd}", flush=True)
        else:
            env_pw = os.environ.get("ADMIN_PASSWORD", "")
            if env_pw:
                conn.execute("UPDATE admin_users SET password=? WHERE role='super'", (env_pw,))
                conn.commit()
            print(f"[Admin] ✅ 已有 {count} 個管理員帳號", flush=True)
        conn.close()
    except Exception as e:
        print(f"[Admin] ❌ 初始化失敗: {e}", flush=True)

_ensure_super_admin()
print("[App] ✅ 啟動完成", flush=True)


@app.route("/api/admin/verify", methods=["POST"])
def admin_verify():
    data = request.json
    username = (data.get("username") or "").strip()
    password = data.get("password", "")
    print(f"[Login] 嘗試登入: username='{username}'", flush=True)

    conn = get_db()
    user = None
    user_type = None

    # 1) 先查管理員（admin_users）— 既有行為不變
    if username:
        user = conn.execute(
            "SELECT * FROM admin_users WHERE username=? AND password=?", (username, password)
        ).fetchone()
    else:
        # 相容舊的純密碼登入
        user = conn.execute(
            "SELECT * FROM admin_users WHERE password=?", (password,)
        ).fetchone()
    if user:
        user_type = "admin"

    # 2) admin 找不到 → 再查代理（agents）
    if not user and username:
        user = conn.execute(
            "SELECT * FROM agents WHERE username=? AND password=? AND status='active'",
            (username, password)
        ).fetchone()
        if user:
            user_type = "agent"
    conn.close()

    if user:
        # 寫入 session
        session.permanent = True
        session["user_type"] = user_type
        session["user_id"] = user["id"]
        session["username"] = user["username"]
        if user_type == "admin":
            session["role"] = user["role"]
            session["agent_id"] = 0  # 0 = 主管理員 / 你的員工，看全部
            print(f"[Login] ✅ admin 登入: {user['username']} ({user['role']})", flush=True)
            return jsonify({
                "success": True,
                "user_type": "admin",
                "username": user["username"],
                "role": user["role"]
            })
        else:
            session["role"] = "agent"
            session["agent_id"] = user["id"]
            session["prefix"] = user["prefix"]
            print(f"[Login] ✅ agent 登入: {user['username']} (prefix={user['prefix']}, id={user['id']})", flush=True)
            return jsonify({
                "success": True,
                "user_type": "agent",
                "username": user["username"],
                "name": user["name"],
                "prefix": user["prefix"]
            })

    print(f"[Login] ❌ 登入失敗", flush=True)
    return jsonify({"success": False, "error": "帳號或密碼錯誤"})


# ===== 身份輔助函式 =====
def current_user():
    """回傳當前登入者資訊（從 session）"""
    if "user_type" not in session:
        return None
    return {
        "user_type": session.get("user_type"),
        "user_id": session.get("user_id"),
        "username": session.get("username"),
        "role": session.get("role"),
        "agent_id": session.get("agent_id", 0),
        "prefix": session.get("prefix", "G"),
    }

def is_super_admin():
    """是否為主管理員（admin_users 表的人，看全部）"""
    return session.get("user_type") == "admin"

def get_current_agent_id():
    """當前代理 id（>0 才是代理，0 = 主管理員看全部）"""
    return int(session.get("agent_id", 0))


@app.route("/api/me", methods=["GET"])
def api_me():
    """前端查當前身份"""
    u = current_user()
    if not u:
        return jsonify({"logged_in": False})
    return jsonify({"logged_in": True, **u})


@app.route("/api/admin/logout", methods=["POST"])
def admin_logout():
    session.clear()
    return jsonify({"success": True})


@app.route("/api/admin/change_password", methods=["POST"])
def admin_change_password():
    data = request.json
    username = (data.get("username") or "admin").strip()
    current = data.get("current", "")
    new_pwd = data.get("new_password", "").strip()
    confirm = data.get("confirm", "").strip()

    conn = get_db()
    user = conn.execute("SELECT * FROM admin_users WHERE username=?", (username,)).fetchone()
    if not user or user["password"] != current:
        conn.close()
        return jsonify({"success": False, "error": "目前密碼錯誤"})
    if not new_pwd or len(new_pwd) < 4:
        conn.close()
        return jsonify({"success": False, "error": "新密碼至少 4 個字元"})
    if new_pwd != confirm:
        conn.close()
        return jsonify({"success": False, "error": "兩次密碼不一致"})

    conn.execute("UPDATE admin_users SET password=? WHERE username=?", (new_pwd, username))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "密碼已更新"})


# ── 管理員帳號管理 ──

@app.route("/api/admin/users", methods=["GET"])
def admin_list_users():
    conn = get_db()
    rows = conn.execute("SELECT id, username, role, created_at FROM admin_users ORDER BY id").fetchall()
    conn.close()
    return jsonify({"success": True, "users": [dict(r) for r in rows]})


# ===== 代理帳號管理（只有主管理員可操作）=====

@app.route("/api/admin/agents", methods=["GET"])
def admin_list_agents():
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    rows = conn.execute(
        "SELECT id, username, prefix, name, min_rate, contact_phone, contact_email, status, note, created_at FROM agents ORDER BY id"
    ).fetchall()
    # 順便統計每個代理底下的會員數
    counts = {}
    for r in conn.execute("SELECT agent_id, COUNT(*) as c FROM members GROUP BY agent_id").fetchall():
        counts[r["agent_id"]] = r["c"]
    conn.close()
    result = []
    for r in rows:
        d = dict(r)
        d["member_count"] = counts.get(r["id"], 0)
        result.append(d)
    return jsonify({"success": True, "agents": result})


@app.route("/api/admin/agents", methods=["POST"])
def admin_create_agent():
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    prefix = (data.get("prefix") or "").strip().upper()
    name = (data.get("name") or "").strip()
    min_rate = float(data.get("min_rate") or 180)

    if not username or not password or not prefix or not name:
        return jsonify({"success": False, "error": "帳號、密碼、前綴、名稱皆為必填"})
    if not re.fullmatch(r"[A-Z]", prefix):
        return jsonify({"success": False, "error": "前綴必須為單一英文字母（A-Z）"})
    if prefix == "G":
        return jsonify({"success": False, "error": "前綴 G 已保留給主管理員"})
    if min_rate < 180:
        return jsonify({"success": False, "error": "最低費率不得低於 NT$180/kg"})

    conn = get_db()
    try:
        conn.execute(
            """INSERT INTO agents (username, password, prefix, name, min_rate, contact_phone, contact_email, status, note, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, 'active', ?, ?)""",
            (username, password, prefix, name, min_rate,
             data.get("contact_phone", ""), data.get("contact_email", ""),
             data.get("note", ""), datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()
        new_id = conn.execute("SELECT last_insert_rowid() as id").fetchone()["id"]
    except sqlite3.IntegrityError as e:
        conn.close()
        msg = str(e)
        if "agents.username" in msg:
            return jsonify({"success": False, "error": f"帳號「{username}」已被使用"})
        if "agents.prefix" in msg:
            return jsonify({"success": False, "error": f"前綴「{prefix}」已被使用"})
        return jsonify({"success": False, "error": f"資料庫錯誤：{msg}"})
    conn.close()
    return jsonify({"success": True, "id": new_id, "message": f"代理「{name}」已建立（前綴 {prefix}）"})


@app.route("/api/admin/agents/<int:agent_id>", methods=["PUT"])
def admin_update_agent(agent_id):
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    conn = get_db()
    existing = conn.execute("SELECT * FROM agents WHERE id=?", (agent_id,)).fetchone()
    if not existing:
        conn.close()
        return jsonify({"success": False, "error": "代理不存在"})

    fields = []
    values = []
    # 可改：name, min_rate, contact_phone, contact_email, status, note, password
    if "name" in data:
        fields.append("name=?"); values.append((data["name"] or "").strip())
    if "min_rate" in data:
        try:
            mr = float(data["min_rate"])
            if mr < 180:
                conn.close()
                return jsonify({"success": False, "error": "最低費率不得低於 NT$180/kg"})
            fields.append("min_rate=?"); values.append(mr)
        except (ValueError, TypeError):
            conn.close()
            return jsonify({"success": False, "error": "費率必須為數字"})
    if "contact_phone" in data:
        fields.append("contact_phone=?"); values.append(data["contact_phone"] or "")
    if "contact_email" in data:
        fields.append("contact_email=?"); values.append(data["contact_email"] or "")
    if "status" in data and data["status"] in ("active", "disabled"):
        fields.append("status=?"); values.append(data["status"])
    if "note" in data:
        fields.append("note=?"); values.append(data["note"] or "")
    if data.get("password"):
        fields.append("password=?"); values.append(data["password"])
    # 前綴與帳號名建立後不可改（避免關聯混亂）

    if not fields:
        conn.close()
        return jsonify({"success": False, "error": "沒有可更新的欄位"})
    values.append(agent_id)
    conn.execute(f"UPDATE agents SET {', '.join(fields)} WHERE id=?", values)
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "已更新"})


@app.route("/api/admin/agents/<int:agent_id>", methods=["DELETE"])
def admin_delete_agent(agent_id):
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    # 安全檢查：若已有會員，不允許刪（避免孤兒資料）
    count = conn.execute("SELECT COUNT(*) as c FROM members WHERE agent_id=?", (agent_id,)).fetchone()["c"]
    if count > 0:
        conn.close()
        return jsonify({"success": False, "error": f"此代理底下尚有 {count} 位會員，無法刪除。可改為「停用（disabled）」狀態"})
    conn.execute("DELETE FROM agents WHERE id=?", (agent_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "代理已刪除"})


# ===== 統一會員查詢（本地 members 優先、找不到回退 Shopify）=====
def get_agent_id_for_g_code(g_code):
    """
    依 g_code 找出歸屬的 agent_id。
    - 在 members 表（代理的客戶）→ 回那個代理的 id
    - 不在 members 表（你 Shopify 來的客戶）→ 回 0（主管理員）
    """
    if not g_code:
        return 0
    try:
        conn = get_db()
        row = conn.execute("SELECT agent_id FROM members WHERE g_code=?", (g_code,)).fetchone()
        conn.close()
        if row:
            return int(row["agent_id"] or 0)
    except Exception as e:
        print(f"[get_agent_id_for_g_code] 失敗: {e}", flush=True)
    return 0


def check_record_ownership(table, record_id):
    """
    檢查當前使用者是否可存取該筆紀錄。
    回傳 (allowed: bool, record_dict_or_none)
    - 主管理員：永遠可以
    - 代理：只有當 record.agent_id == 自己的 agent_id 才可以
    """
    aid = get_current_agent_id()
    conn = get_db()
    row = conn.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
    conn.close()
    if not row:
        return False, None
    if aid == 0 or is_super_admin():
        return True, dict(row)
    return (int(row["agent_id"] or 0) == aid), dict(row)


def get_member_unified(g_code):
    """
    回傳 {g_code, name, agent_id, phone, address, source} 或 None
    - source='local'  → 來自代理建的會員（agent_id > 0）
    - source='shopify' → 來自你的 Shopify（agent_id = 0）
    """
    if not g_code:
        return None
    conn = get_db()
    row = conn.execute("SELECT * FROM members WHERE g_code=?", (g_code,)).fetchone()
    conn.close()
    if row:
        d = dict(row)
        d["source"] = "local"
        return d
    # 回退到 Shopify 快取
    try:
        for c in get_all_goyoutati_customers():
            if (c.get("g_code") or "") == g_code:
                return {
                    "g_code": g_code,
                    "name": c.get("name", ""),
                    "phone": c.get("phone", ""),
                    "address": c.get("address", ""),
                    "agent_id": 0,
                    "source": "shopify"
                }
    except Exception as e:
        print(f"[get_member_unified] Shopify 查詢失敗: {e}", flush=True)
    return None


@app.route("/api/admin/users", methods=["POST"])
def admin_create_user():
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    role = data.get("role", "admin")
    if not username or not password:
        return jsonify({"success": False, "error": "帳號和密碼為必填"})
    if len(password) < 4:
        return jsonify({"success": False, "error": "密碼至少 4 個字元"})
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    try:
        conn.execute("INSERT INTO admin_users (username, password, role, created_at) VALUES (?, ?, ?, ?)",
                     (username, password, role, now))
        conn.commit()
        conn.close()
        return jsonify({"success": True})
    except:
        conn.close()
        return jsonify({"success": False, "error": "帳號已存在"})

@app.route("/api/admin/users/<int:user_id>", methods=["DELETE"])
def admin_delete_user(user_id):
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    user = conn.execute("SELECT role FROM admin_users WHERE id=?", (user_id,)).fetchone()
    if not user:
        conn.close()
        return jsonify({"success": False, "error": "找不到"})
    if user["role"] == "super":
        conn.close()
        return jsonify({"success": False, "error": "無法刪除超級管理員"})
    conn.execute("DELETE FROM admin_users WHERE id=?", (user_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/members", methods=["GET"])
def get_all_members():
    try:
        aid = get_current_agent_id()
        # ===== 代理：只看自己本地建的會員 =====
        if aid > 0:
            conn = get_db()
            agent = conn.execute("SELECT prefix, min_rate FROM agents WHERE id=?", (aid,)).fetchone()
            prefix = agent["prefix"] if agent else "X"
            min_rate = float(agent["min_rate"] or 180) if agent else 180.0
            rows = conn.execute(
                "SELECT * FROM members WHERE agent_id=? ORDER BY g_code", (aid,)
            ).fetchall()
            conn.close()
            members = []
            used_numbers = set()
            for r in rows:
                d = dict(r)
                # 會員專屬費率 > 0 → 用該費率；否則 fallback 到代理 min_rate
                member_rate = float(d.get("shipping_rate") or 0)
                effective_rate = member_rate if member_rate > 0 else min_rate
                members.append({
                    "g_code": d.get("g_code", ""),
                    "name": d.get("name", ""),
                    "phone": d.get("phone", ""),
                    "address": d.get("address", ""),
                    "line_id": d.get("line_id", ""),
                    "email": d.get("email", ""),
                    "shipping_rate": effective_rate,
                    "shipping_rate_raw": member_rate,  # 0 表示沿用 min_rate
                    "note": d.get("note", ""),
                    "status": d.get("status", "active"),
                    "source": "local",
                })
                gc = d.get("g_code", "")
                if gc.startswith(prefix):
                    try:
                        used_numbers.add(int(gc[len(prefix):]))
                    except (ValueError, TypeError):
                        pass
            max_number = max(used_numbers) if used_numbers else 0
            next_number = 1
            while next_number in used_numbers:
                next_number += 1
            next_g_code = f"{prefix}{next_number:04d}"
            return jsonify({
                "success": True,
                "members": members,
                "total": len(members),
                "max_number": max_number,
                "next_g_code": next_g_code,
                "default_shipping_rate": min_rate,
                "twd_to_jpy_rate": TWD_TO_JPY_RATE,
                "min_rate": min_rate,
                "prefix": prefix,
                "source": "agent_local",
            })

        # ===== 主管理員：原有 Shopify 行為（不動）=====
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


@app.route("/api/admin/search_members", methods=["GET"])
def admin_search_members():
    """
    跨會員搜尋（給新增到貨等場景使用，倉庫人員打字即時搜尋）
    - 主管理員：搜本地 members 表 + Shopify 兩邊（含所有代理客戶）
    - 代理：只搜自己 members 表的客戶
    - 比對欄位：g_code、name、phone（去空白後）
    """
    q = (request.args.get("q") or "").strip()
    try:
        limit = int(request.args.get("limit", 15))
    except (ValueError, TypeError):
        limit = 15
    limit = max(1, min(limit, 50))

    if not q or len(q) < 1:
        return jsonify({"success": True, "results": []})

    q_upper = q.upper()
    q_phone = normalize_phone(q)
    pattern = f"%{q_upper}%"
    pattern_phone = f"%{q_phone}%"
    aid = get_current_agent_id()
    results = []

    conn = get_db()
    # 本地 members（代理建的客戶）
    if aid > 0:
        local = conn.execute("""
            SELECT g_code, name, phone, address, agent_id, status
            FROM members
            WHERE agent_id=? AND status='active'
              AND (UPPER(g_code) LIKE ? OR UPPER(name) LIKE ? OR phone LIKE ?)
            ORDER BY g_code LIMIT ?
        """, (aid, pattern, pattern, pattern_phone, limit)).fetchall()
    else:
        local = conn.execute("""
            SELECT m.g_code, m.name, m.phone, m.address, m.agent_id, m.status,
                   a.name as agent_name, a.prefix as agent_prefix
            FROM members m
            LEFT JOIN agents a ON a.id = m.agent_id
            WHERE m.status='active'
              AND (UPPER(m.g_code) LIKE ? OR UPPER(m.name) LIKE ? OR m.phone LIKE ?)
            ORDER BY m.g_code LIMIT ?
        """, (pattern, pattern, pattern_phone, limit)).fetchall()
    conn.close()

    for r in local:
        d = dict(r)
        results.append({
            "g_code": d.get("g_code"),
            "name": d.get("name") or "",
            "phone": d.get("phone") or "",
            "address": d.get("address") or "",
            "source": "agent",
            "agent_id": d.get("agent_id") or 0,
            "agent_name": d.get("agent_name") or "",
        })

    # Shopify 客戶（僅主管理員視角）
    if aid == 0:
        try:
            customers = get_all_goyoutati_customers()
            for c in customers:
                if len(results) >= limit:
                    break
                gc = (c.get("g_code") or "").upper()
                nm = (c.get("name") or "").upper()
                ph = c.get("phone") or ""
                if q_upper in gc or q_upper in nm or (q_phone and q_phone in ph):
                    results.append({
                        "g_code": c.get("g_code"),
                        "name": c.get("name") or "",
                        "phone": ph,
                        "address": c.get("address") or "",
                        "source": "shopify",
                        "agent_id": 0,
                        "agent_name": "",
                    })
        except Exception as e:
            print(f"[search_members] Shopify 搜尋失敗：{e}", flush=True)

    return jsonify({"success": True, "results": results[:limit]})


@app.route("/api/admin/members", methods=["POST"])
def admin_create_member():
    """代理新增會員（自動補前綴與 agent_id）。主管理員建議直接在 Shopify 操作。"""
    aid = get_current_agent_id()
    if aid <= 0:
        return jsonify({"success": False, "error": "主管理員的會員請在 Shopify 後台建立"}), 400
    data = request.json or {}
    name = (data.get("name") or "").strip()
    phone = normalize_phone((data.get("phone") or "").strip())
    address = (data.get("address") or "").strip()
    if not name:
        return jsonify({"success": False, "error": "會員姓名必填"})
    if not phone:
        return jsonify({"success": False, "error": "電話必填（會員用此登入）"})

    conn = get_db()
    ag = conn.execute("SELECT prefix FROM agents WHERE id=?", (aid,)).fetchone()
    if not ag:
        conn.close()
        return jsonify({"success": False, "error": "代理資料異常"}), 500
    prefix = ag["prefix"]

    # 自動產生下一個 g_code（前綴+四位流水）
    g_code_in = (data.get("g_code") or "").strip().upper()
    if g_code_in:
        # 手動指定的：必須以該代理前綴開頭
        if not g_code_in.startswith(prefix):
            conn.close()
            return jsonify({"success": False, "error": f"會員編號必須以「{prefix}」開頭"})
        # 不可重複
        exists = conn.execute("SELECT 1 FROM members WHERE g_code=?", (g_code_in,)).fetchone()
        if exists:
            conn.close()
            return jsonify({"success": False, "error": f"編號「{g_code_in}」已使用"})
        g_code = g_code_in
    else:
        used = set()
        for r in conn.execute("SELECT g_code FROM members WHERE agent_id=?", (aid,)).fetchall():
            gc = r["g_code"] or ""
            if gc.startswith(prefix):
                try:
                    used.add(int(gc[len(prefix):]))
                except (ValueError, TypeError):
                    pass
        n = 1
        while n in used:
            n += 1
        g_code = f"{prefix}{n:04d}"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 處理會員專屬費率（>= 代理 min_rate；留空/0 = 沿用 min_rate）
    raw_rate = data.get("shipping_rate")
    rate_val = 0.0
    if raw_rate not in (None, "", 0, "0"):
        try:
            rate_val = float(raw_rate)
        except (ValueError, TypeError):
            conn.close()
            return jsonify({"success": False, "error": "運費必須為數字"})
        # 取該代理 min_rate
        ag_row = conn.execute("SELECT min_rate FROM agents WHERE id=?", (aid,)).fetchone()
        min_rate = float(ag_row["min_rate"] or 180) if ag_row else 180.0
        if rate_val < min_rate:
            conn.close()
            return jsonify({"success": False, "error": f"運費 NT${rate_val}/kg 低於您的最低費率 NT${int(min_rate)}/kg"})
    try:
        conn.execute(
            """INSERT INTO members (g_code, agent_id, name, phone, address, line_id, email, note, status, created_at, shipping_rate)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'active', ?, ?)""",
            (g_code, aid, name, phone, address,
             (data.get("line_id") or "").strip(), (data.get("email") or "").strip(),
             (data.get("note") or "").strip(), now, rate_val)
        )
        conn.commit()
    except sqlite3.IntegrityError as e:
        conn.close()
        return jsonify({"success": False, "error": f"資料庫錯誤：{e}"})
    conn.close()
    return jsonify({"success": True, "g_code": g_code, "message": f"已建立「{g_code} {name}」"})


@app.route("/api/admin/members/<g_code>", methods=["PUT"])
def admin_update_member(g_code):
    """代理更新自己會員的資料（姓名、電話、地址等）"""
    g_code = g_code.upper()
    aid = get_current_agent_id()
    conn = get_db()
    row = conn.execute("SELECT * FROM members WHERE g_code=?", (g_code,)).fetchone()
    if not row:
        conn.close()
        return jsonify({"success": False, "error": "找不到會員"})
    if aid > 0 and int(row["agent_id"] or 0) != aid:
        conn.close()
        return jsonify({"success": False, "error": "權限不足"}), 403

    data = request.json or {}
    fields, values = [], []
    for col in ["name", "phone", "address", "line_id", "email", "note", "status"]:
        if col in data:
            v = (data[col] or "").strip()
            if col == "phone":
                v = normalize_phone(v)
            if col == "status" and v not in ("active", "disabled"):
                continue
            fields.append(f"{col}=?"); values.append(v)
    # 會員專屬費率
    if "shipping_rate" in data:
        raw = data["shipping_rate"]
        if raw in (None, "", 0, "0"):
            rate_val = 0.0
        else:
            try:
                rate_val = float(raw)
            except (ValueError, TypeError):
                conn.close()
                return jsonify({"success": False, "error": "運費必須為數字"})
            if aid > 0:
                ag_row = conn.execute("SELECT min_rate FROM agents WHERE id=?", (aid,)).fetchone()
                min_rate = float(ag_row["min_rate"] or 180) if ag_row else 180.0
                if rate_val < min_rate:
                    conn.close()
                    return jsonify({"success": False, "error": f"運費 NT${rate_val}/kg 低於您的最低費率 NT${int(min_rate)}/kg"})
        fields.append("shipping_rate=?"); values.append(rate_val)
    if not fields:
        conn.close()
        return jsonify({"success": False, "error": "沒有可更新欄位"})
    values.append(g_code)
    conn.execute(f"UPDATE members SET {', '.join(fields)} WHERE g_code=?", values)
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/admin/members/<g_code>", methods=["DELETE"])
def admin_delete_member(g_code):
    """代理刪除自己會員（若會員底下有任何包裹/預報/出貨紀錄，擋下，建議停用）"""
    g_code = g_code.upper()
    aid = get_current_agent_id()
    conn = get_db()
    row = conn.execute("SELECT * FROM members WHERE g_code=?", (g_code,)).fetchone()
    if not row:
        conn.close()
        return jsonify({"success": False, "error": "找不到會員"})
    if aid > 0 and int(row["agent_id"] or 0) != aid:
        conn.close()
        return jsonify({"success": False, "error": "權限不足"}), 403
    # 安全檢查：是否有關聯資料
    p = conn.execute("SELECT COUNT(*) as c FROM packages WHERE g_code=?", (g_code,)).fetchone()["c"]
    f = conn.execute("SELECT COUNT(*) as c FROM forecasts WHERE g_code=?", (g_code,)).fetchone()["c"]
    s = conn.execute("SELECT COUNT(*) as c FROM shipment_requests WHERE g_code=?", (g_code,)).fetchone()["c"]
    if p + f + s > 0:
        conn.close()
        return jsonify({
            "success": False,
            "error": f"此會員已有 {p} 個包裹/{f} 個預報/{s} 個出貨紀錄，無法刪除。請改為「停用」狀態。"
        })
    conn.execute("DELETE FROM members WHERE g_code=?", (g_code,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "會員已刪除"})


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
    aid = get_current_agent_id()
    conn = get_db()
    if g_code:
        if aid > 0:
            rows = conn.execute(
                "SELECT * FROM packages WHERE g_code=? AND agent_id=? ORDER BY id DESC",
                (g_code.upper(), aid)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM packages WHERE g_code=? ORDER BY id DESC", (g_code.upper(),)
            ).fetchall()
    else:
        if aid > 0:
            rows = conn.execute(
                "SELECT * FROM packages WHERE agent_id=? ORDER BY id DESC", (aid,)
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM packages ORDER BY id DESC").fetchall()
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
    # 開頭非字母 → 補上對應前綴（代理用自己的前綴、主管理員預設 G）
    if not g_code[:1].isalpha():
        aid = get_current_agent_id()
        if aid > 0:
            conn0 = get_db()
            ag = conn0.execute("SELECT prefix FROM agents WHERE id=?", (aid,)).fetchone()
            conn0.close()
            g_code = (ag["prefix"] if ag else "G") + g_code
        else:
            g_code = "G" + g_code

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today = datetime.now().strftime("%Y-%m-%d")
    pkg_agent_id = get_agent_id_for_g_code(g_code)

    # 代理只能幫自己的客戶建包裹
    aid = get_current_agent_id()
    if aid > 0 and pkg_agent_id != aid:
        return jsonify({"success": False, "error": f"客戶編號「{g_code}」不屬於你的代理帳號"}), 403

    conn = get_db()
    cur = conn.execute(
        """INSERT INTO packages (g_code, logis_num, product_name, weight, status, note, in_date, created_at, agent_id)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (g_code, logis_num, product_name, weight, status, note, today, now, pkg_agent_id)
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return jsonify({"success": True, "id": new_id})


@app.route("/api/admin/packages/<int:pkg_id>", methods=["PUT"])
def admin_update_package(pkg_id):
    ok, row = check_record_ownership("packages", pkg_id)
    if not row:
        return jsonify({"success": False, "error": "找不到包裹"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json
    fields = []
    values = []
    for key in ["g_code", "logis_num", "product_name", "weight", "status", "note", "in_date"]:
        if key in data:
            val = data[key]
            if key == "g_code":
                val = val.strip().upper()
                # 改 g_code 時驗證新 g_code 仍屬於同代理
                aid_chk = get_current_agent_id()
                if aid_chk > 0 and get_agent_id_for_g_code(val) != aid_chk:
                    return jsonify({"success": False, "error": f"無法將包裹改至非自己客戶「{val}」"}), 403
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
    aid = get_current_agent_id()
    conn = get_db()
    if aid > 0:
        # 驗證所有 id 都屬於該代理
        placeholders = ",".join(["?"] * len(ids))
        owned = conn.execute(
            f"SELECT id FROM packages WHERE id IN ({placeholders}) AND agent_id=?",
            ids + [aid]
        ).fetchall()
        owned_ids = [r["id"] for r in owned]
        if len(owned_ids) != len(ids):
            conn.close()
            return jsonify({"success": False, "error": "部分包裹不屬於你的代理帳號"}), 403
        ids = owned_ids
    conn.execute(
        f"UPDATE packages SET status='已出貨' WHERE id IN ({','.join(['?']*len(ids))})",
        ids
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "updated": len(ids)})


@app.route("/api/admin/packages/<int:pkg_id>", methods=["DELETE"])
def admin_delete_package(pkg_id):
    ok, row = check_record_ownership("packages", pkg_id)
    if not row:
        return jsonify({"success": False, "error": "找不到包裹"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
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
    # 沒有英文前綴 → 預設加 G（你的客戶）
    if not g_code[:1].isalpha():
        g_code = "G" + g_code
    password_clean = normalize_phone(password)

    # ===== 1) 先查本地 members 表（代理建的會員）=====
    try:
        conn = get_db()
        row = conn.execute("SELECT * FROM members WHERE g_code=?", (g_code,)).fetchone()
        if row:
            m = dict(row)
            # 狀態檢查
            if m.get("status") == "disabled":
                conn.close()
                return jsonify({"success": False, "error": "此會員帳號已停用，請聯絡您的代理"})
            # 比對電話（去除空白、橫線、+886/+81 等）
            stored_phone = normalize_phone(m.get("phone") or "")
            if stored_phone != password_clean:
                conn.close()
                return jsonify({"success": False, "error": "密碼錯誤，請輸入您的手機號碼"})
            # 找該代理的最低費率
            ag = conn.execute("SELECT min_rate, name FROM agents WHERE id=?", (m["agent_id"],)).fetchone()
            conn.close()
            agent_min = float(ag["min_rate"]) if ag and ag["min_rate"] else float(DEFAULT_SHIPPING_RATE)
            # 會員專屬費率 > 0 → 用該費率；否則用代理 min_rate
            member_rate = float(m.get("shipping_rate") or 0)
            rate_twd = int(member_rate if member_rate > 0 else agent_min)
            return jsonify({
                "success": True,
                "customer": {
                    "id": g_code,  # 本地會員無 Shopify customer_id，用 g_code
                    "g_code": g_code,
                    "name": m.get("name") or "會員",
                    "email": m.get("email") or "",
                    "phone": stored_phone,
                    "phone_raw": m.get("phone") or "",
                    "address": m.get("address") or "",
                    "shipping_rate_twd": rate_twd,
                    "shipping_rate_jpy": twd_to_jpy(rate_twd) if rate_twd else 0,
                    "source": "agent",
                    "agent_name": ag["name"] if ag else "",
                }
            })
        conn.close()
    except Exception as e:
        print(f"[verify_customer] 本地查詢失敗：{e}", flush=True)

    # ===== 2) 回退查 Shopify（你的客戶，原有邏輯）=====
    try:
        customers = get_all_goyoutati_customers()
        for c in customers:
            if c["g_code"] == g_code:
                if c["phone"] and c["phone"] == password_clean:
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
                            "shipping_rate_jpy": rate_jpy,
                            "source": "shopify",
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
    aid = get_current_agent_id()
    try:
        conn = get_db()
        if aid > 0:
            rows = conn.execute("""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0 AND agent_id=?
                ORDER BY updated_at ASC
            """, (aid,)).fetchall()
        else:
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
                "consolidation_fee": float(rd.get("consolidation_fee") or 0),
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
    aid = get_current_agent_id()
    try:
        conn = get_db()
        if aid > 0:
            rows = conn.execute("""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0 AND agent_id=?
                ORDER BY updated_at ASC
            """, (aid,)).fetchall()
        else:
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
    aid = get_current_agent_id()
    try:
        conn = get_db()
        if aid > 0:
            rows = conn.execute("""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0 AND agent_id=?
                ORDER BY updated_at DESC
            """, (aid,)).fetchall()
        else:
            rows = conn.execute("""
                SELECT * FROM shipment_requests
                WHERE status='已出貨' AND total_fee > 0
                ORDER BY updated_at DESC
            """).fetchall()
        conn.close()

        monthly = {}
        for row in rows:
            r = dict(row)
            date_str = r.get("updated_at") or r.get("created_at") or ""
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
                    "consolidation_fee": 0,
                    "extra_fee": 0,
                    "total_revenue": 0,
                    "customers": set()
                }

            m = monthly[month_key]
            m["shipments"] += 1
            m["total_kg"] += float(r["billed_weight"] or 0)
            m["shipping_fee"] += float(r["shipping_fee"] or 0)
            m["handling_fee"] += float(r["handling_fee"] or 0)
            m["consolidation_fee"] += float(r.get("consolidation_fee") or 0)
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


# ============ JPD 自動建單 ============

@app.route("/api/admin/shipment_requests/<int:req_id>/jpd_create", methods=["POST"])
def admin_create_jpd_order(req_id):
    """從出貨申請自動在 JPD 建立運單"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到出貨申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    req = conn.execute("SELECT * FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
    req = dict(req)
    g_code = req.get("g_code", "")
    
    # 收件人資訊
    recipient = str(req.get("ship_recipient") or "")
    phone = str(req.get("ship_phone") or "")
    address = str(req.get("ship_address") or "")
    note = str(req.get("note") or "")
    
    if not recipient or not phone or not address:
        conn.close()
        return jsonify({"success": False, "error": "缺少寄送地址資訊"})
    
    # 從預報取得申報品項
    forecasts = conn.execute(
        "SELECT * FROM forecasts WHERE g_code=? AND status='待處理' ORDER BY id", (g_code,)
    ).fetchall()
    
    declare_list = []
    for fc in forecasts:
        fc = dict(fc)
        try:
            items = json.loads(fc.get("items_json") or "[]")
            for item in items:
                declare_list.append({
                    "product_name": item.get("name", "雑貨"),
                    "product_name_local": item.get("name", "雑貨"),
                    "product_num": int(item.get("quantity", 1)),
                    "product_price": int(float(item.get("price", 0))),
                    "product_url": item.get("url", "")
                })
        except Exception as e:
            print(f"[JPD] declare_list 解析失敗: {e}", flush=True)

    # 如果沒有預報品項，給一個預設品項
    if not declare_list:
        declare_list = [{
            "product_name": "雑貨",
            "product_name_local": "雑貨",
            "product_num": 1,
            "product_price": 0,
            "product_url": ""
        }]
    
    # 入庫包裹 ID 由管理員手動輸入（多個用逗號/頓號/空白分隔）
    body = request.get_json(silent=True) or {}
    raw_ids = body.get("package_ids", "")
    if isinstance(raw_ids, list):
        id_tokens = [str(x).strip() for x in raw_ids]
    else:
        id_tokens = re.split(r"[,，、\s]+", str(raw_ids))
    jpd_package_ids = []
    for tok in id_tokens:
        tok = tok.strip()
        if not tok:
            continue
        try:
            jpd_package_ids.append(int(tok))
        except (ValueError, TypeError):
            conn.close()
            return jsonify({"success": False, "error": f"包裹 ID「{tok}」格式錯誤，應為數字"})

    if not jpd_package_ids:
        conn.close()
        return jsonify({"success": False, "error": "請填入 JPD 入庫包裹 ID（可在 JPD 雲倉的入庫列表查到，多個用逗號分隔）"})

    # 客戶運單號前綴：客編 + 日期
    today_str = datetime.now().strftime("%m%d")
    order_prefix = f"{g_code}-{today_str}"

    # 方案 A：一個包裹建一張運單，運單號一律加流水號（-1, -2, -3...）
    created_orders = []
    failed_orders = []

    for idx, pid in enumerate(jpd_package_ids):
        customer_order_id = f"{order_prefix}-{idx + 1}"
        order_data = {
            "customer_order_id": customer_order_id,
            "deliv_id": JPD_DELIV_ID,
            "recipient": recipient,
            "id_issure": "",
            "area": 3,
            "addr1": address,
            "addr2": "",
            "addr3": "",
            "addr4": "",
            "tel": phone,
            "memo": note,
            "create_order_pdf": "n",
            "warehouse_id": JPD_WAREHOUSE_ID,
            "create_package": "n",
            "create_sender": "y",
            "packages": [{"package_id": pid, "declare_list": declare_list}]
        }
        print(f"[JPD] 建立運單: {customer_order_id}, 收件人: {recipient}, 包裹 id={pid}", flush=True)
        result = jpd_request("TCreateOrder", order_data)

        ok = False
        if "OperationResult" in result:
            op = result["OperationResult"]
            if op.get("Request", {}).get("IsValid") == "True":
                res_data = op.get("Result", {})
                if res_data.get("Result") == "SUCCESS":
                    data = res_data.get("Data", {})
                    jpd_order_id = str(data.get("order_id", ""))
                    jpd_logis_num = str(data.get("logis_num", ""))
                    print(f"[JPD] ✅ 建單成功: {customer_order_id} order_id={jpd_order_id}, logis_num={jpd_logis_num}", flush=True)
                    created_orders.append({
                        "customer_order_id": customer_order_id,
                        "jpd_order_id": jpd_order_id,
                        "jpd_logis_num": jpd_logis_num,
                        "package_id": pid
                    })
                    ok = True
                else:
                    err = res_data.get("ErrorMsg", str(res_data))
                    print(f"[JPD] ❌ 建單失敗 {customer_order_id}: {err}", flush=True)
                    failed_orders.append({"customer_order_id": customer_order_id, "error": str(err)})
            else:
                errs = op.get("Request", {}).get("Errors", {})
                print(f"[JPD] ❌ 請求無效 {customer_order_id}: {errs}", flush=True)
                failed_orders.append({"customer_order_id": customer_order_id, "error": str(errs)})
        if not ok and not failed_orders:
            failed_orders.append({"customer_order_id": customer_order_id, "error": "JPD API 無回應"})

    conn.close()

    if created_orders and not failed_orders:
        nums = ', '.join([o["jpd_logis_num"] or o["customer_order_id"] for o in created_orders])
        return jsonify({
            "success": True,
            "created": created_orders,
            "message": f"已建立 {len(created_orders)} 張 JPD 運單：{nums}"
        })
    elif created_orders and failed_orders:
        return jsonify({
            "success": True,
            "created": created_orders,
            "failed": failed_orders,
            "message": f"成功 {len(created_orders)} 張、失敗 {len(failed_orders)} 張。失敗：" +
                       '、'.join([f["customer_order_id"] + '(' + f["error"] + ')' for f in failed_orders])
        })
    else:
        first_err = failed_orders[0]["error"] if failed_orders else "未知錯誤"
        return jsonify({"success": False, "error": f"JPD 建單全部失敗：{first_err}"})


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
    """管理員新增公告（僅主管理員）"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
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
    """管理員更新公告（僅主管理員）"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
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
    """管理員刪除公告（僅主管理員）"""
    if not is_super_admin():
        return jsonify({"success": False, "error": "權限不足"}), 403
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
    sr_agent_id = get_agent_id_for_g_code(g_code)

    conn.execute(
        """INSERT INTO shipment_requests (g_code, customer_name, package_ids, package_summary, status, note, ship_recipient, ship_phone, ship_address, created_at, agent_id)
           VALUES (?, ?, ?, ?, '待處理', ?, ?, ?, ?, ?, ?)""",
        (g_code, customer_name, ids_str, summary, note, ship_recipient, ship_phone, ship_address, now, sr_agent_id)
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


@app.route("/api/admin/old_packages", methods=["GET"])
def admin_old_packages():
    """查詢倉庫滯留超過 N 天的未出貨包裹（預設 30 天）"""
    try:
        days = int(request.args.get("days", 30))
    except (ValueError, TypeError):
        days = 30
    cutoff_date = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    aid = get_current_agent_id()
    conn = get_db()
    if aid > 0:
        rows = conn.execute(
            """SELECT * FROM packages
               WHERE status != '已出貨'
                 AND agent_id = ?
                 AND COALESCE(NULLIF(in_date, ''), substr(created_at, 1, 10)) <= ?
               ORDER BY COALESCE(NULLIF(in_date, ''), substr(created_at, 1, 10)) ASC""",
            (aid, cutoff_date)
        ).fetchall()
    else:
        rows = conn.execute(
            """SELECT * FROM packages
               WHERE status != '已出貨'
                 AND COALESCE(NULLIF(in_date, ''), substr(created_at, 1, 10)) <= ?
               ORDER BY COALESCE(NULLIF(in_date, ''), substr(created_at, 1, 10)) ASC""",
            (cutoff_date,)
        ).fetchall()
    conn.close()
    today = datetime.now().date()
    result = []
    for r in rows:
        d = dict(r)
        ref_date_str = d.get("in_date") or (d.get("created_at") or "")[:10]
        try:
            ref_date = datetime.strptime(ref_date_str, "%Y-%m-%d").date()
            age_days = (today - ref_date).days
        except (ValueError, TypeError):
            age_days = 0
        d["age_days"] = age_days
        d["ref_date"] = ref_date_str
        result.append(d)
    return jsonify({
        "success": True,
        "count": len(result),
        "days": days,
        "packages": result
    })


@app.route("/api/admin/customer_unpaid/<g_code>", methods=["GET"])
def admin_customer_unpaid(g_code):
    """查詢客戶未付款的已出貨筆數和金額（用於出貨警告）"""
    # 代理只能查自己的客戶
    aid = get_current_agent_id()
    if aid > 0 and get_agent_id_for_g_code(g_code.upper()) != aid:
        return jsonify({"success": True, "count": 0, "total": 0, "latest": "", "ids": []})
    conn = get_db()
    rows = conn.execute(
        "SELECT id, total_fee, updated_at FROM shipment_requests "
        "WHERE g_code=? AND status='已出貨' AND total_fee > 0 "
        "AND (payment_last5 IS NULL OR payment_last5='') "
        "ORDER BY updated_at DESC",
        (g_code.upper(),)
    ).fetchall()
    conn.close()
    items = [dict(r) for r in rows]
    return jsonify({
        "success": True,
        "count": len(items),
        "total": sum(int(r.get("total_fee") or 0) for r in items),
        "latest": items[0]["updated_at"] if items else "",
        "ids": [r["id"] for r in items]
    })


@app.route("/api/admin/shipment_requests/<int:req_id>/confirm_payment", methods=["POST"])
def admin_confirm_payment(req_id):
    """管理員確認匯款已收到（可填後五碼、LINE Pay、現金等任意備註）"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    data = request.json or {}
    note = (data.get("last5") or "").strip()

    if len(note) > 20:
        return jsonify({"success": False, "error": "備註請勿超過 20 字"})
    if not note:
        note = "管確認"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    conn.execute(
        "UPDATE shipment_requests SET payment_last5=?, payment_at=? WHERE id=?",
        (note, now, req_id)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "已確認匯款"})


@app.route("/api/admin/shipment_requests/<int:req_id>/unconfirm_payment", methods=["POST"])
def admin_unconfirm_payment(req_id):
    """管理員取消已確認的匯款（誤按時用）"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    conn.execute(
        "UPDATE shipment_requests SET payment_last5='', payment_at='' WHERE id=?",
        (req_id,)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "已取消匯款確認"})


@app.route("/api/admin/shipment_requests", methods=["GET"])
def admin_get_shipment_requests():
    """管理員查看所有出貨申請（含對應客戶的待處理預報資料）"""
    status = request.args.get("status", "")
    aid = get_current_agent_id()
    # 代理過濾片段
    af = " AND agent_id=?" if aid > 0 else ""
    aparams = (aid,) if aid > 0 else ()
    try:
        conn = get_db()
        if status == "已付款":
            sql = "SELECT * FROM shipment_requests WHERE status='已出貨' AND payment_last5 != '' AND payment_last5 IS NOT NULL" + af + " ORDER BY payment_at DESC, id DESC"
            rows = conn.execute(sql, aparams).fetchall()
        elif status == "recent":
            if aid > 0:
                rows = conn.execute(
                    "SELECT * FROM shipment_requests WHERE agent_id=? ORDER BY id DESC LIMIT 50", (aid,)
                ).fetchall()
            else:
                rows = conn.execute("SELECT * FROM shipment_requests ORDER BY id DESC LIMIT 50").fetchall()
        elif status:
            sql = "SELECT * FROM shipment_requests WHERE status=?" + af + " ORDER BY id DESC"
            rows = conn.execute(sql, (status,) + aparams).fetchall()
        else:
            if aid > 0:
                rows = conn.execute(
                    "SELECT * FROM shipment_requests WHERE agent_id=? ORDER BY id DESC LIMIT 50", (aid,)
                ).fetchall()
            else:
                rows = conn.execute("SELECT * FROM shipment_requests ORDER BY id DESC LIMIT 50").fetchall()

        # 一次撈出涉及到的客戶的待處理預報（避免 N+1 查詢）
        g_codes = list({r["g_code"] for r in rows if r["g_code"]})
        forecast_map = {}
        if g_codes:
            placeholders = ",".join(["?"] * len(g_codes))
            fc_rows = conn.execute(
                f"SELECT * FROM forecasts WHERE g_code IN ({placeholders}) AND status='待處理' ORDER BY id",
                g_codes
            ).fetchall()
            for fc in fc_rows:
                try:
                    items = json.loads(fc["items_json"] or "[]")
                except (TypeError, ValueError, json.JSONDecodeError):
                    items = []
                forecast_map.setdefault(fc["g_code"], []).append({
                    "id": fc["id"],
                    "note": fc["note"] or "",
                    "created_at": fc["created_at"] or "",
                    "items": items
                })

        conn.close()

        result = []
        for r in rows:
            d = dict(r)
            d["pending_forecasts"] = forecast_map.get(r["g_code"], [])
            result.append(d)
        return jsonify({"success": True, "requests": result})
    except Exception as e:
        return jsonify({"success": False, "error": str(e), "requests": []})


@app.route("/api/admin/shipment_requests/<int:req_id>", methods=["PUT"])
def admin_update_shipment_request(req_id):
    """管理員更新出貨申請狀態（含帳單資訊）"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
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
    consolidation_fee = data.get("consolidation_fee", 0)
    total_fee = data.get("total_fee", 0)
    tracking_num = data.get("tracking_num", "")
    extra_services = json.dumps(data.get("extra_services", []), ensure_ascii=False)

    # 代理出貨時：強制 rate_per_kg 不得低於自己的最低費率
    aid = get_current_agent_id()
    if aid > 0 and status == "已出貨" and rate_per_kg:
        try:
            rate_val = float(rate_per_kg)
            ag = conn.execute("SELECT min_rate FROM agents WHERE id=?", (aid,)).fetchone()
            min_rate = float(ag["min_rate"]) if ag and ag["min_rate"] else 180.0
            if rate_val < min_rate:
                conn.close()
                return jsonify({
                    "success": False,
                    "error": f"運費 NT${rate_val}/kg 低於你的最低費率 NT${min_rate}/kg"
                })
        except (ValueError, TypeError):
            pass

    if status == "已出貨" and billed_weight:
        conn.execute(
            """UPDATE shipment_requests 
               SET status=?, admin_note=?, updated_at=?,
                   billed_weight=?, rate_per_kg=?, shipping_fee=?, handling_fee=?, consolidation_fee=?, total_fee=?,
                   tracking_num=?, extra_services=?
               WHERE id=?""",
            (status, admin_note, now, billed_weight, rate_per_kg, shipping_fee, handling_fee, consolidation_fee, total_fee, tracking_num, extra_services, req_id)
        )
    else:
        conn.execute(
            "UPDATE shipment_requests SET status=?, admin_note=?, updated_at=? WHERE id=?",
            (status, admin_note, now, req_id)
        )

    # 如果管理員標記為「已出貨」，同步更新包裹狀態 + 預報標為已處理
    g_code_val = ""
    customer_name_val = ""
    if status == "已出貨":
        req = conn.execute("SELECT package_ids, g_code, customer_name FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
        if req:
            g_code_val = req["g_code"]
            customer_name_val = req["customer_name"] or ""
            pkg_ids = [int(x.strip()) for x in req["package_ids"].split(",") if x.strip()]
            if pkg_ids:
                placeholders = ",".join(["?"] * len(pkg_ids))
                conn.execute(
                    f"UPDATE packages SET status='已出貨' WHERE id IN ({placeholders})", pkg_ids
                )
            # 自動把該客戶的待處理預報標為已處理
            conn.execute(
                "UPDATE forecasts SET status='已處理' WHERE g_code=? AND status='待處理'",
                (g_code_val,)
            )
    else:
        req = conn.execute("SELECT g_code, customer_name FROM shipment_requests WHERE id=?", (req_id,)).fetchone()
        if req:
            g_code_val = req["g_code"]
            customer_name_val = req["customer_name"] or ""

    conn.commit()
    conn.close()
    return jsonify({"success": True, "g_code": g_code_val, "customer_name": customer_name_val})


@app.route("/api/admin/shipment_requests/<int:req_id>/revert", methods=["POST"])
def admin_revert_shipment_request(req_id):
    """還原出貨申請：狀態回到待處理，包裹回到已到貨，清空帳單"""
    ok, _row = check_record_ownership("shipment_requests", req_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該申請"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
    conn = get_db()
    req = conn.execute("SELECT * FROM shipment_requests WHERE id=?", (req_id,)).fetchone()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        """UPDATE shipment_requests 
           SET status='待處理', updated_at=?,
               billed_weight=0, rate_per_kg=0, shipping_fee=0, handling_fee=0,
               consolidation_fee=0, total_fee=0,
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
    fc_agent_id = get_agent_id_for_g_code(g_code)
    conn = get_db()
    conn.execute(
        """INSERT INTO forecasts (g_code, customer_name, items_json, status, note, created_at, agent_id)
           VALUES (?, ?, ?, '待處理', ?, ?, ?)""",
        (g_code, customer_name, json.dumps(valid_items, ensure_ascii=False), note, now, fc_agent_id)
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
    g_code = request.args.get("g_code", "").upper()
    aid = get_current_agent_id()
    af = " AND agent_id=?" if aid > 0 else ""
    aparams = (aid,) if aid > 0 else ()
    conn = get_db()
    if g_code and status:
        rows = conn.execute(
            "SELECT * FROM forecasts WHERE g_code=? AND status=?" + af + " ORDER BY id DESC",
            (g_code, status) + aparams
        ).fetchall()
    elif g_code:
        rows = conn.execute(
            "SELECT * FROM forecasts WHERE g_code=?" + af + " ORDER BY id DESC", (g_code,) + aparams
        ).fetchall()
    elif status:
        rows = conn.execute(
            "SELECT * FROM forecasts WHERE status=?" + af + " ORDER BY id DESC", (status,) + aparams
        ).fetchall()
    else:
        if aid > 0:
            rows = conn.execute("SELECT * FROM forecasts WHERE agent_id=? ORDER BY id DESC", (aid,)).fetchall()
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
    ok, _row = check_record_ownership("forecasts", fc_id)
    if not _row:
        return jsonify({"success": False, "error": "找不到該預報"})
    if not ok:
        return jsonify({"success": False, "error": "權限不足"}), 403
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
    ok, _row = check_record_ownership("forecasts", fc_id)
    if not _row:
        return "Not found", 404
    if not ok:
        return "Forbidden", 403
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
