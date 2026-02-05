"""
客人集運預報系統
御用達 × JPD 雲倉
"""

from flask import Flask, request, jsonify, render_template
from datetime import datetime
import requests
import json
import os

app = Flask(__name__)

# ============ 設定區（從環境變數讀取）============
# JPD 雲倉 API
JPD_BASE_URL = "https://biz.cloudwh.jp"
JPD_EMAIL = os.environ.get("JPD_EMAIL", "")
JPD_PASSWORD = os.environ.get("JPD_PASSWORD", "")
JPD_WAREHOUSE_ID = int(os.environ.get("JPD_WAREHOUSE_ID", "1"))

# Shopify API（用於驗證客戶）
SHOPIFY_STORE = os.environ.get("SHOPIFY_STORE", "")
SHOPIFY_ACCESS_TOKEN = os.environ.get("SHOPIFY_ACCESS_TOKEN", "")

# 預設運費（日圓/kg），客戶沒設定時使用
DEFAULT_SHIPPING_RATE = int(os.environ.get("DEFAULT_SHIPPING_RATE", "0"))
# ================================


def normalize_phone(phone_raw):
    """統一手機格式：移除空格橫線，+886 / +81 → 0 開頭"""
    phone = phone_raw.replace(" ", "").replace("-", "")
    if phone.startswith("+886"):
        phone = "0" + phone[4:]
    elif phone.startswith("+81"):
        phone = "0" + phone[3:]
    return phone


def jpd_request(operation, data):
    """JPD 雲倉 API 請求"""
    url = f"{JPD_BASE_URL}/api/json.php?Service=SDC&Operation={operation}"

    payload = {
        "login_email": JPD_EMAIL,
        "login_password": JPD_PASSWORD,
        "data": data
    }

    print(f"\n{'='*50}")
    print(f"📤 JPD API 請求: {operation}")
    print(f"Data: {json.dumps(data, ensure_ascii=False, indent=2)}")

    try:
        response = requests.post(url, json=payload, timeout=30)
        result = response.json()
        print(f"📥 回應: {json.dumps(result, ensure_ascii=False, indent=2)}")
        return result
    except Exception as e:
        print(f"❌ 錯誤: {e}")
        return {"error": str(e)}


def shopify_graphql(query, variables=None):
    """Shopify GraphQL API 請求"""
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
    """Shopify REST API 請求"""
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
    """
    用原始的 metafieldDefinitions 查詢取得所有有 goyoutati_id 的客戶。
    同時也取得每位客戶的 shipping_rate metafield。
    回傳: list of dict
    """
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

                # 取得運費
                rate_mf = owner.get("shippingRate")
                shipping_rate = rate_mf["value"] if rate_mf and rate_mf.get("value") else ""

                customers.append({
                    "g_code": g_code,
                    "customer_id": customer_id,
                    "gid": gid,
                    "name": customer_name,
                    "email": owner.get("email", ""),
                    "phone": phone,
                    "phone_raw": phone_raw,
                    "shipping_rate": shipping_rate,
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


@app.route("/api/admin/verify", methods=["POST"])
def admin_verify():
    """Admin 密碼驗證"""
    data = request.json
    password = data.get("password", "")
    admin_password = os.environ.get("ADMIN_PASSWORD", "admin123")

    if password == admin_password:
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "密碼錯誤"})


@app.route("/api/admin/members", methods=["GET"])
def get_all_members():
    """取得所有已分配 G 編號的會員（含運費設定）"""
    try:
        members = get_all_goyoutati_customers()

        # 按 G 編號排序
        members.sort(key=lambda x: x["g_code"])

        # 收集所有已用的編號
        used_numbers = set()
        for m in members:
            if m["g_code"].startswith("G"):
                try:
                    used_numbers.add(int(m["g_code"][1:]))
                except:
                    pass

        max_number = max(used_numbers) if used_numbers else 0

        # 找出最小可用編號（跳號優先填補）
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
            "default_shipping_rate": DEFAULT_SHIPPING_RATE
        })

    except Exception as e:
        print(f"❌ 錯誤: {e}")
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/admin/shipping_rate", methods=["POST"])
def set_shipping_rate():
    """設定客戶的每公斤運費（存入 Shopify Customer Metafield）"""
    data = request.json
    customer_gid = data.get("customer_gid", "")
    shipping_rate = data.get("shipping_rate", "")

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
            metafields {
                key
                value
            }
            userErrors {
                field
                message
            }
        }
    }
    """

    variables = {
        "metafields": [
            {
                "ownerId": customer_gid,
                "namespace": "custom",
                "key": "shipping_rate",
                "type": "single_line_text_field",
                "value": str(rate_val)
            }
        ]
    }

    try:
        result = shopify_graphql(mutation, variables)
        print(f"📥 設定運費回應: {json.dumps(result, ensure_ascii=False)[:1000]}")

        if "data" in result:
            mutation_result = result["data"].get("metafieldsSet", {})
            user_errors = mutation_result.get("userErrors", [])

            if user_errors:
                error_msg = "; ".join([e["message"] for e in user_errors])
                return jsonify({"success": False, "error": error_msg})

            metafields = mutation_result.get("metafields", [])
            if metafields:
                return jsonify({"success": True, "shipping_rate": rate_val})

        if "errors" in result:
            return jsonify({"success": False, "error": str(result["errors"])})

        return jsonify({"success": False, "error": "設定失敗，請重試"})

    except Exception as e:
        print(f"❌ 錯誤: {e}")
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/verify_customer", methods=["POST"])
def verify_customer():
    """驗證客戶 G 編號 + 手機密碼，回傳含運費資訊"""
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

    print(f"\n{'='*50}")
    print(f"🔍 查詢會員編號: {g_code}")

    try:
        customers = get_all_goyoutati_customers()

        for c in customers:
            if c["g_code"] == g_code:
                print(f"📱 客戶手機: {c['phone']}, 輸入密碼: {password_clean}")

                if c["phone"] and c["phone"] == password_clean:
                    try:
                        shipping_rate = int(c["shipping_rate"]) if c["shipping_rate"] else DEFAULT_SHIPPING_RATE
                    except (ValueError, TypeError):
                        shipping_rate = DEFAULT_SHIPPING_RATE

                    print(f"✅ 登入成功: {c['name']} (ID: {c['customer_id']}, 運費: {shipping_rate} 日圓/kg)")

                    return jsonify({
                        "success": True,
                        "customer": {
                            "id": c["customer_id"],
                            "g_code": g_code,
                            "name": c["name"] or "會員",
                            "email": c["email"],
                            "phone": c["phone"],
                            "shipping_rate": shipping_rate
                        }
                    })
                else:
                    print(f"❌ 密碼錯誤")
                    return jsonify({"success": False, "error": "密碼錯誤，請輸入您的手機號碼"})

        print(f"❌ 找不到會員編號: {g_code}")
        return jsonify({"success": False, "error": "找不到此會員編號，請確認後重試"})

    except Exception as e:
        print(f"❌ 錯誤: {e}")
        return jsonify({"success": False, "error": f"查詢失敗: {str(e)}"})


@app.route("/api/forecast", methods=["POST"])
def create_forecast():
    """建立預報包裹"""
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

        results.append({
            "success": False,
            "local_logis_num": local_logis_num,
            "error": "預報失敗"
        })

    return jsonify({
        "success": all(r["success"] for r in results),
        "results": results
    })


@app.route("/api/packages", methods=["GET"])
def get_packages():
    """查詢客戶的包裹列表"""
    g_code = request.args.get("g_code") or request.args.get("customer_id")

    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})

    result = jpd_request("TSearchPackages", {
        "client_cid": g_code,
        "warehouse_id": JPD_WAREHOUSE_ID
    })

    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            packages = op_result.get("Result", {}).get("Data", [])

            formatted = []
            for pkg in packages:
                formatted.append({
                    "package_id": pkg.get("package_id"),
                    "local_logis_num": pkg.get("local_logis_num"),
                    "client_pid": pkg.get("client_pid"),
                    "status": pkg.get("status_name", "未知"),
                    "status_id": pkg.get("status_id"),
                    "weight": pkg.get("weight", "0"),
                    "product_name": pkg.get("product_name"),
                    "product_num": pkg.get("product_num"),
                    "create_date": pkg.get("create_date"),
                    "in_date": pkg.get("in_date"),
                    "declare_list": pkg.get("declare_list", [])
                })

            return jsonify({"success": True, "packages": formatted})

    return jsonify({"success": False, "error": "查詢失敗"})


@app.route("/api/orders", methods=["GET"])
def get_orders():
    """查詢客戶的運單列表"""
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"

    print(f"""
    ╔═══════════════════════════════════════════════════════════╗
    ║       客人集運預報系統                                      ║
    ║       御用達 × JPD 雲倉                                     ║
    ╚═══════════════════════════════════════════════════════════╝

    🌐 服務啟動於 Port: {port}
    """)
    app.run(host="0.0.0.0", port=port, debug=debug)
