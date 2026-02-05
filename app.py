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


def jpd_request(operation: str, data: dict) -> dict:
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


def shopify_graphql(query: str, variables: dict = None) -> dict:
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


def shopify_request(endpoint: str, method: str = "GET", data: dict = None) -> dict:
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


@app.route("/admin")
def admin_page():
    """Admin 頁面"""
    return render_template("admin.html")


@app.route("/api/admin/members", methods=["GET"])
def get_all_members():
    """取得所有已分配 G 編號的會員（含運費設定）"""
    
    # 查詢所有有 goyoutati_id 的客戶
    graphql_query = """
    {
        customers(first: 100, query: "metafield_namespace:custom metafield_key:goyoutati_id") {
            edges {
                node {
                    id
                    firstName
                    lastName
                    email
                    phone
                    createdAt
                    defaultAddress {
                        phone
                    }
                    gCode: metafield(namespace: "custom", key: "goyoutati_id") {
                        value
                    }
                    shippingRate: metafield(namespace: "custom", key: "shipping_rate") {
                        value
                    }
                }
            }
        }
    }
    """
    
    try:
        result = shopify_graphql(graphql_query)
        
        members = []
        max_number = 0
        
        if "data" in result:
            customers = result["data"].get("customers", {}).get("edges", [])
            
            for edge in customers:
                node = edge["node"]
                g_code_mf = node.get("gCode")
                g_code = g_code_mf["value"] if g_code_mf else ""
                
                if not g_code:
                    continue
                
                # 提取編號數字
                if g_code.startswith("G"):
                    try:
                        num = int(g_code[1:])
                        if num > max_number:
                            max_number = num
                    except:
                        pass
                
                # 取得運費
                rate_mf = node.get("shippingRate")
                shipping_rate = rate_mf["value"] if rate_mf else ""
                
                gid = node.get("id", "")
                customer_id = gid.split("/")[-1] if "/" in gid else gid
                
                customer_name = f"{node.get('lastName', '')}{node.get('firstName', '')}".strip()
                if not customer_name:
                    customer_name = node.get("email", "")
                
                default_address = node.get("defaultAddress") or {}
                phone = default_address.get("phone") or node.get("phone") or ""
                
                members.append({
                    "g_code": g_code,
                    "customer_id": customer_id,
                    "gid": gid,
                    "name": customer_name,
                    "email": node.get("email", ""),
                    "phone": phone,
                    "shipping_rate": shipping_rate,
                    "created_at": node.get("createdAt", "")
                })
        
        # 按 G 編號排序
        members.sort(key=lambda x: x["g_code"])
        
        # 計算下一個可用編號
        next_number = max_number + 1
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
    customer_gid = data.get("customer_gid", "")  # e.g. gid://shopify/Customer/12345
    shipping_rate = data.get("shipping_rate", "")
    
    if not customer_gid:
        return jsonify({"success": False, "error": "缺少客戶 ID"})
    
    if shipping_rate == "" or shipping_rate is None:
        return jsonify({"success": False, "error": "請輸入運費"})
    
    # 驗證是數字
    try:
        rate_val = int(shipping_rate)
        if rate_val < 0:
            return jsonify({"success": False, "error": "運費不能為負數"})
    except ValueError:
        return jsonify({"success": False, "error": "運費必須為整數"})
    
    # 使用 metafieldsSet mutation
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
                "type": "number_integer",
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
                return jsonify({
                    "success": True,
                    "shipping_rate": rate_val
                })
        
        # 檢查是否有 errors
        if "errors" in result:
            error_msg = str(result["errors"])
            return jsonify({"success": False, "error": error_msg})
        
        return jsonify({"success": False, "error": "設定失敗，請重試"})
        
    except Exception as e:
        print(f"❌ 錯誤: {e}")
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/admin/verify", methods=["POST"])
def admin_verify():
    """Admin 密碼驗證"""
    data = request.json
    password = data.get("password", "")
    
    # 從環境變數取得 admin 密碼，預設為 "admin123"
    admin_password = os.environ.get("ADMIN_PASSWORD", "admin123")
    
    if password == admin_password:
        return jsonify({"success": True})
    
    return jsonify({"success": False, "error": "密碼錯誤"})


@app.route("/")
def index():
    """首頁 - 客人預報表單"""
    return render_template("index.html")


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
    
    # 確保格式正確（G 開頭）
    if not g_code.startswith("G"):
        g_code = "G" + g_code
    
    # 清理密碼格式（移除空格、橫線等）
    password_clean = password.replace(" ", "").replace("-", "").replace("+886", "0")
    
    print(f"\n{'='*50}")
    print(f"🔍 查詢會員編號: {g_code}")
    
    # 用新的 customers query 搜尋
    graphql_query = """
    {
        customers(first: 100, query: "metafield_namespace:custom metafield_key:goyoutati_id") {
            edges {
                node {
                    id
                    firstName
                    lastName
                    email
                    phone
                    defaultAddress {
                        phone
                    }
                    gCode: metafield(namespace: "custom", key: "goyoutati_id") {
                        value
                    }
                    shippingRate: metafield(namespace: "custom", key: "shipping_rate") {
                        value
                    }
                }
            }
        }
    }
    """
    
    try:
        result = shopify_graphql(graphql_query)
        print(f"📥 GraphQL 回應: {json.dumps(result, ensure_ascii=False)[:1500]}")
        
        if "data" in result:
            customers = result["data"].get("customers", {}).get("edges", [])
            
            for edge in customers:
                node = edge["node"]
                g_code_mf = node.get("gCode")
                node_g_code = g_code_mf["value"] if g_code_mf else ""
                
                if node_g_code == g_code:
                    # 找到匹配的會員
                    default_address = node.get("defaultAddress") or {}
                    customer_phone = default_address.get("phone") or node.get("phone") or ""
                    
                    # 清理手機號碼格式
                    phone_clean = customer_phone.replace(" ", "").replace("-", "").replace("+886", "0")
                    
                    print(f"📱 客戶手機: {phone_clean}, 輸入密碼: {password_clean}")
                    
                    # 驗證密碼（手機號碼）
                    if phone_clean and phone_clean == password_clean:
                        gid = node.get("id", "")
                        customer_id = gid.split("/")[-1] if "/" in gid else gid
                        
                        customer_name = f"{node.get('lastName', '')}{node.get('firstName', '')}".strip()
                        if not customer_name:
                            customer_name = node.get("email", "會員")
                        
                        # 取得運費
                        rate_mf = node.get("shippingRate")
                        shipping_rate = int(rate_mf["value"]) if rate_mf and rate_mf["value"] else DEFAULT_SHIPPING_RATE
                        
                        print(f"✅ 登入成功: {customer_name} (ID: {customer_id}, 運費: {shipping_rate} 日圓/kg)")
                        
                        return jsonify({
                            "success": True,
                            "customer": {
                                "id": customer_id,
                                "g_code": g_code,
                                "name": customer_name,
                                "email": node.get("email", ""),
                                "phone": customer_phone,
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
    
    customer_id = data.get("customer_id")  # Shopify Customer ID
    g_code = data.get("g_code", "")  # G 編號
    packages = data.get("packages", [])
    
    if not customer_id:
        return jsonify({"success": False, "error": "缺少客戶編號"})
    
    if not packages:
        return jsonify({"success": False, "error": "請至少填寫一個包裹"})
    
    results = []
    
    for idx, pkg in enumerate(packages):
        # 產生預報編號：G編號 + 日期 + 序號
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        local_logis_num = f"{g_code}-{timestamp}-{idx+1}"
        
        # 組裝申報列表
        declare_list = []
        for item in pkg.get("items", []):
            declare_list.append({
                "product_name": item.get("name", "商品"),
                "product_name_local": item.get("name", "商品"),
                "product_num": int(item.get("quantity", 1)),
                "product_price": int(float(item.get("price", 0))),
                "product_url": item.get("url", "")
            })
        
        # 計算總數量和總價
        total_num = sum(int(item.get("quantity", 1)) for item in pkg.get("items", []))
        total_price = sum(int(float(item.get("price", 0))) * int(item.get("quantity", 1)) for item in pkg.get("items", []))
        
        # 呼叫 JPD API 預報
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
    
    # 查詢該客戶的包裹（用 G 編號）
    result = jpd_request("TSearchPackages", {
        "client_cid": g_code
    })
    
    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            packages = op_result.get("Result", {}).get("Data", [])
            
            # 整理包裹資訊
            formatted_packages = []
            for pkg in packages:
                formatted_packages.append({
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
            
            return jsonify({
                "success": True,
                "packages": formatted_packages
            })
    
    return jsonify({"success": False, "error": "查詢失敗"})


@app.route("/api/orders", methods=["GET"])
def get_orders():
    """查詢客戶的運單列表"""
    g_code = request.args.get("g_code") or request.args.get("customer_id")
    
    if not g_code:
        return jsonify({"success": False, "error": "缺少會員編號"})
    
    # 查詢該客戶的運單
    result = jpd_request("TSearchOrders", {
        "client_cid": g_code
    })
    
    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            orders = op_result.get("Result", {}).get("Data", [])
            
            formatted_orders = []
            for order in orders:
                formatted_orders.append({
                    "order_id": order.get("order_id"),
                    "customer_order_id": order.get("customer_order_id"),
                    "logis_num": order.get("logis_num"),
                    "status": order.get("status_name"),
                    "recipient": order.get("recipient"),
                    "create_date": order.get("create_date"),
                    "weight": order.get("weight"),
                    "deliv_fee": order.get("deliv_fee")
                })
            
            return jsonify({
                "success": True,
                "orders": formatted_orders
            })
    
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
