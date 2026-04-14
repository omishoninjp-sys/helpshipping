#!/usr/bin/env python3
"""
Shopify × JPD 雲倉 串接工具
御用達-光頭哥 專用
"""

from flask import Flask, render_template, request, jsonify
import requests
import json
import os
from datetime import datetime, timedelta

app = Flask(__name__)

# ============ 從環境變數讀取設定 ============
# Shopify 設定
SHOPIFY_STORE = os.environ.get("SHOPIFY_STORE", "")
SHOPIFY_ACCESS_TOKEN = os.environ.get("SHOPIFY_ACCESS_TOKEN", "")

# JPD 雲倉設定
JPD_EMAIL = os.environ.get("JPD_EMAIL", "")
JPD_PASSWORD = os.environ.get("JPD_PASSWORD", "")
JPD_BASE_URL = os.environ.get("JPD_BASE_URL", "https://biz.cloudwh.jp")
JPD_WAREHOUSE_ID = int(os.environ.get("JPD_WAREHOUSE_ID", "1"))   # 足立倉庫
JPD_DELIV_ID = int(os.environ.get("JPD_DELIV_ID", "40"))          # 台灣空運線
# =============================================


def shopify_request(endpoint: str, method: str = "GET", data: dict = None) -> dict:
    """Shopify API 請求"""
    url = f"https://{SHOPIFY_STORE}.myshopify.com/admin/api/2026-01/{endpoint}"
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    
    try:
        if method == "GET":
            response = requests.get(url, headers=headers, timeout=30, verify=True)
        elif method == "POST":
            response = requests.post(url, headers=headers, json=data, timeout=30, verify=True)
        elif method == "PUT":
            response = requests.put(url, headers=headers, json=data, timeout=30, verify=True)
        
        return response.json()
    except requests.exceptions.SSLError:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        try:
            if method == "GET":
                response = requests.get(url, headers=headers, timeout=30, verify=False)
            elif method == "POST":
                response = requests.post(url, headers=headers, json=data, timeout=30, verify=False)
            elif method == "PUT":
                response = requests.put(url, headers=headers, json=data, timeout=30, verify=False)
            return response.json()
        except Exception as e:
            return {"error": f"SSL 錯誤: {str(e)}"}
    except Exception as e:
        return {"error": str(e)}


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


@app.route("/")
def index():
    """首頁"""
    return render_template("index.html")


@app.route("/health")
def health():
    """健康檢查"""
    return jsonify({
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
        "shopify_store": SHOPIFY_STORE,
        "jpd_configured": bool(JPD_EMAIL)
    })


@app.route("/api/shopify/orders")
def get_shopify_orders():
    """取得 Shopify 訂單列表"""
    status = request.args.get("status", "unfulfilled")
    limit = request.args.get("limit", 50)
    
    result = shopify_request(f"orders.json?status=any&fulfillment_status={status}&limit={limit}")
    
    print(f"\n{'='*50}")
    print(f"📦 Shopify API 回應:")
    print(json.dumps(result, ensure_ascii=False, indent=2)[:1000])
    
    if "orders" in result:
        orders = []
        for order in result["orders"]:
            shipping = order.get("shipping_address", {}) or {}
            customer_info = order.get("customer", {}) or {}
            billing = order.get("billing_address", {}) or {}
            
            print(f"\n--- 訂單 {order.get('name', '?')} 姓名來源 ---")
            print(f"  shipping.name        = '{shipping.get('name', '')}'")
            print(f"  shipping.first_name  = '{shipping.get('first_name', '')}'")
            print(f"  shipping.last_name   = '{shipping.get('last_name', '')}'")
            print(f"  customer.first_name  = '{customer_info.get('first_name', '')}'")
            print(f"  customer.last_name   = '{customer_info.get('last_name', '')}'")
            print(f"  billing.name         = '{billing.get('name', '')}'")
            print(f"  order.contact_email  = '{order.get('contact_email', '')}'")
            
            invalid_names = {"本人", "本人本人", "本人 本人", "同上", "同收件人", "test", "測試", ".", "-", ""}
            
            def is_valid(name):
                return name and name.strip() not in invalid_names
            
            shipping_name = (shipping.get("name") or "").strip()
            
            s_last  = (shipping.get("last_name")  or "").strip()
            s_first = (shipping.get("first_name") or "").strip()
            shipping_combined = f"{s_last}{s_first}".strip()
            
            c_last  = (customer_info.get("last_name")  or "").strip()
            c_first = (customer_info.get("first_name") or "").strip()
            customer_combined = f"{c_last}{c_first}".strip()
            
            billing_name = (billing.get("name") or "").strip()
            
            if is_valid(shipping_combined):
                customer_name = shipping_combined
            elif is_valid(customer_combined):
                customer_name = customer_combined
            elif is_valid(shipping_name):
                customer_name = shipping_name
            elif is_valid(billing_name):
                customer_name = billing_name
            else:
                customer_name = shipping_name or shipping_combined or customer_combined or "N/A"
            
            print(f"  ➡️ 最終使用: '{customer_name}'")

            # ===== 過濾已取消/已退款品項（fulfillable_quantity = 0）=====
            active_items = [
                {
                    "title": item["title"],
                    "variant_title": item.get("variant_title", ""),
                    "quantity": item.get("fulfillable_quantity", item["quantity"]),
                    "price": item["price"],
                    "sku": item.get("sku", "")
                }
                for item in order["line_items"]
                if item.get("fulfillable_quantity", item["quantity"]) > 0
            ]

            # 所有品項都已取消 → 跳過此訂單
            if not active_items:
                print(f"  ⚠️ 訂單 {order.get('name')} 所有品項已取消，略過")
                continue
            # =============================================================

            orders.append({
                "id": order["id"],
                "order_number": order["order_number"],
                "name": order["name"],
                "created_at": order["created_at"],
                "total_price": order.get("current_total_price", order["total_price"]),
                "currency": order["currency"],
                "fulfillment_status": order["fulfillment_status"] or "unfulfilled",
                "customer_name": customer_name,
                "phone": shipping.get("phone", ""),
                "address": " ".join(filter(None, [
                    shipping.get("province", ""),
                    shipping.get("city", ""),
                    shipping.get("address1", ""),
                    shipping.get("address2", "")
                ])).strip(),
                "line_items": active_items
            })
        return jsonify({"success": True, "orders": orders})
    
    error_msg = result.get("error") or result.get("errors") or str(result)
    return jsonify({"success": False, "error": error_msg})


@app.route("/api/shopify/order/<order_id>")
def get_shopify_order(order_id):
    """取得單一 Shopify 訂單詳情"""
    result = shopify_request(f"orders/{order_id}.json")
    
    if "order" in result:
        return jsonify({"success": True, "order": result["order"]})
    
    return jsonify({"success": False, "error": result.get("error", "Order not found")})


@app.route("/api/jpd/packages")
def get_jpd_packages():
    """取得 JPD 倉庫的包裹列表，並合併運單收件人資訊"""
    result = jpd_request("TSearchPackages", {
        "stock_date_from": (datetime.now().replace(day=1)).strftime("%Y-%m-%d 00:00:00")
    })
    
    if "OperationResult" not in result:
        return jsonify({"success": False, "error": "Unknown error"})
    
    op_result = result["OperationResult"]
    if op_result["Request"]["IsValid"] != "True":
        errors = op_result["Request"].get("Errors", {})
        return jsonify({"success": False, "error": errors})
    
    packages = op_result["Result"].get("Data", [])
    
    # 收集有 order_id 的包裹，去撈運單的收件人資訊
    order_ids = set()
    for pkg in packages:
        oid = pkg.get("order_id")
        if oid and str(oid) != "0":
            order_ids.add(str(oid))
    
    order_map = {}  # order_id -> {recipient, tel, addr1, customer_order_id}
    if order_ids:
        # 用同月份範圍撈運單
        orders_result = jpd_request("TSearchOrders", {
            "create_date_from": (datetime.now().replace(day=1)).strftime("%Y-%m-%d"),
            "create_date_to": datetime.now().strftime("%Y-%m-%d")
        })
        if "OperationResult" in orders_result:
            orders_op = orders_result["OperationResult"]
            if orders_op["Request"]["IsValid"] == "True":
                orders_data = orders_op.get("Result", {}).get("Data", [])
                for o in orders_data:
                    oid = str(o.get("order_id", ""))
                    if oid:
                        order_map[oid] = {
                            "recipient": o.get("recipient", ""),
                            "tel": o.get("tel", ""),
                            "addr1": o.get("addr1", ""),
                            "customer_order_id": o.get("customer_order_id", ""),
                            "logis_num": o.get("logis_num", ""),
                        }
    
    # 合併到包裹資料
    for pkg in packages:
        oid = str(pkg.get("order_id", ""))
        info = order_map.get(oid, {})
        pkg["recipient"] = info.get("recipient", "")
        pkg["tel"] = info.get("tel", "")
        pkg["addr1"] = info.get("addr1", "")
        pkg["customer_order_id"] = info.get("customer_order_id", "")
        pkg["logis_num"] = info.get("logis_num", "")
    
    return jsonify({"success": True, "packages": packages})


@app.route("/api/jpd/orders")
def get_jpd_orders():
    """取得 JPD 運單列表（支援日期範圍）"""
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")
    
    search_params = {}
    if date_from:
        search_params["create_date_from"] = date_from
    if date_to:
        search_params["create_date_to"] = date_to
    if not date_from and not date_to:
        # 預設：近 30 天
        search_params["create_date_from"] = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
        search_params["create_date_to"] = datetime.now().strftime("%Y-%m-%d")
    
    result = jpd_request("TSearchOrders", search_params)
    
    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            orders = op_result["Result"].get("Data", [])
            return jsonify({"success": True, "orders": orders})
    
    return jsonify({"success": False, "error": "Failed to fetch JPD orders"})


@app.route("/api/jpd/create_order", methods=["POST"])
def create_jpd_order():
    """創建 JPD 運單"""
    data = request.json
    mode = data.get("mode", "self")
    
    declare_list = []
    for item in data.get("declare_list", []):
        declare_list.append({
            "product_name": item.get("product_name", "商品"),
            "product_name_local": item.get("product_name_local", item.get("product_name", "商品")),
            "product_num": int(item.get("product_num", 1)),
            "product_price": int(item.get("product_price", 100))
        })
    
    total_num = sum(int(item.get("product_num", 1)) for item in data.get("declare_list", []))
    total_price = sum(int(item.get("product_price", 0)) * int(item.get("product_num", 1)) for item in data.get("declare_list", []))
    
    package_ids = []
    
    if mode == "warehouse":
        if not data.get("package_ids"):
            return jsonify({"success": False, "error": "倉庫代發模式需要選擇已入庫的包裹"})
        package_ids = data["package_ids"]
    else:
        forecast_data = {
            "packages": [{
                "local_logis_num": data["customer_order_id"],
                "client_cid": data["customer_order_id"],
                "client_pid": data["customer_order_id"],
                "client_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "warehouse_id": JPD_WAREHOUSE_ID,
                "product_name": declare_list[0]["product_name"] if declare_list else "商品",
                "product_num": total_num,
                "product_price": total_price,
                "declare_list": declare_list
            }]
        }
        
        forecast_result = jpd_request("TForecastPackage", forecast_data)
        
        if "OperationResult" in forecast_result:
            op_result = forecast_result["OperationResult"]
            if op_result["Request"]["IsValid"] == "True":
                result_data = op_result.get("Result", {})
                if result_data.get("Result") == "SUCCESS":
                    packages_data = result_data.get("Data", [])
                    for pkg in packages_data:
                        if pkg.get("package_id"):
                            package_ids.append(pkg["package_id"])
                else:
                    return jsonify({
                        "success": False, 
                        "error": f"預報包裹失敗: {result_data.get('Data', {}).get('msg', '未知錯誤')}"
                    })
            else:
                errors = op_result["Request"].get("Errors", {})
                return jsonify({"success": False, "error": f"預報包裹失敗: {errors}"})
        else:
            return jsonify({"success": False, "error": "預報包裹 API 回應異常"})
        
        if not package_ids:
            return jsonify({"success": False, "error": "預報包裹失敗：未取得 package_id"})
    
    recipient = data["recipient"]
    shopify_order_id = data.get("shopify_order_id")
    if shopify_order_id:
        order_detail = shopify_request(f"orders/{shopify_order_id}.json")
        if "order" in order_detail:
            orig_order = order_detail["order"]
            orig_shipping = orig_order.get("shipping_address", {}) or {}
            orig_customer = orig_order.get("customer", {}) or {}
            orig_billing  = orig_order.get("billing_address", {}) or {}
            
            invalid_names = {"本人", "本人本人", "本人 本人", "同上", "同收件人", "test", "測試", ".", "-", ""}
            
            def is_valid_name(name):
                return name and name.strip() not in invalid_names
            
            s_last  = (orig_shipping.get("last_name")  or "").strip()
            s_first = (orig_shipping.get("first_name") or "").strip()
            shipping_combined = f"{s_last}{s_first}".strip()
            
            c_last  = (orig_customer.get("last_name")  or "").strip()
            c_first = (orig_customer.get("first_name") or "").strip()
            customer_combined = f"{c_last}{c_first}".strip()
            
            b_last  = (orig_billing.get("last_name")  or "").strip()
            b_first = (orig_billing.get("first_name") or "").strip()
            billing_combined = f"{b_last}{b_first}".strip()
            
            if is_valid_name(shipping_combined):
                recipient = shipping_combined
            elif is_valid_name(customer_combined):
                recipient = customer_combined
            elif is_valid_name(billing_combined):
                recipient = billing_combined
            
            print(f"📝 JPD 收件人: '{recipient}' (原始: shipping={shipping_combined}, customer={customer_combined})")
    
    order_data = {
        "customer_order_id": data["customer_order_id"],
        "deliv_id": JPD_DELIV_ID,
        "recipient": recipient,
        "id_issure": "",
        "area": 3,
        "addr1": data["address"],
        "addr2": "",
        "addr3": "",
        "addr4": "",
        "tel": data["phone"],
        "memo": data.get("memo", ""),
        "create_order_pdf": "y",
        "warehouse_id": JPD_WAREHOUSE_ID,
        "create_package": "n",
        "create_sender": "y",
        "packages": [{"package_id": int(pid), "declare_list": declare_list} for pid in package_ids]
    }
    
    result = jpd_request("TCreateOrder", order_data)
    
    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            result_data = op_result["Result"]
            if result_data.get("Result") == "SUCCESS":
                jpd_data = result_data.get("Data", {})
                return jsonify({
                    "success": True,
                    "order_id": jpd_data.get("order_id"),
                    "logis_num": jpd_data.get("logis_num"),
                    "message": "運單創建成功"
                })
            else:
                return jsonify({
                    "success": False,
                    "error": result_data.get("Data", {}).get("msg", "創建失敗")
                })
        else:
            errors = op_result["Request"].get("Errors", {})
            error_list = errors.get("Error", [])
            if isinstance(error_list, dict):
                error_list = [error_list]
            
            is_duplicate = any("已存在" in str(e.get("Message", "")) for e in error_list)
            
            if is_duplicate:
                search_result = jpd_request("TSearchOrders", {
                    "customer_order_id": data["customer_order_id"]
                })
                
                order_info = {}
                if "OperationResult" in search_result:
                    search_op = search_result["OperationResult"]
                    if search_op["Request"]["IsValid"] == "True":
                        search_data = search_op.get("Result", {}).get("Data", [])
                        if search_data and len(search_data) > 0:
                            order_info = search_data[0]
                
                return jsonify({
                    "success": True,
                    "order_id": order_info.get("order_id", ""),
                    "logis_num": order_info.get("logis_num", ""),
                    "message": "此運單已存在，無需重複建立"
                })
            
            return jsonify({"success": False, "error": str(errors)})
    
    return jsonify({"success": False, "error": "API 回應異常"})


@app.route("/api/jpd/confirm_order", methods=["POST"])
def confirm_jpd_order():
    """確定發貨"""
    data = request.json
    customer_order_id = data.get("customer_order_id")
    
    result = jpd_request("TConfirmOrder", {
        "customer_order_id": customer_order_id
    })
    
    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            result_data = op_result["Result"]
            if result_data.get("Result") == "SUCCESS":
                return jsonify({"success": True, "message": "確定發貨成功"})
    
    return jsonify({"success": False, "error": "確定發貨失敗"})


@app.route("/api/jpd/cancel_order", methods=["POST"])
def cancel_jpd_order():
    """取消訂單"""
    data = request.json
    customer_order_id = data.get("customer_order_id")
    
    result = jpd_request("TDeleteOrder", {
        "customer_order_id": customer_order_id
    })
    
    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            result_data = op_result["Result"]
            if result_data.get("Result") == "SUCCESS":
                return jsonify({"success": True, "message": "訂單取消成功"})
    
    return jsonify({"success": False, "error": "取消訂單失敗"})


@app.route("/api/shopify/fulfill", methods=["POST"])
def fulfill_shopify_order():
    """回寫 Shopify 出貨資訊"""
    data = request.json
    order_id = data.get("shopify_order_id")
    tracking_number = data.get("tracking_number")
    
    print(f"\n{'='*50}")
    print(f"📝 回寫 Shopify 訂單: {order_id}")
    print(f"📦 追蹤號: {tracking_number}")
    
    fo_result = shopify_request(f"orders/{order_id}/fulfillment_orders.json")
    print(f"📥 Fulfillment Orders: {json.dumps(fo_result, ensure_ascii=False)[:500]}")
    
    if "fulfillment_orders" not in fo_result:
        return jsonify({"success": False, "error": "無法取得訂單資訊"})
    
    for fo in fo_result["fulfillment_orders"]:
        if fo["status"] in ["open", "in_progress"]:
            fulfill_data = {
                "fulfillment": {
                    "line_items_by_fulfillment_order": [
                        {
                            "fulfillment_order_id": fo["id"]
                        }
                    ],
                    "tracking_info": {
                        "number": tracking_number,
                        "company": "SG 速貴專線",
                        "url": f"https://www.sgxpress.com/query/?logic_num={tracking_number}"
                    },
                    "notify_customer": True
                }
            }
            
            print(f"📤 Fulfillment 請求: {json.dumps(fulfill_data, ensure_ascii=False)}")
            
            fulfill_result = shopify_request("fulfillments.json", "POST", fulfill_data)
            print(f"📥 Fulfillment 回應: {json.dumps(fulfill_result, ensure_ascii=False)[:500]}")
            
            if "fulfillment" in fulfill_result:
                return jsonify({
                    "success": True, 
                    "message": "出貨資訊已回寫 Shopify",
                    "fulfillment_id": fulfill_result["fulfillment"]["id"]
                })
            else:
                error_msg = fulfill_result.get("errors") or fulfill_result.get("error") or str(fulfill_result)
                return jsonify({"success": False, "error": f"回寫失敗: {error_msg}"})
    
    return jsonify({"success": False, "error": "找不到可出貨的訂單項目（可能已出貨）"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"""
    ╔═══════════════════════════════════════════════════════════╗
    ║       Shopify × JPD 雲倉 串接工具                         ║
    ║       御用達-光頭哥 專用                                   ║
    ╚═══════════════════════════════════════════════════════════╝
    
    🌐 請打開瀏覽器訪問: http://localhost:{port}
    
    按 Ctrl+C 停止服務
    """)
    app.run(debug=True, host="0.0.0.0", port=port)
