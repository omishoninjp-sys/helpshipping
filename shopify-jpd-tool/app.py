#!/usr/bin/env python3
"""
Shopify × JPD 雲倉 串接工具
御用達-光頭哥 專用
"""

from flask import Flask, render_template, request, jsonify
import requests
import json
import os
from datetime import datetime

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
        # SSL 錯誤時嘗試不驗證（僅限本地開發使用）
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
    
    # 除錯：輸出請求內容
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
    status = request.args.get("status", "unfulfilled")  # unfulfilled, any, fulfilled
    limit = request.args.get("limit", 50)
    
    result = shopify_request(f"orders.json?status=any&fulfillment_status={status}&limit={limit}")
    
    # 除錯：輸出 Shopify 回應
    print(f"\n{'='*50}")
    print(f"📦 Shopify API 回應:")
    print(json.dumps(result, ensure_ascii=False, indent=2)[:1000])  # 只印前1000字元
    
    if "orders" in result:
        orders = []
        for order in result["orders"]:
            shipping = order.get("shipping_address", {}) or {}
            customer_info = order.get("customer", {}) or {}
            billing = order.get("billing_address", {}) or {}
            
            # ===== 除錯：印出所有姓名來源 =====
            print(f"\n--- 訂單 {order.get('name', '?')} 姓名來源 ---")
            print(f"  shipping.name        = '{shipping.get('name', '')}'")
            print(f"  shipping.first_name  = '{shipping.get('first_name', '')}'")
            print(f"  shipping.last_name   = '{shipping.get('last_name', '')}'")
            print(f"  customer.first_name  = '{customer_info.get('first_name', '')}'")
            print(f"  customer.last_name   = '{customer_info.get('last_name', '')}'")
            print(f"  billing.name         = '{billing.get('name', '')}'")
            print(f"  order.contact_email  = '{order.get('contact_email', '')}'")
            # ===================================
            
            # 無效姓名（結帳時常見的佔位字）
            invalid_names = {"本人", "本人本人", "本人 本人", "同上", "同收件人", "test", "測試", ".", "-", ""}
            
            def is_valid(name):
                return name and name.strip() not in invalid_names
            
            # 收件人姓名判斷（多重來源 fallback）
            # 來源 1: shipping_address.name（Shopify 自動組合的完整名字）
            shipping_name = shipping.get("name", "").strip()
            
            # 來源 2: shipping_address.last_name + first_name（自己拼接）
            s_last = shipping.get("last_name", "").strip()
            s_first = shipping.get("first_name", "").strip()
            shipping_combined = f"{s_last}{s_first}".strip()
            
            # 來源 3: customer 物件
            c_last = customer_info.get("last_name", "").strip()
            c_first = customer_info.get("first_name", "").strip()
            customer_combined = f"{c_last}{c_first}".strip()
            
            # 來源 4: billing_address.name
            billing_name = billing.get("name", "").strip()
            
            # 優先順序判斷
            if is_valid(shipping_name):
                customer_name = shipping_name
            elif is_valid(shipping_combined):
                customer_name = shipping_combined
            elif is_valid(customer_combined):
                customer_name = customer_combined
            elif is_valid(billing_name):
                customer_name = billing_name
            else:
                # 最後 fallback
                customer_name = shipping_name or shipping_combined or customer_combined or "N/A"
            
            print(f"  ➡️ 最終使用: '{customer_name}'")
            
            orders.append({
                "id": order["id"],
                "order_number": order["order_number"],
                "name": order["name"],  # #1001 格式
                "created_at": order["created_at"],
                "total_price": order["total_price"],
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
                "line_items": [
                    {
                        "title": item["title"],
                        "variant_title": item.get("variant_title", ""),
                        "quantity": item["quantity"],
                        "price": item["price"],
                        "sku": item.get("sku", "")
                    }
                    for item in order["line_items"]
                ]
            })
        return jsonify({"success": True, "orders": orders})
    
    # 回傳更詳細的錯誤資訊
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
    """取得 JPD 倉庫的包裹列表"""
    # 查詢最近入庫的包裹
    result = jpd_request("TSearchPackages", {
        "stock_date_from": (datetime.now().replace(day=1)).strftime("%Y-%m-%d 00:00:00")
    })
    
    if "OperationResult" in result:
        op_result = result["OperationResult"]
        if op_result["Request"]["IsValid"] == "True":
            packages = op_result["Result"].get("Data", [])
            return jsonify({"success": True, "packages": packages})
        else:
            errors = op_result["Request"].get("Errors", {})
            return jsonify({"success": False, "error": errors})
    
    return jsonify({"success": False, "error": "Unknown error"})


@app.route("/api/jpd/orders")
def get_jpd_orders():
    """取得 JPD 運單列表"""
    days = request.args.get("days", 7, type=int)
    
    result = jpd_request("TSearchOrders", {
        "create_date": datetime.now().strftime("%Y-%m-%d")
    })
    
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
    
    # 組裝申報列表
    declare_list = []
    for item in data.get("declare_list", []):
        declare_list.append({
            "product_name": item.get("product_name", "商品"),
            "product_name_local": item.get("product_name_local", item.get("product_name", "商品")),
            "product_num": int(item.get("product_num", 1)),
            "product_price": int(item.get("product_price", 100))
        })
    
    # 計算總數量和總價
    total_num = sum(int(item.get("product_num", 1)) for item in data.get("declare_list", []))
    total_price = sum(int(item.get("product_price", 0)) * int(item.get("product_num", 1)) for item in data.get("declare_list", []))
    
    package_ids = []
    
    if mode == "warehouse":
        # 倉庫代發：使用已入庫的包裹
        if not data.get("package_ids"):
            return jsonify({"success": False, "error": "倉庫代發模式需要選擇已入庫的包裹"})
        package_ids = data["package_ids"]
    else:
        # 自出貨：先預報包裹
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
                    # 取得預報成功的 package_id
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
    
    # 組裝運單資料
    order_data = {
        "customer_order_id": data["customer_order_id"],
        "deliv_id": JPD_DELIV_ID,
        "recipient": data["recipient"],
        "id_issure": "",
        "area": 3,  # 台灣
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
    
    # 呼叫 JPD API 創建運單
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
            # 檢查是否是「已存在」的錯誤
            errors = op_result["Request"].get("Errors", {})
            error_list = errors.get("Error", [])
            if isinstance(error_list, dict):
                error_list = [error_list]
            
            is_duplicate = any("已存在" in str(e.get("Message", "")) for e in error_list)
            
            if is_duplicate:
                # 運單已存在，視為成功（可能是重複提交）
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
    
    # 取得該訂單的 fulfillment orders
    fo_result = shopify_request(f"orders/{order_id}/fulfillment_orders.json")
    print(f"📥 Fulfillment Orders: {json.dumps(fo_result, ensure_ascii=False)[:500]}")
    
    if "fulfillment_orders" not in fo_result:
        return jsonify({"success": False, "error": "無法取得訂單資訊"})
    
    # 找到狀態為 open 或 in_progress 的 fulfillment order
    for fo in fo_result["fulfillment_orders"]:
        if fo["status"] in ["open", "in_progress"]:
            # 組裝 fulfillment 請求
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
