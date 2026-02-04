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


def shopify_request(endpoint: str) -> dict:
    """Shopify API 請求"""
    url = f"https://{SHOPIFY_STORE}/admin/api/2026-01/{endpoint}"
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        return response.json()
    except Exception as e:
        return {"error": str(e)}


@app.route("/")
def index():
    """首頁 - 客人預報表單"""
    return render_template("index.html")


@app.route("/api/verify_customer", methods=["POST"])
def verify_customer():
    """驗證客戶 ID"""
    data = request.json
    customer_id = data.get("customer_id", "").strip()
    
    if not customer_id:
        return jsonify({"success": False, "error": "請輸入客戶編號"})
    
    # 呼叫 Shopify API 驗證客戶
    result = shopify_request(f"customers/{customer_id}.json")
    
    if "customer" in result:
        customer = result["customer"]
        return jsonify({
            "success": True,
            "customer": {
                "id": customer["id"],
                "name": f"{customer.get('last_name', '')}{customer.get('first_name', '')}".strip() or customer.get("email", ""),
                "email": customer.get("email", ""),
                "phone": customer.get("phone", "")
            }
        })
    
    return jsonify({"success": False, "error": "找不到此客戶編號，請確認後重試"})


@app.route("/api/forecast", methods=["POST"])
def create_forecast():
    """建立預報包裹"""
    data = request.json
    
    customer_id = data.get("customer_id")
    packages = data.get("packages", [])
    
    if not customer_id:
        return jsonify({"success": False, "error": "缺少客戶編號"})
    
    if not packages:
        return jsonify({"success": False, "error": "請至少填寫一個包裹"})
    
    results = []
    
    for idx, pkg in enumerate(packages):
        # 產生預報編號：客戶ID + 日期 + 序號
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        local_logis_num = f"{customer_id}-{timestamp}-{idx+1}"
        
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
                "client_cid": str(customer_id),  # 客戶會員 ID
                "client_pid": pkg.get("client_pid") or local_logis_num,  # 客戶自訂包裹編號
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
    customer_id = request.args.get("customer_id")
    
    if not customer_id:
        return jsonify({"success": False, "error": "缺少客戶編號"})
    
    # 查詢該客戶的包裹
    result = jpd_request("TSearchPackages", {
        "client_cid": str(customer_id)
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
                    "in_date": pkg.get("in_date"),  # 入庫日期
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
    customer_id = request.args.get("customer_id")
    
    if not customer_id:
        return jsonify({"success": False, "error": "缺少客戶編號"})
    
    # 查詢該客戶的運單
    result = jpd_request("TSearchOrders", {
        "client_cid": str(customer_id)
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
