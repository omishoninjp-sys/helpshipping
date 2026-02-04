# 客人集運預報系統

御用達 × JPD 雲倉

## Zeabur 部署步驟

### 1. 上傳到 GitHub
將此資料夾上傳到你的 GitHub 倉庫

### 2. Zeabur 建立服務
- 登入 Zeabur
- 新增專案 → 從 GitHub 導入
- 選擇此倉庫

### 3. 設定環境變數
在 Zeabur 的 Variables 頁面設定：

```
JPD_EMAIL=你的JPD帳號
JPD_PASSWORD=你的JPD密碼
JPD_WAREHOUSE_ID=1
SHOPIFY_STORE=你的商店.myshopify.com
SHOPIFY_ACCESS_TOKEN=shpat_xxxxx
```

### 4. 綁定網域
在 Networking 頁面綁定你的網域，例如：
- `forecast.goyoutati.com`

## 功能說明

### 客人端功能
1. **登入** - 用 Shopify Customer ID 登入
2. **預報包裹** - 填寫商品資訊後預報到 JPD
3. **我的包裹** - 查看包裹狀態、重量
4. **運單查詢** - 查看已發貨運單、追蹤

### API 端點
- `POST /api/verify_customer` - 驗證客戶
- `POST /api/forecast` - 建立預報
- `GET /api/packages?customer_id=xxx` - 查詢包裹
- `GET /api/orders?customer_id=xxx` - 查詢運單

## 本地測試

```bash
# 設定環境變數
export JPD_EMAIL=xxx
export JPD_PASSWORD=xxx
export SHOPIFY_STORE=xxx.myshopify.com
export SHOPIFY_ACCESS_TOKEN=shpat_xxx

# 啟動
python app.py
```

開啟 http://localhost:5001
