# 2026-09 REAL affinity 事件記錄

`shipment_requests` 七個欄位被宣告成 REAL，SQLite 自動把字串轉成數字，
前導零與 `+` 號被永久吃掉。已修復，記錄如下。

---

## 現象

`PRAGMA table_info(shipment_requests)` 顯示以下七個欄位的型別是 **REAL**：

```
payment_last5, payment_at, tracking_num, extra_services,
ship_recipient, ship_phone, ship_address
```

但 `app.py` 的 ALTER 遷移清單裡，它們全部寫的是 `TEXT`。
**線上 schema 不是這份程式碼建出來的** —— 這是整起事件所有判斷的前提。

SQLite 的 type affinity：REAL 欄位收到「看起來像數字」的字串會自動轉成數字。
所以：

| 寫入 | 實際存成 | typeof |
|---|---|---|
| `'0912345678'` | `912345678.0` | real |
| `'00463'` | `463.0` | real |
| `'+886912345678'` | `886912345678.0` | real |
| `'0912-345-999'` | `'0912-345-999'` | text（有橫線，逃過一劫）|

三條寫入路徑的 `zfill(5)` 全部失效 —— 補好的零在存進去那一刻就被吃掉了。

## 影響

| 欄位 | 受損筆數 | 說明 |
|---|---|---|
| `ship_phone` | 833 | 開頭 `0` 被吃掉變 9 碼。**已流入交給清關行的出貨單** |
| `payment_last5` | 100 | 前導零被吃掉（`00463` → `463`）|
| `tracking_num` | 3 | 測試資料 |

其餘四欄（`payment_at` / `extra_services` / `ship_recipient` / `ship_address`）
值本身是日期、JSON、中文，不會被判定成數字，所以沒有實際損害，但型別仍須修正
以免未來寫入被轉型。

`members` / `agent_payouts` / `addresses` / `declarants` 經全庫掃描確認乾淨，未受影響。

## 根因

某次遷移迴圈的型別帶錯，把這一批欄位加成了 REAL。程式碼後來修成 TEXT，
但 **`ALTER TABLE ADD COLUMN` 只在欄位不存在時才會執行** ——
既有欄位不會因為程式碼改了就自己變回去。線上 schema 從此與程式碼長期不一致。

## 修復

分兩步，各自可獨立驗證：

**任務 J2 — 型別修正（重建表）**
從 `sqlite_master` 讀線上實際建表 SQL，逐欄只替換型別 token `REAL → TEXT`，
DEFAULT 與其他修飾一字不動；搬資料時只去除型別造成的 `.0`（`463.0` → `'463'`），
**不補零**。備份用 SQLite 官方 backup API（WAL 安全），
在同一個 transaction 內驗證列數、id 集合、其他欄位逐欄比對、typeof、去 `.0` 正確性，
任一不過就 ROLLBACK。

**任務 M — 前導零還原**
型別修好也救不回前導零（是寫入當下就沒了），依三條規則還原：

- **R1** `payment_last5` 純數字且長度 1~4 → `zfill(5)`
- **R2** `ship_phone` 長度 9、純數字、開頭 `9` → 前面補 `0`
- **R3** `ship_phone` `886`/`+886` 開頭且其後為 9 位開頭 `9` → 前綴換 `0`

實際執行結果：**R1=100 R2=833 R3=4**。

**正確性驗證：`operation_logs`**
管理員確認付款會寫 `log_op("帳單確認付款", "出貨單#N", "後五碼 XXXXX")`，
而 `operation_logs.detail` 是真正的 TEXT，**原始輸入含前導零完整保存**。
拿它當 ground truth 比對修復後的 `payment_last5`，`mismatch = 0`。
這是本次修復正確性的最終證明，不是靠推斷。

`normalize_phone` 也補了一條：沒有 `+` 的 `886` 前綴（affinity 把 `+` 吃掉後殘留的）
→ 換成 `0` 開頭。`+81` 規則維持原樣未動。

## 未處理（刻意保留）

這些**故意留著錯誤外觀**，讓客服看到會去問客人，不要自作聰明補值：

| 位置 | 值 | 為什麼不動 |
|---|---|---|
| `ship_phone` id=664 (G0387) | `97912928` | 8 碼，補 0 只有 9 碼仍是錯的；原始輸入本來就少一碼 |
| `ship_phone` id=259 (G0020) | `93529536` | 同上 |
| `ship_phone` id=467 | `Hj` | 非數字 |
| `ship_phone` id=115 (G0202) | `Hj` | 非數字 |
| `ship_phone` id=490 | 全形數字電話 | 需人工確認 |
| `ship_phone` id=725 | 含空格的電話 | 需人工確認 |

Schema 層面尚未處理：

- `contact_book.mobile_no` 宣告為 REAL
- `contact_book.misc` 無型別宣告（BLOB affinity）

這兩個目前沒有實際損害，但同樣是誤宣告，日後若要存字串進去會踩同一個坑。

## 留下的工具

三個**唯讀**診斷端點保留在正式站，下次懷疑資料有問題時直接用（`is_boss` 限定）：

| 端點 | 用途 |
|---|---|
| `GET /api/admin/maintenance/last5_diag` | `payment_last5` 儲存型別、時間軸、模糊樣本、從 `operation_logs` 還原原值 |
| `GET /api/admin/maintenance/affinity_diag` | 七欄損害範圍、`ship_phone` 長度分布、超長 `tracking_num`、全庫可疑欄位 |
| `GET /api/admin/maintenance/phone_affinity_diag` | 全庫電話欄位 affinity、`members` 細看、登入比對邏輯說明 |

一次性的**寫入**工具（`clean_last5` / `fix_column_types` / `restore_leading_zeros`）
在修復完成後已移除 —— 能重建或改寫資料表的東西沒有理由長期掛在正式站上。
需要時從 git 歷史取回（見 `chore(maintenance): 移除一次性資料修復端點` 之前的版本）。

相容處理 `fmtLast5` / `hasPayment` / `_norm_pay_mark` / admin 端去 `.0` 防禦
**永久保留**，即使資料已清乾淨。

---

## 給未來的人

**新增欄位後，務必用 `PRAGMA table_info` 確認實際型別，不要相信程式碼裡的 ALTER 清單寫什麼。**

```sql
SELECT name, type, dflt_value FROM pragma_table_info('你的表名');
```

`ALTER TABLE ADD COLUMN` 只在欄位不存在時執行，所以程式碼與線上 schema 一旦分岔就不會自己收斂。

其他幾個踩過的坑：

- **REAL 欄位存不出 REAL 以外的東西**：TEXT affinity 會把數字轉成字串，
  REAL affinity 會把數字字串轉成數字。宣告型別 ≠ 實際 storage class，
  用 `typeof(欄位)` 才看得到真相。
- **JSON 序列化會吃掉型別資訊**：REAL 值經 `jsonify` 就變成數字，
  在瀏覽器看不出異狀。診斷工具一律 `CAST(... AS TEXT)` 再輸出，並另外附上 `typeof`。
- **超長數字進 REAL 會失去精度且不可逆**：19 碼的追蹤號會變成 `1.23456789012346e+18`，
  尾端數字永久消失，補零救不回來。
- **WAL 模式下 `shutil.copy2` 備份不完整**：`commit()` 只保證寫進 `-wal` 旁檔，
  不保證 checkpoint 回主檔。備份一律用 `conn.backup(dest)`（SQLite 官方 backup API）。
- **備份檔名的時間戳只到秒**：同一秒內跑兩次會撞名並蓋掉前一份備份，撞名要加序號。
