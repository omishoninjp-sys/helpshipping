"""
廠商出貨檔案範本（config-driven）

每個廠商一份 dict，描述：
  - id: 內部識別字串（'nigel' / 'jpd'）
  - display_name: UI 顯示用
  - filename_template: 下載檔名格式
  - columns: Excel 欄位定義（header + getter）
  - row_strategy: 'one_per_package'（一個包裹一行）

新增廠商只要加一份 dict 到 VENDORS 即可，不用改其他 code。
"""
from __future__ import annotations
import random
import re
from datetime import datetime


# ============ 隨機資料產生器（出檔案一律用白名單隨機，不使用客戶預報資料） ============

# 唯一允許出現在廠商 Excel 的品名（客戶預報的真實品名不外流）
_PRODUCT_NAME_POOL = [
    "玩具", "糖果", "上衣", "文具用品", "髮夾", "廚房用品", "貴鞋子",
    "化妝品", "娃娃", "包包", "毛巾", "保健食品", "便宜鞋子", "飾品－A",
    "吊飾", "水壺", "襪子", "蝦餅", "口罩", "小朋友涼鞋", "飾品－Ｂ",
    "沐浴球", "卡片",
]
_ORIGIN_POOL = ["japan"] * 6 + ["china"] * 4  # 60% japan, 40% china

MAX_QTY = 10  # 每個品項數量上限


def _seed_for(shipment_id: int, package_id: int = 0) -> random.Random:
    """每筆出貨用固定 seed → 同一筆每次匯出結果一樣（避免亂跳）"""
    return random.Random(f"{shipment_id}-{package_id}")


def _random_price(rng) -> int:
    """單品 JPY 申報價：200~2000、整百"""
    return rng.randint(2, 20) * 100


def build_item_plan(shipment_id: int, n: int) -> list[dict]:
    """為一筆出貨單配置 n 個品項（每列一個）。

    規則：
      • 品名只從白名單取
      • 同一筆出貨單內【品名不重複】→ 自然不會有「同品名不同價格」
      • 每個品名的單價固定（同名必同價）
      • 數量 1~MAX_QTY(10)
      • 用固定 seed → 同一筆每次匯出結果一致
    n 若超過白名單長度（極少見），才循環重複品名，且重複時沿用同一單價。
    """
    rng = _seed_for(shipment_id, 0)
    pool = list(_PRODUCT_NAME_POOL)
    rng.shuffle(pool)

    plan = []
    price_by_name = {}
    for i in range(max(n, 0)):
        name = pool[i % len(pool)]           # 不重複；超過池長才循環
        if name not in price_by_name:        # 同名必同價
            price_by_name[name] = _random_price(rng)
        plan.append({
            "name": name,
            "price": price_by_name[name],
            "qty": rng.randint(1, MAX_QTY),
            "origin": rng.choice(_ORIGIN_POOL),
        })
    return plan


# ============ 廠商範本 ============

NIGEL = {
    "id": "nigel",
    "display_name": "Nigel",
    "filename_template": "{date}_Nigel_出貨單.xlsx",
    "row_strategy": "one_per_package",
    "columns": [
        # (header, getter — 接收 ctx dict 回字串/數字)
        ("客戶編號",       lambda ctx: f"{ctx['g_code']}-{ctx['packaging_mmdd']}"),
        ("清關號碼",       lambda ctx: ctx["tracking_num"]),  # 出貨追蹤號碼（多箱換行）
        ("收件人",         lambda ctx: ctx["ship_recipient"]),
        ("收件人詳細地址", lambda ctx: ctx["ship_address"]),
        ("收件人電話號碼", lambda ctx: ctx["ship_phone"]),
        ("申報人",         lambda ctx: ctx["ship_recipient"]),         # 預設 = 收件人
        ("申報人詳細地址", lambda ctx: ctx["ship_address"]),
        ("申報人電話號碼", lambda ctx: ctx["ship_phone"]),
        ("品名",           lambda ctx: ctx["item"]["name"]),
        ("數量",           lambda ctx: ctx["item"]["qty"]),
        ("金額",           lambda ctx: ctx["item"]["price"]),
        ("產地",           lambda ctx: ctx["item"]["origin"]),
        ("URL/JanCode",    lambda ctx: ""),
    ],
}


JPD = {
    "id": "jpd",
    "display_name": "JpD 集運（小客戶大價值）",
    "filename_template": "{date}_JpD_出貨單.xlsx",
    "row_strategy": "one_per_package",
    "columns": [
        ("客戶運單號",      lambda ctx: f"{ctx['g_code']}-{ctx['packaging_mmdd']}"),
        ("JpD包裹ID",       lambda ctx: ctx["tracking_num"]),  # 出貨追蹤號碼（多箱換行）
        ("運單ID",          lambda ctx: ""),
        ("包裹特殊服務",     lambda ctx: ""),
        ("收件人",          lambda ctx: ctx["ship_recipient"]),
        ("收件人身份證ID",   lambda ctx: ""),  # 不需要
        ("收件人詳細地址",   lambda ctx: ctx["ship_address"]),
        ("收件人电话号码",   lambda ctx: ctx["ship_phone"]),
        ("備註",            lambda ctx: ""),
        ("特殊服务",         lambda ctx: ""),
        ("渠道ID",          lambda ctx: 40),  # JpD 固定 40
        ("申報人",          lambda ctx: ctx["ship_recipient"]),
        ("申報人身份證ID",   lambda ctx: ""),
        ("申報人詳細地址",   lambda ctx: ctx["ship_address"]),
        ("申報人电话号码",   lambda ctx: ctx["ship_phone"]),
        ("品名",            lambda ctx: ctx["item"]["name"]),
        ("数量",            lambda ctx: ctx["item"]["qty"]),
        ("金额",            lambda ctx: ctx["item"]["price"]),
        ("材質",            lambda ctx: ""),
        ("產地",            lambda ctx: ctx["item"]["origin"]),
        ("URL/JanCode",     lambda ctx: ""),
    ],
}


VENDORS = {
    "nigel": NIGEL,
    "jpd": JPD,
}


def list_vendors() -> list[dict]:
    """給前端 UI 顯示用的廠商清單"""
    return [{"id": v["id"], "display_name": v["display_name"]} for v in VENDORS.values()]


def get_vendor(vendor_id: str) -> dict | None:
    return VENDORS.get(vendor_id)


def _packaging_mmdd(s: dict) -> str:
    """打包日 MMDD：updated_at（標記已出貨那刻）> created_at（客戶申請）> 今天。"""
    for src in (s.get("updated_at"), s.get("created_at")):
        if src and len(str(src)) >= 10:
            try:
                return datetime.strptime(str(src)[:10], "%Y-%m-%d").strftime("%m%d")
            except ValueError:
                pass
    return datetime.now().strftime("%m%d")


def export_code_for(s: dict) -> str:
    """出檔案給廠商的客戶編號 = {g_code}-{MMDD}（與 Excel 內容一致，供台灣配送貨況比對）。"""
    return f"{s.get('g_code', '')}-{_packaging_mmdd(s)}"


def build_rows(vendor_id: str, shipments: list[dict]) -> tuple[list[str], list[list]]:
    """
    依範本產生 (headers, rows)。
    shipments 結構：每個 dict 包含 shipment 本身 + 該 shipment 內的 packages list。

    Returns:
      headers: List[str] — 第一列標頭
      rows:    List[List] — 資料列（每個 package 一列）
    """
    vendor = get_vendor(vendor_id)
    if not vendor:
        raise ValueError(f"unknown vendor: {vendor_id}")

    headers = [col[0] for col in vendor["columns"]]
    rows = []

    for s in shipments:
        # packaging_mmdd = 客戶申請出單（admin 打包）的日期 → MMDD
        # 來源優先：updated_at（admin 標記已出貨那刻）> created_at（客戶申請時）> 今天
        packaging_mmdd = _packaging_mmdd(s)

        # 出貨追蹤號碼正規化：把換行／半形逗號／全形逗號／頓號都統一成換行分隔，去空行
        # 多箱 → 多行（與後台「多箱請換行」一致）；單箱 → 單一字串
        tracking_num = "\n".join(
            t.strip()
            for t in re.split(r"[\n,，、]+", str(s.get("tracking_num") or ""))
            if t.strip()
        )

        # 每筆出貨單先配置品項（白名單、同筆內品名不重複、同名同價、數量≤10）
        item_plan = build_item_plan(s["id"], len(s["packages"]))

        for pkg_index, pkg in enumerate(s["packages"]):
            ctx = {
                "shipment_id":          s["id"],
                "g_code":               s["g_code"],
                "packaging_mmdd":       packaging_mmdd,   # 客戶申請出單日（打包日）MMDD
                "tracking_num":         tracking_num,     # 出貨追蹤號碼（Nigel→清關號碼 / JpD→JpD包裹ID）
                # 品名/數量/金額/產地一律來自白名單配置（不使用客戶預報資料）
                "item":                 item_plan[pkg_index],
                "pkg_index":            pkg_index,
                # str() 防 DB 把 phone 存成 float（912345678.0）造成下游 .strip() 炸
                "ship_recipient":       str(s["ship_recipient"]) if s["ship_recipient"] else "",
                "ship_address":         str(s["ship_address"]) if s["ship_address"] else "",
                "ship_phone":           str(s["ship_phone"]) if s["ship_phone"] else "",
                "billed_weight":        s.get("billed_weight") or 0,
                "total_fee":            s.get("total_fee") or 0,
                "package_id":           pkg["id"],
                "package_logis_num":    pkg.get("logis_num") or "",
                "package_product_name": pkg.get("product_name") or "",
                "package_weight":       pkg.get("weight") or 0,
            }
            row = [getter(ctx) for _, getter in vendor["columns"]]
            rows.append(row)

    return headers, rows


def filename_for(vendor_id: str, when: datetime | None = None) -> str:
    vendor = get_vendor(vendor_id)
    if not vendor:
        return "export.xlsx"
    when = when or datetime.now()
    return vendor["filename_template"].format(date=when.strftime("%Y%m%d_%H%M%S"))
