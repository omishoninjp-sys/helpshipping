"""
廠商出貨檔案範本（config-driven）

shopify-jpd-tool 專用版本，跟 helpshipping 的 vendors.py 同一個設計，
但因為這邊資料模型不同（沒有會員編號 g_code），改用 Shopify 訂單號當識別。

每個廠商一份 dict，描述：
  - id:               內部識別字串（'nigel' / 'jpd'）
  - display_name:     UI 顯示用
  - filename_template: 下載檔名格式
  - columns:          Excel 欄位定義（header + getter）
  - row_strategy:     'one_per_package'（一個包裹一行）

客戶編號規則：`{customer_order_id}-{MMDD}`
  例如：1234-0616 表示 Shopify 訂單 #1234，於 6/16 建單。
  MMDD 來源：order_history.created_at（建立 JPD 運單那刻）。

新增廠商只要加一份 dict 到 VENDORS 即可，不用改其他 code。
"""
from __future__ import annotations
import random
from datetime import datetime


# ============ 隨機資料產生器 ============

_PRODUCT_NAME_POOL = [
    "雑貨", "衣類", "玩具", "鞋子", "包包", "化妝品", "食品",
    "日用品", "餅乾", "髮夾", "上衣", "褲子", "外套", "手帕",
    "毛巾", "資料夾", "杯子", "公仔", "面膜", "卡片",
]
_ORIGIN_POOL = ["japan"] * 6 + ["china"] * 4  # 60% japan, 40% china


def _seed_for(order_id: int, package_id: int = 0) -> random.Random:
    """每筆出貨用固定 seed → 同一筆每次匯出結果一樣（避免亂跳）"""
    return random.Random(f"jpd-{order_id}-{package_id}")


def random_product_name(order_id: int, package_id: int = 0) -> str:
    return _seed_for(order_id, package_id).choice(_PRODUCT_NAME_POOL)


def random_quantity(order_id: int, package_id: int = 0) -> int:
    return _seed_for(order_id, package_id + 1).randint(1, 30)


def random_jpy_amount(order_id: int, package_id: int = 0) -> int:
    """單品 JPY 申報價：200~2000、四捨五入到百"""
    rng = _seed_for(order_id, package_id + 2)
    return rng.randint(2, 20) * 100


def random_origin(order_id: int, package_id: int = 0) -> str:
    return _seed_for(order_id, package_id + 3).choice(_ORIGIN_POOL)


# ============ 廠商範本 ============

NIGEL = {
    "id": "nigel",
    "display_name": "Nigel",
    "filename_template": "{date}_Nigel_出貨單.xlsx",
    "row_strategy": "one_per_package",
    "columns": [
        # (header, getter — 接收 ctx dict 回字串/數字)
        ("客戶編號",       lambda ctx: f"{ctx['customer_order_id']}-{ctx['packaging_mmdd']}"),
        ("清關號碼",       lambda ctx: ""),  # 由 Nigel 端填入
        ("收件人",         lambda ctx: ctx["recipient"]),
        ("收件人詳細地址", lambda ctx: ctx["address"]),
        ("收件人電話號碼", lambda ctx: ctx["phone"]),
        ("申報人",         lambda ctx: ctx["recipient"]),         # 預設 = 收件人
        ("申報人詳細地址", lambda ctx: ctx["address"]),
        ("申報人電話號碼", lambda ctx: ctx["phone"]),
        ("品名",           lambda ctx: random_product_name(ctx["order_id"], ctx["package_id"])),
        ("數量",           lambda ctx: random_quantity(ctx["order_id"], ctx["package_id"])),
        ("金額",           lambda ctx: random_jpy_amount(ctx["order_id"], ctx["package_id"])),
        ("產地",           lambda ctx: random_origin(ctx["order_id"], ctx["package_id"])),
        ("URL/JanCode",    lambda ctx: ""),
    ],
}


JPD = {
    "id": "jpd",
    "display_name": "小客戶大價值 (JpD)",
    "filename_template": "{date}_JpD_出貨單.xlsx",
    "row_strategy": "one_per_package",
    "columns": [
        ("客戶運單號",      lambda ctx: f"{ctx['customer_order_id']}-{ctx['packaging_mmdd']}"),
        ("JpD包裹ID",       lambda ctx: ""),  # 由 JpD 端填入
        ("運單ID",          lambda ctx: ""),
        ("包裹特殊服務",     lambda ctx: ""),
        ("收件人",          lambda ctx: ctx["recipient"]),
        ("收件人身份證ID",   lambda ctx: ""),  # 不需要
        ("收件人詳細地址",   lambda ctx: ctx["address"]),
        ("收件人电话号码",   lambda ctx: ctx["phone"]),
        ("備註",            lambda ctx: ""),
        ("特殊服务",         lambda ctx: ""),
        ("渠道ID",          lambda ctx: 40),  # JpD 固定 40
        ("申報人",          lambda ctx: ctx["recipient"]),
        ("申報人身份證ID",   lambda ctx: ""),
        ("申報人詳細地址",   lambda ctx: ctx["address"]),
        ("申報人电话号码",   lambda ctx: ctx["phone"]),
        ("品名",            lambda ctx: random_product_name(ctx["order_id"], ctx["package_id"])),
        ("数量",            lambda ctx: random_quantity(ctx["order_id"], ctx["package_id"])),
        ("金额",            lambda ctx: random_jpy_amount(ctx["order_id"], ctx["package_id"])),
        ("材質",            lambda ctx: ""),
        ("產地",            lambda ctx: random_origin(ctx["order_id"], ctx["package_id"])),
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


def build_rows(vendor_id: str, orders: list[dict]) -> tuple[list[str], list[list]]:
    """
    依範本產生 (headers, rows)。
    orders 結構：每個 dict = 一筆 order_history + 該筆的 package_ids list。

    每個 package_id 一列。

    Returns:
      headers: List[str] — 第一列標頭
      rows:    List[List] — 資料列（每個 package 一列）
    """
    vendor = get_vendor(vendor_id)
    if not vendor:
        raise ValueError(f"unknown vendor: {vendor_id}")

    headers = [col[0] for col in vendor["columns"]]
    rows = []

    for o in orders:
        # packaging_mmdd = 建立 JPD 運單那一刻的 MMDD
        # 來源優先：created_at（本地存 order_history 時的時間） > 今天
        packaging_mmdd = ""
        for src in (o.get("created_at"),):
            if src and len(str(src)) >= 10:
                try:
                    # 接受 'YYYY-MM-DD HH:MM:SS' 或 'YYYY-MM-DD'
                    dt = datetime.strptime(str(src)[:10], "%Y-%m-%d")
                    packaging_mmdd = dt.strftime("%m%d")
                    break
                except ValueError:
                    pass
        if not packaging_mmdd:
            packaging_mmdd = datetime.now().strftime("%m%d")

        # customer_order_id 可能含 # 開頭、空白，要清乾淨
        coid = (o.get("customer_order_id") or "").strip().lstrip("#").strip()
        if not coid:
            coid = f"ORD{o.get('id', 0)}"  # fallback：本地流水號

        # 解析 package_ids（CSV 字串）
        pkg_ids = [
            int(x.strip()) for x in str(o.get("package_ids") or "").split(",")
            if x.strip().isdigit()
        ]
        if not pkg_ids:
            # 沒包裹 ID → 至少出 1 行（用 order.id 當 seed），避免空行
            pkg_ids = [o.get("id", 0)]

        for pid in pkg_ids:
            ctx = {
                "order_id":          o.get("id", 0),  # 本地 order_history.id（seed 用）
                "customer_order_id": coid,
                "packaging_mmdd":    packaging_mmdd,
                "recipient":         (o.get("recipient") or "").strip(),
                "address":           (o.get("address") or "").strip(),
                "phone":             (o.get("phone") or "").strip(),
                "logis_num":         (o.get("logis_num") or "").strip(),
                "shopify_order_id":  o.get("shopify_order_id") or "",
                "package_id":        pid,
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
