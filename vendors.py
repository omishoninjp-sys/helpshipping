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
from datetime import datetime


# ============ 隨機資料產生器（用戶說「亂填即可」） ============

_PRODUCT_NAME_POOL = [
    "雑貨", "衣類", "玩具", "鞋子", "包包", "化妝品", "食品",
    "日用品", "餅乾", "髮夾", "上衣", "褲子", "外套", "手帕",
    "毛巾", "資料夾", "杯子", "公仔", "面膜", "卡片",
]
_ORIGIN_POOL = ["japan"] * 6 + ["china"] * 4  # 60% japan, 40% china


def _seed_for(shipment_id: int, package_id: int = 0) -> random.Random:
    """每筆出貨用固定 seed → 同一筆每次匯出結果一樣（避免亂跳）"""
    return random.Random(f"{shipment_id}-{package_id}")


def random_product_name(shipment_id: int, package_id: int = 0) -> str:
    return _seed_for(shipment_id, package_id).choice(_PRODUCT_NAME_POOL)


def random_quantity(shipment_id: int, package_id: int = 0) -> int:
    return _seed_for(shipment_id, package_id + 1).randint(1, 30)


def random_jpy_amount(shipment_id: int, package_id: int = 0) -> int:
    """單品 JPY 申報價：200~2000、四捨五入到百"""
    rng = _seed_for(shipment_id, package_id + 2)
    return rng.randint(2, 20) * 100


def random_origin(shipment_id: int, package_id: int = 0) -> str:
    return _seed_for(shipment_id, package_id + 3).choice(_ORIGIN_POOL)


# ============ 廠商範本 ============

NIGEL = {
    "id": "nigel",
    "display_name": "Nigel",
    "filename_template": "{date}_Nigel_出貨單.xlsx",
    "row_strategy": "one_per_package",
    "columns": [
        # (header, getter — 接收 ctx dict 回字串/數字)
        ("客戶編號",       lambda ctx: f"{ctx['g_code']}-{ctx['packaging_mmdd']}"),
        ("清關號碼",       lambda ctx: ""),  # 由 Nigel 端填入
        ("收件人",         lambda ctx: ctx["ship_recipient"]),
        ("收件人詳細地址", lambda ctx: ctx["ship_address"]),
        ("收件人電話號碼", lambda ctx: ctx["ship_phone"]),
        ("申報人",         lambda ctx: ctx["ship_recipient"]),         # 預設 = 收件人
        ("申報人詳細地址", lambda ctx: ctx["ship_address"]),
        ("申報人電話號碼", lambda ctx: ctx["ship_phone"]),
        ("品名",           lambda ctx: random_product_name(ctx["shipment_id"], ctx["package_id"])),
        ("數量",           lambda ctx: random_quantity(ctx["shipment_id"], ctx["package_id"])),
        ("金額",           lambda ctx: random_jpy_amount(ctx["shipment_id"], ctx["package_id"])),
        ("產地",           lambda ctx: random_origin(ctx["shipment_id"], ctx["package_id"])),
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
        ("JpD包裹ID",       lambda ctx: ""),  # 由 JpD 端填入
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
        ("品名",            lambda ctx: random_product_name(ctx["shipment_id"], ctx["package_id"])),
        ("数量",            lambda ctx: random_quantity(ctx["shipment_id"], ctx["package_id"])),
        ("金额",            lambda ctx: random_jpy_amount(ctx["shipment_id"], ctx["package_id"])),
        ("材質",            lambda ctx: ""),
        ("產地",            lambda ctx: random_origin(ctx["shipment_id"], ctx["package_id"])),
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
        packaging_mmdd = ""
        for src in (s.get("updated_at"), s.get("created_at")):
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

        for pkg in s["packages"]:
            ctx = {
                "shipment_id":          s["id"],
                "g_code":               s["g_code"],
                "packaging_mmdd":       packaging_mmdd,   # 客戶申請出單日（打包日）MMDD
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
