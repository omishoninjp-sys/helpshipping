# -*- coding: utf-8 -*-
"""
帳單來源：從 helpshipping 的 SQLite 撈出指定付款日區間的帳單。

⚠️ 需確認的三件事（見 README 的「接線」段落）：
   1. 帳單資料表名稱
   2. 欄位對應（BILL_SQL 內的 AS 別名要維持不變）
   3. 付款日欄位（目前假設是 paid_at，用來做日期區間篩選）
"""
from __future__ import annotations

import os
import sqlite3
from datetime import date, datetime

from .matcher import BillRecord

DB_PATH = os.environ.get("HELPSHIPPING_DB", "helpshipping.db")

# 自家收款帳號（用來偵測「末五碼誤填成自家帳號片段」）
OWN_ACCOUNT = os.environ.get("OWN_ACCOUNT", "699515361956")

# ---- 這段是唯一需要照實際 schema 改的地方 -------------------------------
BILL_SQL = """
SELECT
    b.id                AS bill_id,
    b.customer_code     AS customer_code,
    b.customer_name     AS customer_name,
    b.ship_date         AS ship_date,
    b.total             AS amount,
    b.pay_mark          AS pay_mark,
    b.paid_at           AS paid_at
FROM bills b
WHERE date(b.paid_at) BETWEEN date(:start) AND date(:end)
ORDER BY b.ship_date DESC
"""
# ------------------------------------------------------------------------


def _to_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, date):
        return v
    s = str(v)[:19].replace("/", "-")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%m-%d"):
        try:
            d = datetime.strptime(s, fmt)
            return d.replace(year=date.today().year).date() if fmt == "%m-%d" else d.date()
        except ValueError:
            continue
    return None


def connect(db_path: str | None = None) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path or DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def fetch_bills(conn, start: date, end: date, sql: str | None = None) -> list[BillRecord]:
    rows = conn.execute(sql or BILL_SQL,
                        {"start": start.isoformat(), "end": end.isoformat()}).fetchall()
    out = []
    for r in rows:
        out.append(BillRecord(
            bill_id=str(r["bill_id"]),
            customer_code=str(r["customer_code"] or ""),
            customer_name=str(r["customer_name"] or ""),
            ship_date=_to_date(r["ship_date"]),
            amount=int(round(float(r["amount"] or 0))),
            pay_mark=str(r["pay_mark"] or "").strip(),
            paid_at=_to_date(r["paid_at"]),
        ))
    return out


def inspect_schema(conn) -> dict:
    """列出所有資料表與欄位，方便確認 BILL_SQL 該怎麼寫。"""
    out = {}
    for (name,) in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"):
        out[name] = [c[1] for c in conn.execute(f"PRAGMA table_info({name})")]
    return out
