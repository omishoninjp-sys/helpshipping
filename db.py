# -*- coding: utf-8 -*-
"""
帳單來源：從 helpshipping 的 SQLite 撈出指定付款日區間的帳單。

已對照 app.py 實際 schema：
  資料表 shipment_requests
    g_code         客戶編號
    customer_name  客戶姓名
    total_fee      合計
    payment_last5  匯款欄（末五碼，或「現金 / 後付 / 管確認 / 收日幣現金」）
    payment_at     銷帳時間 ← 用來做付款日區間篩選
    updated_at     標記已出貨那天（＝出貨日），舊單 fallback created_at
    agent_id       0 = 主管理員
"""
from __future__ import annotations

import os
import sqlite3
from datetime import date, datetime

from .matcher import BillRecord

# 與 app.py 一致
DB_PATH = os.environ.get("DB_PATH", "packages.db")

# 自家收款帳號（偵測「末五碼誤填成自家帳號片段」）
OWN_ACCOUNT = os.environ.get("OWN_ACCOUNT", "699515361956")

BILL_SQL = """
SELECT
    sr.id                                                 AS bill_id,
    sr.g_code                                             AS customer_code,
    sr.customer_name                                      AS customer_name,
    COALESCE(NULLIF(sr.updated_at, ''), sr.created_at)    AS ship_date,
    sr.total_fee                                          AS amount,
    sr.payment_last5                                      AS pay_mark,
    sr.payment_at                                         AS paid_at
FROM shipment_requests sr
WHERE sr.status = '已出貨'
  AND COALESCE(sr.payment_last5, '') <> ''
  AND date(replace(substr(sr.payment_at, 1, 10), '/', '-'))
      BETWEEN date(:start) AND date(:end)
  {agent_filter}
ORDER BY sr.payment_at DESC, sr.id DESC
"""


def _to_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, date):
        return v
    s = str(v)[:19].replace("/", "-")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def connect(db_path: str | None = None) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path or DB_PATH, timeout=5)
    conn.row_factory = sqlite3.Row
    return conn


def fetch_bills(conn, start: date, end: date, agent_id: int = 0,
                sql: str | None = None) -> list[BillRecord]:
    params = {"start": start.isoformat(), "end": end.isoformat()}
    if sql is None:
        sql = BILL_SQL.format(
            agent_filter="AND sr.agent_id = :agent_id" if agent_id > 0 else "")
        if agent_id > 0:
            params["agent_id"] = agent_id

    out = []
    for r in conn.execute(sql, params).fetchall():
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
    out = {}
    for (name,) in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"):
        out[name] = [c[1] for c in conn.execute(f"PRAGMA table_info({name})")]
    return out
