"""
MOPS 金融股月自結損益 Web 版 — PythonAnywhere 專用 WSGI / Flask 部署版 (flask_app.py)

特點：
1. 同時匯出 `app` 與 `application` (PEP 3333 WSGI callable)，直接相容 PythonAnywhere 的「Flask」或「Manual Configuration」模式。
2. 內建 SQLite3 資料庫持久化 (`mops_self_profit.db`)，隨附已預載之 25 檔金融股 11509 清單與 114/115 年度完整月自結損益資料。
3. 針對 PythonAnywhere 免費版 (Free Tier) 外部網路白名單限制進行容錯處理：
   - 若選擇「強制重抓」但受限於 PythonAnywhere 免費版 Proxy 無法連外時，若資料庫已有該筆資料則自動回退使用資料庫快取，避免畫面報錯。
   - 點擊 `/api/open_mops` 開啟原始報表時，若雲端伺服器連外受限，自動轉址至 MOPS 官方查詢頁由使用者瀏覽器直接開啟。
4. 亦支援在本機直接執行 `python flask_app.py` 啟動測試伺服器。
"""
import os
import re
import sys
import io
import json
import time
import sqlite3
import datetime
import threading
import urllib.parse
import urllib.request
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

import pandas as pd
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("DB_PATH", os.path.join(HERE, "mops_self_profit.db"))
DB_LOCK = threading.Lock()

FIN_MARK = "不適用"
API_URL = "https://mops.twse.com.tw/mops/api/redirectToOld"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
    "Content-Type": "application/json",
    "Referer": "https://mops.twse.com.tw/mops/",
    "Origin": "https://mops.twse.com.tw",
}
DELAY = 0.8


# ================================================================ SQLite3 資料庫層
def init_db():
    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS mops_list_meta (
                ym TEXT NOT NULL,
                market TEXT NOT NULL,
                total_companies INTEGER NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (ym, market)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS mops_list (
                ym TEXT NOT NULL,
                market TEXT NOT NULL,
                co_id TEXT NOT NULL,
                name TEXT NOT NULL,
                report_type TEXT NOT NULL,
                list_month_pretax REAL,
                list_cum_pretax REAL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (ym, market, co_id)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS company_year_meta (
                co_id TEXT NOT NULL,
                year INTEGER NOT NULL,
                name TEXT NOT NULL,
                months_count INTEGER NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (co_id, year)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS company_monthly_profit (
                co_id TEXT NOT NULL,
                year INTEGER NOT NULL,
                month INTEGER NOT NULL,
                month_pretax REAL,
                cum_pretax REAL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (co_id, year, month)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS financial_invest_metrics (
                co_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                price REAL,
                volume INTEGER,
                capital REAL,
                nav REAL,
                eps REAL,
                eps_quarter TEXT,
                dividend REAL,
                dividend_year TEXT,
                pe REAL,
                yield REAL,
                updated_at TEXT NOT NULL
            )
        """)
        conn.commit()


def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def db_get_list(ym, market):
    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        meta = conn.execute(
            "SELECT total_companies, updated_at FROM mops_list_meta WHERE ym = ? AND market = ?",
            (str(ym), str(market)),
        ).fetchone()
        if not meta:
            return None
        rows = conn.execute(
            """SELECT co_id, name, report_type, list_month_pretax, list_cum_pretax
               FROM mops_list WHERE ym = ? AND market = ? ORDER BY co_id ASC""",
            (str(ym), str(market)),
        ).fetchall()
        if not rows:
            return None
        stocks = []
        for r in rows:
            m_val = int(r["list_month_pretax"]) if r["list_month_pretax"] is not None else None
            c_val = int(r["list_cum_pretax"]) if r["list_cum_pretax"] is not None else None
            stocks.append({
                "co_id": r["co_id"],
                "name": r["name"],
                "label": f"{r['co_id']} {r['name']}",
                "report_type": r["report_type"],
                "list_month_pretax": m_val,
                "list_cum_pretax": c_val,
            })
        return {
            "ym": str(ym),
            "year": int(str(ym)[:-2]),
            "month": int(str(ym)[-2:]),
            "market": str(market),
            "total_companies": meta["total_companies"],
            "financial_count": len(stocks),
            "stocks": stocks,
            "source": "db",
            "updated_at": meta["updated_at"],
        }


def db_save_list(ym, market, total_companies, stocks):
    ts = now_str()
    ym, market = str(ym), str(market)
    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """INSERT OR REPLACE INTO mops_list_meta (ym, market, total_companies, updated_at)
               VALUES (?, ?, ?, ?)""",
            (ym, market, total_companies, ts),
        )
        conn.execute("DELETE FROM mops_list WHERE ym = ? AND market = ?", (ym, market))
        conn.executemany(
            """INSERT INTO mops_list (ym, market, co_id, name, report_type, list_month_pretax, list_cum_pretax, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            [
                (
                    ym,
                    market,
                    s["co_id"],
                    s["name"],
                    s["report_type"],
                    s["list_month_pretax"],
                    s["list_cum_pretax"],
                    ts,
                )
                for s in stocks
            ],
        )
        conn.commit()
    return ts


def db_get_company_year(co_id, year):
    co_id, year = str(co_id), int(year)
    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        meta = conn.execute(
            "SELECT name, months_count, updated_at FROM company_year_meta WHERE co_id = ? AND year = ?",
            (co_id, year),
        ).fetchone()
        if not meta or meta["months_count"] <= 0:
            return None
        rows = conn.execute(
            """SELECT month, month_pretax, cum_pretax
               FROM company_monthly_profit WHERE co_id = ? AND year = ? ORDER BY month ASC""",
            (co_id, year),
        ).fetchall()
        if not rows:
            return None
        month_map = {}
        for r in rows:
            month_map[int(r["month"])] = {
                "month_pretax": int(r["month_pretax"]) if r["month_pretax"] is not None else None,
                "cum_pretax": int(r["cum_pretax"]) if r["cum_pretax"] is not None else None,
            }
        return {
            "name": meta["name"],
            "month_map": month_map,
            "updated_at": meta["updated_at"],
            "source": "db",
        }


def db_save_company_year(co_id, year, name, month_map):
    co_id, year = str(co_id), int(year)
    ts = now_str()
    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """INSERT OR REPLACE INTO company_year_meta (co_id, year, name, months_count, updated_at)
               VALUES (?, ?, ?, ?, ?)""",
            (co_id, year, name or "", len(month_map), ts),
        )
        conn.execute("DELETE FROM company_monthly_profit WHERE co_id = ? AND year = ?", (co_id, year))
        conn.executemany(
            """INSERT INTO company_monthly_profit (co_id, year, month, month_pretax, cum_pretax, updated_at)
               VALUES (?, ?, ?, ?, ?, ?)""",
            [
                (co_id, year, int(m), d.get("month_pretax"), d.get("cum_pretax"), ts)
                for m, d in month_map.items()
            ],
        )
        conn.commit()
    return ts


# ================================================================ MOPS 爬蟲層
def mops_get(api_name, params, retries=2):
    body = json.dumps({"apiName": api_name, "parameters": params}).encode("utf-8")
    last_err = None
    for i in range(retries):
        try:
            req = urllib.request.Request(API_URL, data=body, headers=HEADERS, method="POST")
            with urllib.request.urlopen(req, timeout=15) as r:
                res = json.loads(r.read().decode("utf-8"))
            if res.get("code") != 200 or "result" not in res or "url" not in res["result"]:
                raise RuntimeError(f"MOPS API 回應非 200 或無資料: {res.get('message', res)}")
            url = res["result"]["url"]
            hreq = urllib.request.Request(
                url,
                headers={"User-Agent": HEADERS["User-Agent"], "Referer": HEADERS["Referer"]},
            )
            with urllib.request.urlopen(hreq, timeout=15) as r:
                html = r.read().decode("utf-8", errors="ignore")
            if "FOR SECURITY REASONS" in html:
                raise RuntimeError("MOPS 安全機制阻擋，稍後重試")
            time.sleep(DELAY)
            return html
        except Exception as e:
            last_err = e
            time.sleep(DELAY * (i + 1))
    raise RuntimeError(f"{api_name} 請求失敗 (若部署於 PythonAnywhere 免費版，可能受限於外部連線白名單，建議切換為「🗄️ 資料庫優先」): {last_err}")


def to_num(s):
    if s is None:
        return None
    s = str(s).strip().replace(",", "")
    if s == "" or FIN_MARK in s:
        return None
    try:
        return int(s)
    except ValueError:
        try:
            return float(s)
        except ValueError:
            return None


def norm_label(s):
    s = s.strip()
    for t in ["（仟元）：", "（最多60個中文字）：", "：", "合併"]:
        s = s.replace(t, "")
    return s


def calc_yoy(cur, prev):
    if cur is None or prev is None or prev == 0:
        return None
    return (cur - prev) / abs(prev)


def fetch_list_data(ym, market="1", force=False):
    ym, market = str(ym), str(market)
    if not force:
        cached = db_get_list(ym, market)
        if cached:
            return cached

    try:
        html = mops_get("ajax_t138sb01", {
            "CK2": "1",
            "YM1": ym,
            "CK1": market,
            "encodeURIComponent": 1,
            "firstin": 1,
            "step": 1,
            "first": 1,
        })
    except Exception:
        # 若雲端環境連外受阻但資料庫已有該年月快取，優雅退回資料庫資料
        cached = db_get_list(ym, market)
        if cached:
            return cached
        raise

    soup = BeautifulSoup(html, "html.parser")
    all_rows = []
    for tbl in soup.find_all("table"):
        ths = tbl.find_all("th")
        if not ths or "公司代號" not in ths[0].text:
            continue
        prev = tbl.find_previous("table", class_="noBorder")
        rtype = "合併" if prev and "合併" in prev.text else "個別"
        headers = []
        for th in ths:
            h = norm_label(th.text)
            if h not in headers:
                headers.append(h)
        for tr in tbl.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) != len(headers):
                continue
            rec = {h: td.text.strip().replace("\xa0", " ") for h, td in zip(headers, tds)}
            rec["報表類型"] = rtype
            all_rows.append(rec)

    fin_stocks = []
    for r in all_rows:
        op_val = r.get("本月營業損益", "")
        if FIN_MARK in op_val:
            fin_stocks.append({
                "co_id": r.get("公司代號", ""),
                "name": r.get("公司簡稱", ""),
                "label": f"{r.get('公司代號', '')} {r.get('公司簡稱', '')}",
                "report_type": r.get("報表類型", ""),
                "list_month_pretax": to_num(r.get("本月稅前損益")),
                "list_cum_pretax": to_num(r.get("本年累計稅前損益")),
            })

    fin_stocks.sort(key=lambda x: x["co_id"])
    ts = db_save_list(ym, market, len(all_rows), fin_stocks) if fin_stocks else now_str()
    return {
        "ym": ym,
        "year": int(ym[:-2]),
        "month": int(ym[-2:]),
        "market": market,
        "total_companies": len(all_rows),
        "financial_count": len(fin_stocks),
        "stocks": fin_stocks,
        "source": "mops",
        "updated_at": ts,
    }


def find_latest_ym(market="1"):
    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT ym FROM mops_list_meta WHERE market = ? ORDER BY ym DESC LIMIT 1",
            (str(market),),
        ).fetchone()
    if row:
        cached = db_get_list(row[0], market)
        if cached and cached.get("stocks"):
            return cached

    today = datetime.date.today()
    y, m = today.year - 1911, today.month
    for _ in range(6):
        m -= 1
        if m == 0:
            y, m = y - 1, 12
        ym = f"{y}{m:02d}"
        try:
            res = fetch_list_data(ym, market, force=False)
            if res["total_companies"] > 0:
                return res
        except Exception:
            pass
    raise RuntimeError("找不到最近 6 個月的月自結損益清單")


def fetch_company_year_data(co_id, year, force=False):
    co_id, year = str(co_id), int(year)
    if not force:
        db_res = db_get_company_year(co_id, year)
        if db_res:
            return db_res["name"], db_res["month_map"], "db", db_res["updated_at"]

    try:
        q1 = mops_get("ajax_t138sb02_q1", {
            "co_id": co_id,
            "year": str(year),
            "encodeURIComponent": 1,
            "step": 1,
            "firstin": 1,
            "off": 1,
        })
        hidden = dict(re.findall(r"name='(\w+)' value='([^']*)'", q1))
        params = {"encodeURIComponent": 1, "step": 1, "firstin": 1, "off": 1}
        for k in ["CK2", "BK1", "TYPEK", "YEAR", "COMP"]:
            if k in hidden:
                params[k] = hidden[k]
        if "COMP" not in params:
            params["COMP"] = co_id
            params["YEAR"] = str(year)
            params["CK2"] = "1"
            params["BK1"] = "2"
            params["TYPEK"] = "sii"

        html = mops_get("ajax_t138sb02", params)
    except Exception:
        # 若雲端連外受阻但資料庫已有該年度資料，自動回退使用資料庫快取
        db_res = db_get_company_year(co_id, year)
        if db_res:
            return db_res["name"], db_res["month_map"], "db", db_res["updated_at"]
        return "", {}, "mops", now_str()

    soup = BeautifulSoup(html, "html.parser")
    name = ""
    for h in soup.find_all("h3"):
        mm = re.search(r"公司名稱：\s*(\S+)", h.text)
        if mm:
            name = mm.group(1).strip()

    caps = [t for t in soup.find_all("table", class_="noBorder") if re.search(rf"{year}年\s*\d+月", t.text)]
    datas = soup.find_all("table", class_="hasBorder")
    month_map = {}
    for cap, data in zip(caps, datas):
        m_match = re.search(rf"{year}年\s*(\d+)月", cap.text)
        if not m_match:
            continue
        month = int(m_match.group(1))
        rec = {}
        for tr in data.find_all("tr"):
            th, td = tr.find("th"), tr.find("td")
            if th and td:
                rec[norm_label(th.text)] = td.text.strip()
        month_map[month] = {
            "month_pretax": to_num(rec.get("本月稅前損益")),
            "cum_pretax": to_num(rec.get("本年累計稅前損益")),
        }

    ts = db_save_company_year(co_id, year, name, month_map) if month_map else now_str()
    return name, month_map, "mops", ts


def build_company_comparison(co_id, y1, y2, ref_month=None, force=False):
    y1, y2 = int(y1), int(y2)
    name1, map1, src1, ts1 = fetch_company_year_data(co_id, y1, force=force)
    name2, map2, src2, ts2 = fetch_company_year_data(co_id, y2, force=force)
    name = name2 or name1 or ""

    rows = []
    latest_m_in_y2 = max(map2.keys()) if map2 else None
    target_m = int(ref_month) if ref_month and int(ref_month) in map2 else (int(ref_month) if ref_month else latest_m_in_y2)

    for m in range(1, 13):
        d1 = map1.get(m, {})
        d2 = map2.get(m, {})
        y1_mon = d1.get("month_pretax")
        y1_cum = d1.get("cum_pretax")
        y2_mon = d2.get("month_pretax")
        y2_cum = d2.get("cum_pretax")
        rows.append({
            "month": m,
            "month_label": f"{m}月",
            "y1_month_pretax": y1_mon,
            "y1_cum_pretax": y1_cum,
            "y2_month_pretax": y2_mon,
            "y2_cum_pretax": y2_cum,
            "month_yoy": calc_yoy(y2_mon, y1_mon),
            "cum_yoy": calc_yoy(y2_cum, y1_cum),
        })

    ref_row = rows[target_m - 1] if (target_m and 1 <= target_m <= 12) else None
    overall_source = "db" if (src1 == "db" and src2 == "db") else "mops"
    return {
        "co_id": str(co_id),
        "name": name,
        "label": f"{co_id} {name}".strip(),
        "y1": y1,
        "y2": y2,
        "ref_month": target_m,
        "latest_row": ref_row,
        "rows": rows,
        "source": overall_source,
        "updated_at": max(ts1, ts2),
    }


def style_worksheet(ws, pct_cols=()):
    hdr_fill = PatternFill("solid", start_color="1F497D")
    hdr_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    font = Font(name="微軟正黑體", size=10)
    border = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
    headers = [c.value for c in ws[1]]
    for c in ws[1]:
        c.fill, c.font = hdr_fill, hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 36
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font, c.border = font, border
            h = headers[c.column - 1]
            if h in pct_cols:
                c.number_format = "0.00%;[Red]-0.00%"
                c.alignment = Alignment(horizontal="right")
            elif isinstance(c.value, (int, float)):
                c.number_format = "#,##0;[Red]-#,##0"
                c.alignment = Alignment(horizontal="right")
            else:
                c.alignment = Alignment(horizontal="center" if c.column in (1, 2) else "left")
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[letter].width = min(max(width * 1.6, 13), 40)
    ws.freeze_panes = "B2"


def generate_excel_bytes(payload):
    ym = payload.get("ym", "11509")
    y1 = payload.get("y1", 114)
    y2 = payload.get("y2", 115)
    stocks = payload.get("stocks", [])
    company_details = payload.get("companies", {})

    cur_y, cur_m = int(ym[:-2]), int(ym[-2:])
    ym_lbl = f"{cur_y}/{cur_m:02d}"

    col_y1_m = f"{y1}年度 本月稅前損益"
    col_y1_c = f"{y1}年度 本年累計稅前損益"
    col_y2_m = f"{y2}年度 本月稅前損益"
    col_y2_c = f"{y2}年度 本年累計稅前損益"

    summary_rows = []
    for s in stocks:
        co_id = s["co_id"]
        comp = company_details.get(co_id, {})
        latest = comp.get("latest_row") or {}
        summary_rows.append({
            "金融股": s.get("label", f"{co_id} {s.get('name', '')}"),
            "報表類型": s.get("report_type", ""),
            f"{ym_lbl} 本月稅前損益(清單)": s.get("list_month_pretax"),
            f"{ym_lbl} 本年累計稅前損益(清單)": s.get("list_cum_pretax"),
            col_y1_m: latest.get("y1_month_pretax"),
            col_y1_c: latest.get("y1_cum_pretax"),
            col_y2_m: latest.get("y2_month_pretax"),
            col_y2_c: latest.get("y2_cum_pretax"),
            "累計稅前 YoY": latest.get("cum_yoy"),
        })

    df_summary = pd.DataFrame(summary_rows)
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="自結清單", index=False)
        for co_id, comp in company_details.items():
            rows = comp.get("rows", [])
            if not rows:
                continue
            sheet_rows = []
            for r in rows:
                sheet_rows.append({
                    "月份": r["month_label"],
                    col_y1_m: r["y1_month_pretax"],
                    col_y1_c: r["y1_cum_pretax"],
                    col_y2_m: r["y2_month_pretax"],
                    col_y2_c: r["y2_cum_pretax"],
                    "單月稅前 YoY": r["month_yoy"],
                    "累計稅前 YoY": r["cum_yoy"],
                })
            sname = comp.get("label", co_id)[:31]
            pd.DataFrame(sheet_rows).to_excel(writer, sheet_name=sname, index=False)

    bio.seek(0)
    wb = openpyxl.load_workbook(bio)
    for ws in wb.worksheets:
        style_worksheet(ws, pct_cols=("累計稅前 YoY", "單月稅前 YoY"))
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def enrich_invest_items(items, ym, self_profit_map):
    month_num = int(str(ym)[-2:])
    for item in items:
        c = item["co_id"]
        sp = self_profit_map.get(c)
        cap = item.get("capital")
        if sp and cap and cap > 0 and (sp.get("cum") is not None or sp.get("mon") is not None):
            cum_profit = sp.get("cum")
            mon_profit = sp.get("mon")
            item["has_self_profit"] = True
            item["self_profit_month"] = month_num
            item["self_profit_ym"] = str(ym)
            item["self_profit_cum"] = cum_profit
            item["self_profit_mon"] = mon_profit
            item["eps_self_cum"] = round(cum_profit / (cap * 10000.0), 2) if cum_profit is not None else None
            item["eps_self_mon"] = round(mon_profit / (cap * 10000.0), 2) if mon_profit is not None else None
        else:
            item["has_self_profit"] = False
            item["self_profit_month"] = month_num
            item["self_profit_ym"] = str(ym)
            item["self_profit_cum"] = None
            item["self_profit_mon"] = None
            item["eps_self_cum"] = None
            item["eps_self_mon"] = None


def fetch_invest_metrics(force=False, ym=None):
    ts = now_str()
    if not ym:
        with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
            row = conn.execute("SELECT ym FROM mops_list_meta ORDER BY ym DESC LIMIT 1").fetchone()
            ym = row[0] if row else "11509"
    month_num = int(str(ym)[-2:])

    self_profit_map = {}
    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        sp_rows = conn.execute(
            "SELECT co_id, list_month_pretax, list_cum_pretax FROM mops_list WHERE ym = ?",
            (str(ym),)
        ).fetchall()
        for r in sp_rows:
            self_profit_map[r["co_id"]] = {
                "mon": r["list_month_pretax"],
                "cum": r["list_cum_pretax"]
            }

    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cached = conn.execute("""
            SELECT co_id, name, price, volume, capital, nav, eps, eps_quarter, dividend, dividend_year, pe, yield, updated_at
            FROM financial_invest_metrics ORDER BY co_id ASC
        """).fetchall()
        if cached and len(cached) >= 20 and not force:
            items = [dict(r) for r in cached]
            enrich_invest_items(items, ym, self_profit_map)
            return {
                "items": items,
                "source": "db",
                "self_profit_month": month_num,
                "self_profit_ym": str(ym),
                "updated_at": cached[0]["updated_at"],
                "count": len(items)
            }

    codes = []
    names = {}
    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute("SELECT co_id, name FROM mops_list GROUP BY co_id ORDER BY co_id ASC").fetchall()
        for r in rows:
            codes.append(r[0])
            names[r[0]] = r[1]

    if not codes:
        codes = [
            "2801", "2812", "2834", "2836", "2838", "2845", "2849", "2855",
            "2880", "2881", "2882", "2883", "2884", "2885", "2886", "2887",
            "2889", "2890", "2891", "2892", "2897", "5876", "5880", "6005", "6024"
        ]

    # A. BWIBBU_d (股價, 殖利率, 本益比, 股價淨值比)
    bw_map = {}
    try:
        req = urllib.request.Request("https://www.twse.com.tw/rwd/zh/afterTrading/BWIBBU_d?response=json", headers=HEADERS)
        with urllib.request.urlopen(req, timeout=12) as r:
            bw_json = json.loads(r.read().decode("utf-8"))
        for row in bw_json.get("data", []):
            if row[0] in codes:
                bw_map[row[0]] = {
                    "price": to_num(row[2]),
                    "yield": to_num(row[3]),
                    "dividend_year": row[4],
                    "pe": to_num(row[5]),
                    "pb": to_num(row[6]),
                    "name": row[1]
                }
    except Exception as e:
        print("[BWIBBU_d fetch error]", e)

    # B. STOCK_DAY_ALL (成交股數)
    day_map = {}
    try:
        req = urllib.request.Request("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL", headers=HEADERS)
        with urllib.request.urlopen(req, timeout=12) as r:
            day_json = json.loads(r.read().decode("utf-8"))
        for item in day_json:
            c = item.get("Code")
            if c in codes:
                day_map[c] = {
                    "volume": int(item.get("TradeVolume", 0) or 0),
                    "closing_price": to_num(item.get("ClosingPrice")),
                    "name": item.get("Name")
                }
    except Exception as e:
        print("[STOCK_DAY_ALL fetch error]", e)

    # C. Dividend t187ap45_L
    div_map = {}
    try:
        req = urllib.request.Request("https://openapi.twse.com.tw/v1/opendata/t187ap45_L", headers=HEADERS)
        with urllib.request.urlopen(req, timeout=12) as r:
            div_json = json.loads(r.read().decode("utf-8"))
        for item in div_json:
            c = item.get("公司代號")
            if c in codes:
                c1 = float(item.get("股東配發-盈餘分配之現金股利(元/股)", 0) or 0)
                c2 = float(item.get("股東配發-法定盈餘公積發放之現金(元/股)", 0) or 0)
                c3 = float(item.get("股東配發-資本公積發放之現金(元/股)", 0) or 0)
                tot = c1 + c2 + c3
                y = item.get("股利年度", "")
                if c not in div_map or str(y) > str(div_map[c].get("year", "")):
                    div_map[c] = {"cash": tot, "year": str(y)}
    except Exception as e:
        print("[Dividend fetch error]", e)

    # D. Balance sheets for NAV & Capital
    nav_map = {}
    for ep in ["t187ap07_L_fh", "t187ap07_L_basi", "t187ap07_L_bd"]:
        try:
            req = urllib.request.Request(f"https://openapi.twse.com.tw/v1/opendata/{ep}", headers=HEADERS)
            with urllib.request.urlopen(req, timeout=12) as r:
                bs_json = json.loads(r.read().decode("utf-8"))
            for item in bs_json:
                c = item.get("公司代號")
                if c in codes:
                    nav_map[c] = {
                        "nav": to_num(item.get("每股參考淨值")),
                        "capital_thousand": to_num(item.get("股本")),
                        "quarter": f"{item.get('年度')}/Q{item.get('季別')}"
                    }
        except Exception as e:
            print(f"[BS {ep} error]", e)

    # D2. Common Stock Capital from t187ap03_L (排除特別股，改採已發行普通股股本)
    common_cap_map = {}
    try:
        req = urllib.request.Request("https://openapi.twse.com.tw/v1/opendata/t187ap03_L", headers=HEADERS)
        with urllib.request.urlopen(req, timeout=12) as r:
            co_json = json.loads(r.read().decode("utf-8"))
        for item in co_json:
            c = item.get("公司代號")
            if c in codes:
                shs = item.get("已發行普通股數或TDR原股發行股數")
                if shs:
                    try:
                        # 股數 * 10 / 100,000,000 = 股數 / 10,000,000 (換算為億元)
                        common_cap_map[c] = round(float(shs) / 10000000.0, 2)
                    except Exception:
                        pass
    except Exception as e:
        print("[Company profile t187ap03_L fetch error]", e)

    # E. Income statement for EPS (t187ap14_L)
    eps_map = {}
    try:
        req = urllib.request.Request("https://openapi.twse.com.tw/v1/opendata/t187ap14_L", headers=HEADERS)
        with urllib.request.urlopen(req, timeout=12) as r:
            is_json = json.loads(r.read().decode("utf-8"))
        for item in is_json:
            c = item.get("公司代號")
            if c in codes:
                eps_map[c] = {
                    "eps": to_num(item.get("基本每股盈餘(元)")),
                    "quarter": f"{item.get('年度')}/Q{item.get('季別')}"
                }
    except Exception as e:
        print("[EPS fetch error]", e)

    items = []
    db_records = []
    for c in codes:
        b = bw_map.get(c, {})
        d = day_map.get(c, {})
        dv = div_map.get(c, {})
        nv = nav_map.get(c, {})
        ep = eps_map.get(c, {})

        name = names.get(c) or b.get("name") or d.get("name") or c
        price = b.get("price") if b.get("price") is not None else d.get("closing_price")
        volume = d.get("volume", 0)

        # 優先採用排除特別股之「普通股股本」
        capital = common_cap_map.get(c)
        if capital is None and nv.get("capital_thousand") is not None:
            capital = round(nv["capital_thousand"] / 100000.0, 2)

        nav = nv.get("nav")
        if nav is None and price is not None and b.get("pb"):
            try:
                nav = round(price / b["pb"], 2)
            except Exception:
                pass

        eps = ep.get("eps")
        eps_q = ep.get("quarter", "")

        dividend = dv.get("cash")
        dividend_y = dv.get("year", "")
        if dividend is None and price is not None and b.get("yield"):
            try:
                dividend = round(price * (b["yield"] / 100.0), 2)
                dividend_y = str(b.get("dividend_year", ""))
            except Exception:
                pass

        pe = b.get("pe")
        yield_val = b.get("yield")

        row_dict = {
            "co_id": c,
            "name": name,
            "price": price,
            "volume": volume,
            "capital": capital,
            "nav": nav,
            "eps": eps,
            "eps_quarter": eps_q,
            "dividend": dividend,
            "dividend_year": dividend_y,
            "pe": pe,
            "yield": yield_val,
            "updated_at": ts
        }
        items.append(row_dict)
        db_records.append((
            c, name, price, volume, capital, nav, eps, eps_q, dividend, dividend_y, pe, yield_val, ts
        ))

    with DB_LOCK, sqlite3.connect(DB_PATH) as conn:
        conn.executemany("""
            INSERT OR REPLACE INTO financial_invest_metrics (
                co_id, name, price, volume, capital, nav, eps, eps_quarter, dividend, dividend_year, pe, yield, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, db_records)
        conn.commit()

    enrich_invest_items(items, ym, self_profit_map)
    return {
        "items": items,
        "source": "twse",
        "self_profit_month": month_num,
        "self_profit_ym": str(ym),
        "updated_at": ts,
        "count": len(items)
    }


def generate_invest_excel_bytes(items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "金融股投資表"

    sample_m = 9
    for it in items:
        if it.get("self_profit_month"):
            sample_m = it["self_profit_month"]
            break

    headers = [
        "代號", "名稱", "股價", "成交量(張)", "普通股本(億元)", "每股淨值",
        f"EPS({sample_m}月自結)", "EPS", "EPS季別", "配息(元)", "股利年度", "本益比", "殖利率"
    ]
    ws.append(headers)

    for item in items:
        vol = item.get("volume", 0) // 1000 if item.get("volume") else None
        yd = (item.get("yield") / 100.0) if item.get("yield") is not None else None
        eps_self = item.get("eps_self_cum") if item.get("has_self_profit") else "-"
        ws.append([
            item.get("co_id"),
            item.get("name"),
            item.get("price"),
            vol,
            item.get("capital"),
            item.get("nav"),
            eps_self,
            item.get("eps"),
            item.get("eps_quarter"),
            item.get("dividend"),
            item.get("dividend_year"),
            item.get("pe"),
            yd,
        ])

    hdr_fill = PatternFill("solid", fgColor="1E3A8A")
    hdr_font = Font(name="Microsoft JhengHei", size=11, bold=True, color="FFFFFF")
    border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for cell in ws[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[1].height = 28

    for row in ws.iter_rows(min_row=2):
        ws.row_dimensions[row[0].row].height = 20
        for cell in row:
            cell.font = Font(name="Microsoft JhengHei", size=10)
            cell.border = border
            col_idx = cell.column
            if col_idx in (1, 2, 9, 11):
                cell.alignment = Alignment(horizontal="center")
            elif col_idx == 13:
                cell.number_format = "0.00%"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx in (3, 5, 6, 7, 8, 10, 12):
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 4:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right")

    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[letter].width = min(max(width * 1.6, 12), 25)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def generate_company_excel_bytes(comp):
    wb = openpyxl.Workbook()
    ws = wb.active
    label = comp.get("label", comp.get("co_id", "個股自結"))
    ws.title = str(label)[:31]
    y1 = comp.get("y1", 114)
    y2 = comp.get("y2", 115)
    headers = [
        "月份",
        f"{y1}年度 本月稅前損益",
        f"{y1}年度 本年累計稅前損益",
        f"{y2}年度 本月稅前損益",
        f"{y2}年度 本年累計稅前損益",
        "單月稅前 YoY",
        "累計稅前 YoY"
    ]
    ws.append(headers)
    for r in comp.get("rows", []):
        ws.append([
            r.get("month_label", f"{r.get('month')}月"),
            r.get("y1_month_pretax"),
            r.get("y1_cum_pretax"),
            r.get("y2_month_pretax"),
            r.get("y2_cum_pretax"),
            r.get("month_yoy"),
            r.get("cum_yoy")
        ])
    style_worksheet(ws, pct_cols=("單月稅前 YoY", "累計稅前 YoY"))
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def generate_ezsearch_excel_bytes(items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "公告快易查"
    headers = [
        "公告日期", "公告時間", "市場產業別", "產業別代碼", "公司代號", "公司簡稱", "公告項目代碼", "公告項目", "公告主旨", "詳細連結"
    ]
    ws.append(headers)
    for it in items:
        ws.append([
            it.get("CDATE", ""),
            it.get("CTIME", ""),
            it.get("TYPEK", ""),
            it.get("CODE_NAME", ""),
            it.get("COMPANY_ID", ""),
            it.get("COMPANY_NAME", ""),
            it.get("AN_CODE", ""),
            it.get("AN_NAME", ""),
            it.get("SUBJECT", ""),
            it.get("HYPERLINK", "")
        ])
    style_worksheet(ws)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def fetch_ezsearch_announcements(sdate="20261001", edate="20261023", subject="自結", co_id="", custom_url=""):
    if custom_url and custom_url.strip():
        url = custom_url.strip()
        try:
            parsed = urllib.parse.urlparse(url)
            params = urllib.parse.parse_qs(parsed.query)
            if "SUBJECT" in params and params["SUBJECT"]:
                subject = params["SUBJECT"][0]
            if "SDATE" in params and params["SDATE"]:
                sdate = params["SDATE"][0]
            if "EDATE" in params and params["EDATE"]:
                edate = params["EDATE"][0]
            if "CO_ID" in params and params["CO_ID"]:
                co_id = params["CO_ID"][0]
        except Exception:
            pass
    else:
        sdate_clean = str(sdate).strip().replace("-", "")
        edate_clean = str(edate).strip().replace("-", "")
        encoded_subj = urllib.parse.quote(subject)
        encoded_sdate = urllib.parse.quote(sdate_clean)
        encoded_edate = urllib.parse.quote(edate_clean)
        encoded_coid = urllib.parse.quote(co_id)
        url = f"https://mopsov.twse.com.tw/mops/web/ezsearch_query?step=00&RADIO_CM=1&TYPEK=sii&CO_MARKET=17&CO_ID={encoded_coid}&PRO_ITEM=M99&SUBJECT={encoded_subj}&SDATE={encoded_sdate}&EDATE={encoded_edate}&lang=TW&AN="

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://mopsov.twse.com.tw/mops/web/ezsearch"
    }
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=15) as r:
        raw_text = r.read().decode("utf-8", errors="ignore")

    clean_text = raw_text.strip().lstrip("\ufeff").strip()
    data_json = json.loads(clean_text)
    items = data_json.get("data", [])
    return {
        "items": items,
        "count": len(items),
        "url": url,
        "sdate": sdate,
        "edate": edate,
        "subject": subject,
        "updated_at": now_str()
    }


# ================================================================ PEP 3333 WSGI Application (相容 PythonAnywhere)
init_db()


def _json_response(start_response, data, status_str="200 OK"):
    raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
    headers = [
        ("Content-Type", "application/json; charset=utf-8"),
        ("Content-Length", str(len(raw))),
        ("Cache-Control", "no-store"),
    ]
    start_response(status_str, headers)
    return [raw]


def application(environ, start_response):
    method = environ.get("REQUEST_METHOD", "GET").upper()
    path = environ.get("PATH_INFO", "/")
    query_string = environ.get("QUERY_STRING", "")
    qs = urllib.parse.parse_qs(query_string)

    if method == "GET":
        if path.startswith("/assets/"):
            filename = os.path.basename(path)
            asset_path = os.path.join(HERE, "assets", filename)
            if os.path.exists(asset_path):
                ext = os.path.splitext(filename)[1].lower()
                mime = "image/png" if ext == ".png" else "application/octet-stream"
                with open(asset_path, "rb") as f:
                    content = f.read()
                start_response("200 OK", [
                    ("Content-Type", mime),
                    ("Content-Length", str(len(content))),
                    ("Cache-Control", "public, max-age=86400"),
                ])
                return [content]
            else:
                return _json_response(start_response, {"ok": False, "error": "Asset not found"}, "404 Not Found")

        if path == "/api/ezsearch":
            sdate = qs.get("sdate", ["20261001"])[0].strip()
            edate = qs.get("edate", ["20261023"])[0].strip()
            subject = qs.get("subject", ["自結"])[0].strip()
            co_id = qs.get("co_id", [""])[0].strip()
            custom_url = qs.get("url", [""])[0].strip()
            try:
                data = fetch_ezsearch_announcements(sdate=sdate, edate=edate, subject=subject, co_id=co_id, custom_url=custom_url)
                return _json_response(start_response, {"ok": True, **data})
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

        if path in ("/", "/index.html"):
            index_path = os.path.join(HERE, "index.html")
            with open(index_path, "rb") as f:
                content = f.read()
            start_response("200 OK", [
                ("Content-Type", "text/html; charset=utf-8"),
                ("Content-Length", str(len(content))),
            ])
            return [content]

        if path == "/api/init":
            try:
                today = datetime.date.today()
                cur_roc_year = today.year - 1911
                latest_data = find_latest_ym("1")
                return _json_response(start_response, {
                    "ok": True,
                    "default_y1": cur_roc_year - 1,
                    "default_y2": cur_roc_year,
                    "latest_ym": latest_data["ym"],
                    "list_data": latest_data,
                })
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

        if path == "/api/list":
            ym = qs.get("ym", ["11509"])[0].strip()
            market = qs.get("market", ["1"])[0].strip()
            force = qs.get("force", ["0"])[0] == "1"
            try:
                data = fetch_list_data(ym, market, force=force)
                return _json_response(start_response, {"ok": True, "list_data": data})
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

        if path == "/api/company":
            co_id = qs.get("co_id", [""])[0].strip()
            y1 = int(qs.get("y1", ["114"])[0])
            y2 = int(qs.get("y2", ["115"])[0])
            ref_m = qs.get("month", [None])[0]
            force = qs.get("force", ["0"])[0] == "1"
            if not co_id:
                return _json_response(start_response, {"ok": False, "error": "缺少 co_id 參數"}, "400 Bad Request")
            try:
                comp = build_company_comparison(co_id, y1, y2, ref_month=ref_m, force=force)
                return _json_response(start_response, {"ok": True, "company": comp})
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

        if path == "/api/invest_table":
            force = qs.get("force", ["0"])[0] == "1"
            ym = qs.get("ym", [None])[0]
            try:
                data = fetch_invest_metrics(force=force, ym=ym)
                return _json_response(start_response, {"ok": True, **data})
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

        if path == "/api/open_mops":
            qtype = qs.get("type", ["list"])[0].strip()
            fallback_url = "https://mops.twse.com.tw/mops/#/web/t138sb01" if qtype == "list" else "https://mops.twse.com.tw/mops/#/web/t138sb02_q1"
            try:
                if qtype == "list":
                    ym = qs.get("ym", ["11509"])[0].strip()
                    market = qs.get("market", ["1"])[0].strip()
                    body = {
                        "apiName": "ajax_t138sb01",
                        "parameters": {
                            "CK2": "1",
                            "YM1": ym,
                            "CK1": market,
                            "encodeURIComponent": 1,
                            "firstin": 1,
                            "step": 1,
                            "first": 1,
                        },
                    }
                elif qtype == "company_q1":
                    co_id = qs.get("co_id", ["2881"])[0].strip()
                    year = qs.get("year", ["115"])[0].strip()
                    body = {
                        "apiName": "ajax_t138sb02_q1",
                        "parameters": {
                            "co_id": co_id,
                            "year": year,
                            "encodeURIComponent": 1,
                            "step": 1,
                            "firstin": 1,
                            "off": 1,
                        },
                    }
                else:
                    co_id = qs.get("co_id", ["2881"])[0].strip()
                    year = qs.get("year", ["115"])[0].strip()
                    body = {
                        "apiName": "ajax_t138sb02",
                        "parameters": {
                            "encodeURIComponent": 1,
                            "step": 1,
                            "firstin": 1,
                            "off": 1,
                            "CK2": "1",
                            "BK1": "2",
                            "TYPEK": "sii",
                            "YEAR": year,
                            "COMP": co_id,
                        },
                    }
                req = urllib.request.Request(
                    API_URL,
                    data=json.dumps(body).encode("utf-8"),
                    headers=HEADERS,
                    method="POST",
                )
                with urllib.request.urlopen(req, timeout=8) as r:
                    res = json.loads(r.read().decode("utf-8"))
                target_url = res["result"]["url"]
            except Exception:
                # 若 PythonAnywhere 免費版擋外連，自動轉址到 MOPS 查詢頁
                target_url = fallback_url

            start_response("302 Found", [("Location", target_url)])
            return [b""]

    elif method == "POST":
        if path == "/api/export_excel":
            try:
                length = int(environ.get("CONTENT_LENGTH") or 0)
                raw_body = environ["wsgi.input"].read(length).decode("utf-8")
                body = json.loads(raw_body)
                xlsx_bytes = generate_excel_bytes(body)
                ym = body.get("ym", "11509")
                y1 = body.get("y1", 114)
                y2 = body.get("y2", 115)
                fname = f"MOPS_financial_self_profit_{ym}_{y1}vs{y2}.xlsx"
                start_response("200 OK", [
                    ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    ("Content-Disposition", f'attachment; filename="{fname}"'),
                    ("Content-Length", str(len(xlsx_bytes))),
                ])
                return [xlsx_bytes]
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

        if path == "/api/export_invest_excel":
            try:
                length = int(environ.get("CONTENT_LENGTH") or 0)
                raw_body = environ["wsgi.input"].read(length).decode("utf-8")
                body = json.loads(raw_body)
                items = body.get("items", [])
                xlsx_bytes = generate_invest_excel_bytes(items)
                today_str = datetime.date.today().strftime("%Y%m%d")
                fname = f"financial_stocks_invest_table_{today_str}.xlsx"
                start_response("200 OK", [
                    ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    ("Content-Disposition", f'attachment; filename="{fname}"'),
                    ("Content-Length", str(len(xlsx_bytes))),
                ])
                return [xlsx_bytes]
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

        if path == "/api/export_company_excel":
            try:
                length = int(environ.get("CONTENT_LENGTH") or 0)
                raw_body = environ["wsgi.input"].read(length).decode("utf-8")
                body = json.loads(raw_body)
                comp = body.get("comp", {})
                xlsx_bytes = generate_company_excel_bytes(comp)
                co_id = comp.get("co_id", "company")
                name = comp.get("name", "")
                ascii_fname = f"{co_id}_self_profit.xlsx"
                utf8_fname = urllib.parse.quote(f"{co_id}_{name}_自結損益_{comp.get('y1',114)}vs{comp.get('y2',115)}.xlsx")
                start_response("200 OK", [
                    ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    ("Content-Disposition", f'attachment; filename="{ascii_fname}"; filename*=UTF-8\'\'{utf8_fname}'),
                    ("Content-Length", str(len(xlsx_bytes))),
                ])
                return [xlsx_bytes]
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

        if path == "/api/export_ezsearch_excel":
            try:
                length = int(environ.get("CONTENT_LENGTH") or 0)
                raw_body = environ["wsgi.input"].read(length).decode("utf-8")
                body = json.loads(raw_body)
                items = body.get("items", [])
                xlsx_bytes = generate_ezsearch_excel_bytes(items)
                today_str = datetime.date.today().strftime("%Y%m%d")
                fname = f"MOPS_ezsearch_{today_str}.xlsx"
                start_response("200 OK", [
                    ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    ("Content-Disposition", f'attachment; filename="{fname}"'),
                    ("Content-Length", str(len(xlsx_bytes))),
                ])
                return [xlsx_bytes]
            except Exception as e:
                return _json_response(start_response, {"ok": False, "error": str(e)}, "500 Internal Server Error")

    return _json_response(start_response, {"ok": False, "error": "Not found"}, "404 Not Found")


# 讓 PythonAnywhere 的預設 `from flask_app import app as application` 直接生效
app = application


if __name__ == "__main__":
    from wsgiref.simple_server import make_server
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", 8765))
    host = os.environ.get("HOST", "0.0.0.0")
    print(f"============================================================")
    print(f" MOPS 金融股月自結損益 (WSGI / PythonAnywhere 版) 已啟動！")
    print(f" 本機測試網址: http://127.0.0.1:{port}")
    print(f"============================================================")
    with make_server(host, port, application) as httpd:
        httpd.serve_forever()
