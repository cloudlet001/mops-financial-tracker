import os
import sys
import json
import urllib.request
import pandas as pd
from bs4 import BeautifulSoup

# Ensure UTF-8 stdout
sys.stdout.reconfigure(encoding='utf-8')

OUTPUT_DIR = r"C:\Users\USER\.gemini\antigravity\scratch\mops_monthly_profit"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def fetch_mops_monthly_profit(year_month="11509", market_type="1"):
    """
    下載公開資訊觀測站月自結損益彙總表
    year_month: 民國年月 (例: 11509 代表民國115年9月)
    market_type: 市場別 (1: 上市, 2: 上櫃, 3: 興櫃)
    """
    api_url = "https://mops.twse.com.tw/mops/api/redirectToOld"
    payload = {
        "apiName": "ajax_t138sb01",
        "parameters": {
            "CK2": "1",
            "YM1": year_month,
            "CK1": market_type,
            "encodeURIComponent": 1,
            "firstin": 1,
            "step": 1,
            "first": 1
        }
    }

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Content-Type": "application/json",
        "Referer": "https://mops.twse.com.tw/mops/",
        "Origin": "https://mops.twse.com.tw"
    }

    print(f"[*] 正在請求 MOPS API (民國 {year_month} 年月, 市場別 {market_type})...")
    req = urllib.request.Request(
        api_url, 
        data=json.dumps(payload).encode('utf-8'), 
        headers=headers, 
        method="POST"
    )

    with urllib.request.urlopen(req, timeout=30) as resp:
        res_json = json.loads(resp.read().decode('utf-8'))

    if res_json.get("code") != 200 or "result" not in res_json:
        raise RuntimeError(f"API 回應錯誤: {res_json}")

    target_url = res_json["result"]["url"]
    print(f"[*] 取得資料來源 URL: {target_url}")

    # 下載報表網頁
    html_req = urllib.request.Request(target_url, headers={
        "User-Agent": headers["User-Agent"],
        "Referer": "https://mops.twse.com.tw/mops/"
    })
    with urllib.request.urlopen(html_req, timeout=30) as html_resp:
        raw_html = html_resp.read().decode('utf-8', errors='ignore')

    soup = BeautifulSoup(raw_html, 'html.parser')
    tables = soup.find_all('table')

    # 解析個別自結損益 (Table 2)
    df_indiv = pd.DataFrame()
    if len(tables) > 2:
        table_indiv = tables[2]
        headers_indiv = [th.text.strip() for th in table_indiv.find_all('th')]
        rows_indiv = []
        for tr in table_indiv.find_all('tr'):
            tds = tr.find_all('td')
            if tds and len(tds) == len(headers_indiv):
                rows_indiv.append([td.text.strip().replace('\xa0', ' ') for td in tds])
        if headers_indiv:
            df_indiv = pd.DataFrame(rows_indiv, columns=headers_indiv)

    # 解析合併自結損益 (Table 4)
    df_merge = pd.DataFrame()
    if len(tables) > 4:
        table_merge = tables[4]
        headers_merge = []
        for th in table_merge.find_all('th'):
            h = th.text.strip()
            if h not in headers_merge:
                headers_merge.append(h)
        rows_merge = []
        for tr in table_merge.find_all('tr'):
            tds = tr.find_all('td')
            if tds and len(tds) == len(headers_merge):
                rows_merge.append([td.text.strip().replace('\xa0', ' ') for td in tds])
        if headers_merge:
            df_merge = pd.DataFrame(rows_merge, columns=headers_merge)

    print(f"[+] 解析完成: 個別自結損益共 {len(df_indiv)} 筆，合併自結損益共 {len(df_merge)} 筆。")

    # 輸出檔案
    indiv_csv = os.path.join(OUTPUT_DIR, f"mops_monthly_profit_{year_month}_individual.csv")
    merge_csv = os.path.join(OUTPUT_DIR, f"mops_monthly_profit_{year_month}_consolidated.csv")
    excel_file = os.path.join(OUTPUT_DIR, f"mops_monthly_profit_{year_month}_all.xlsx")
    json_file = os.path.join(OUTPUT_DIR, f"mops_monthly_profit_{year_month}.json")

    # 存為 UTF-8 with BOM (utf-8-sig)，讓 Excel 開啟不亂碼
    df_indiv.to_csv(indiv_csv, index=False, encoding='utf-8-sig')
    df_merge.to_csv(merge_csv, index=False, encoding='utf-8-sig')

    # 存為 Excel 多分頁
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        df_indiv.to_excel(writer, sheet_name='個別自結損益', index=False)
        df_merge.to_excel(writer, sheet_name='合併自結損益', index=False)

    # 存為 JSON
    all_data = {
        "period": year_month,
        "query_time": res_json.get("datetime", ""),
        "individual": df_indiv.to_dict(orient="records"),
        "consolidated": df_merge.to_dict(orient="records")
    }
    with open(json_file, 'w', encoding='utf-8') as jf:
        json.dump(all_data, jf, ensure_ascii=False, indent=2)

    print(f"[+] 個別 CSV 儲存於: {indiv_csv}")
    print(f"[+] 合併 CSV 儲存於: {merge_csv}")
    print(f"[+] Excel 儲存於: {excel_file}")
    print(f"[+] JSON 儲存於: {json_file}")

    return df_indiv, df_merge

if __name__ == "__main__":
    fetch_mops_monthly_profit("11509", "1")
