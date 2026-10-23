import urllib.request
import json
import sys
from bs4 import BeautifulSoup
import re

sys.stdout.reconfigure(encoding='utf-8')

url = "https://mops.twse.com.tw/mops/api/redirectToOld"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/json",
    "Referer": "https://mops.twse.com.tw/mops/",
    "Origin": "https://mops.twse.com.tw"
}

def get_year_data(year):
    payload = {
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
            "COMP": "2881"
        }
    }
    req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        res = json.loads(resp.read().decode('utf-8'))
    target_url = res["result"]["url"]
    hreq = urllib.request.Request(target_url, headers={"User-Agent": headers["User-Agent"]})
    with urllib.request.urlopen(hreq, timeout=30) as hresp:
        html = hresp.read().decode('utf-8', errors='ignore')
    return html

for y in ["114", "115"]:
    html = get_year_data(y)
    with open(f"C:/Users/USER/.gemini/antigravity/scratch/mops_monthly_profit/2881_{y}_full.html", "w", encoding="utf-8") as f:
        f.write(html)
    soup = BeautifulSoup(html, 'html.parser')
    
    # Find all tables or month sections
    tables = soup.find_all('table')
    print(f"\n================ Year {y} (Tables: {len(tables)}) ================")
    
    # Let's inspect text
    # Usually it has: 114年1月 (合併), followed by a table with metrics
    months_found = re.findall(rf'{y}年\s*(\d+)月', html)
    print(f"Months found in text: {months_found}")
    
    # Let's see how each month is formatted
    # Look for tables with class 'hasBorder'
    border_tables = soup.find_all('table', class_='hasBorder')
    print(f"Border tables count: {len(border_tables)}")
    if border_tables:
        first_table = border_tables[0]
        for tr in first_table.find_all('tr'):
            th = tr.find('th')
            td = tr.find('td')
            th_text = th.text.strip() if th else ""
            td_text = td.text.strip() if td else ""
            print(f"  {th_text}: {td_text}")
