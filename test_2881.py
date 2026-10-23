import urllib.request
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = "https://mops.twse.com.tw/mops/api/redirectToOld"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/json",
    "Referer": "https://mops.twse.com.tw/mops/",
    "Origin": "https://mops.twse.com.tw"
}

for year in ["114", "115"]:
    payload = {
        "apiName": "ajax_t138sb02_q1",
        "parameters": {
            "co_id": "2881",
            "year": year,
            "encodeURIComponent": 1,
            "step": 1,
            "firstin": 1,
            "off": 1
        }
    }
    req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            print(f"Year {year} API response code:", data.get("code"))
            target_url = data["result"]["url"]
            print(f"Year {year} target url:", target_url)
            
            # Fetch target page
            html_req = urllib.request.Request(target_url, headers={
                "User-Agent": headers["User-Agent"],
                "Referer": "https://mops.twse.com.tw/mops/"
            })
            with urllib.request.urlopen(html_req, timeout=30) as hresp:
                raw_html = hresp.read().decode('utf-8', errors='ignore')
                with open(f"C:/Users/USER/.gemini/antigravity/scratch/mops_monthly_profit/raw_2881_{year}.html", "w", encoding="utf-8") as out:
                    out.write(raw_html)
                print(f"Year {year} saved, length: {len(raw_html)}")
    except Exception as e:
        print(f"Year {year} error:", e)
