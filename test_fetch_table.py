import urllib.request
import json
import urllib.parse
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Approach 1: redirectToOld with ajax_t138sb02
url = "https://mops.twse.com.tw/mops/api/redirectToOld"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/json",
    "Referer": "https://mops.twse.com.tw/mops/",
    "Origin": "https://mops.twse.com.tw"
}

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
        "YEAR": "114",
        "COMP": "2881"
    }
}

try:
    req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        res = json.loads(resp.read().decode('utf-8'))
        print("Approach 1 resp:", res.get("code"))
        target_url = res.get("result", {}).get("url")
        print("target_url:", target_url)
        if target_url:
            hreq = urllib.request.Request(target_url, headers={"User-Agent": headers["User-Agent"]})
            with urllib.request.urlopen(hreq, timeout=30) as hresp:
                html = hresp.read().decode('utf-8', errors='ignore')
                print("HTML length:", len(html))
                print("Snippet:")
                print(html[:1000])
except Exception as e:
    print("Approach 1 error:", e)

# Approach 2: Direct POST to mopsov.twse.com.tw/mops/web/ajax_t138sb02
print("\n--- Approach 2: Direct POST ---")
form_data = urllib.parse.urlencode({
    "encodeURIComponent": 1,
    "step": "1",
    "firstin": "true",
    "off": "1",
    "CK2": "1",
    "BK1": "2",
    "TYPEK": "sii",
    "YEAR": "114",
    "COMP": "2881"
}).encode('utf-8')

try:
    req2 = urllib.request.Request("https://mopsov.twse.com.tw/mops/web/ajax_t138sb02", data=form_data, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://mopsov.twse.com.tw/mops/web/ajax_t138sb02_q1"
    })
    with urllib.request.urlopen(req2, timeout=30) as resp2:
        html2 = resp2.read().decode('utf-8', errors='ignore')
        print("Approach 2 HTML length:", len(html2))
        print("Approach 2 Snippet:")
        print(html2[:1000])
except Exception as e:
    print("Approach 2 error:", e)
