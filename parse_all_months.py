import sys
from bs4 import BeautifulSoup
import re
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

def parse_year(year):
    with open(f"C:/Users/USER/.gemini/antigravity/scratch/mops_monthly_profit/2881_{year}_full.html", "r", encoding="utf-8") as f:
        html = f.read()
    
    soup = BeautifulSoup(html, 'html.parser')
    
    # We have pair of noBorder table (containing "114年X月") and hasBorder table (containing metrics)
    # Let's verify
    no_border_tables = [t for t in soup.find_all('table', class_='noBorder') if f'{year}年' in t.text]
    has_border_tables = soup.find_all('table', class_='hasBorder')
    
    print(f"Year {year}: noBorder count = {len(no_border_tables)}, hasBorder count = {len(has_border_tables)}")
    
    records = []
    for nb, hb in zip(no_border_tables, has_border_tables):
        # Extract month
        m_match = re.search(rf'{year}年\s*(\d+)月', nb.text)
        month = int(m_match.group(1)) if m_match else None
        
        row_dict = {"月份": f"{month}月"}
        for tr in hb.find_all('tr'):
            th = tr.find('th')
            td = tr.find('td')
            if th and td:
                key = th.text.strip().replace('（仟元）：', '').replace('（最多60個中文字）：', '').replace('：', '')
                val = td.text.strip().replace('\xa0', ' ')
                row_dict[key] = val
        records.append(row_dict)
    
    df = pd.DataFrame(records)
    return df

df_114 = parse_year("114")
print("\n--- Year 114 ---")
print(df_114.to_string())

df_115 = parse_year("115")
print("\n--- Year 115 ---")
print(df_115.to_string())
