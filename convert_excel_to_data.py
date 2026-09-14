# -*- coding: utf-8 -*-
"""
用途：把整理好的 Excel 或 CSV 名單，轉換成網頁用的 data.js
使用方式：
    python3 convert_excel_to_data.py 你的檔名.xlsx
    python3 convert_excel_to_data.py 你的檔名.csv

支援 .xlsx / .xls / .csv 三種格式，程式會自動依副檔名判斷。

欄位需求（跟原始「一碼掃名單」一樣）：
    序號 / 商店名稱 / 地址 / 區域 / 電話 / 類別 / 連結 / (可選)Unnamed: 7 第二連結 / (可選)精選

類別欄位請填：住 / 運 / 遊
    住 -> 住宿
    運 -> 交通票券
    遊 -> 門票・體驗

「精選」欄位是選填的：填「是」「V」「1」都算精選，其餘（含空白）都不算。
不加這個欄位也完全沒問題，程式會當作沒有精選店家處理。

執行後會在同一個資料夾產生新的 data.js，直接覆蓋舊檔即可。
"""
import sys
import os
import json
import re
import pandas as pd

CAT_MAP = {'住': '住宿', '運': '交通票券', '遊': '門票・體驗'}

PLATFORM_RULES = [
    ('booking.com', 'Booking.com'),
    ('agoda.com', 'Agoda'),
    ('klook.com', 'Klook'),
    ('kkday.com', 'KKday'),
    ('eztravel.com.tw', '易遊網 ezTravel'),
    ('ctrip.com', '攜程 Ctrip'),
    ('expedia.com', 'Expedia'),
]


def normalize_url(url):
    url = url.strip()
    if not url:
        return None
    if not url.startswith('http'):
        url = 'https://' + url
    return url


def detect_platform(url):
    for domain, label in PLATFORM_RULES:
        if domain in url:
            return label
    m = re.search(r'https?://(?:www\.)?([^/]+)', url)
    return m.group(1) if m else '訂購連結'


def load_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        # 先試 utf-8-sig（處理 Excel 匯出 CSV 常見的 BOM），失敗再退回系統編碼
        try:
            return pd.read_csv(path, encoding='utf-8-sig')
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding='big5')
    elif ext in ('.xlsx', '.xls'):
        return pd.read_excel(path)
    else:
        raise ValueError(f"不支援的檔案格式：{ext}（請提供 .xlsx 或 .csv）")


def convert(src_path, out_path='data.js'):
    df = load_table(src_path)
    records = []
    skipped = []

    extra_link_col = 'Unnamed: 7' if 'Unnamed: 7' in df.columns else None

    for i, row in df.iterrows():
        name = str(row['商店名稱']).strip()
        address = str(row['地址']).strip()
        region = str(row['區域']).strip()
        phone_raw = row.get('電話')
        phone = '' if pd.isna(phone_raw) else str(phone_raw).strip()
        cat_code = str(row['類別']).strip()
        category = CAT_MAP.get(cat_code, cat_code)

        featured_raw = row.get('精選')
        featured = False
        if not pd.isna(featured_raw):
            featured = str(featured_raw).strip().upper() in ('是', 'Y', 'YES', 'V', '1', 'TRUE')

        links = []
        seen = set()
        cols = ['連結'] + ([extra_link_col] if extra_link_col else [])
        for col in cols:
            val = row.get(col)
            if isinstance(val, str) and val.strip():
                url = normalize_url(val)
                if url and url not in seen:
                    seen.add(url)
                    links.append({'platform': detect_platform(url), 'url': url})

        if not links:
            skipped.append(name)
            continue

        records.append({
            'id': i + 1,
            'name': name,
            'address': address,
            'region': region,
            'phone': phone,
            'category': category,
            'featured': featured,
            'links': links,
        })

    js = "// 屏東好好玩 - 票券資料（由 convert_excel_to_data.py 自動產生，請勿手動大幅修改結構）\n"
    js += "window.PINGTUNG_DATA = " + json.dumps(records, ensure_ascii=False, indent=2) + ";\n"

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(js)

    print(f"完成！共轉換 {len(records)} 筆，輸出至 {out_path}")
    if skipped:
        print(f"以下 {len(skipped)} 筆因為沒有訂購連結被跳過：")
        for s in skipped:
            print("  -", s)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法：python3 convert_excel_to_data.py 你的檔名.xlsx（或 .csv）")
        sys.exit(1)
    convert(sys.argv[1])
