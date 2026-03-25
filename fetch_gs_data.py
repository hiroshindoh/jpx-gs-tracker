import json
import requests
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl

TARGET = "ゴールドマン"
HEADERS = {"User-Agent": "Mozilla/5.0"}

def last_friday():
    d = datetime.today()
    days_back = (d.weekday() + 3) % 7 or 7
    return d - timedelta(days=days_back)

def build_url(dt):
    yyyymm   = dt.strftime("%Y%m")
    yyyymmdd = dt.strftime("%Y%m%d")
    return (
        "https://www.jpx.co.jp/automation/markets/derivatives/"
        f"open-interest/files/{yyyymm}/{yyyymmdd}_nk225op_oi_by_tp_2.xlsx"
    )def download_excel(url):
    print(f"DL: {url}")
    r = requests.get(url, headers=HEADERS, timeout=30)
    if r.status_code == 200:
        print(f"OK: {len(r.content):,} bytes")
        return r.content
    print(f"NG: HTTP {r.status_code}")
    return None

def parse_excel(raw, date_str):
    wb = openpyxl.load_workbook(BytesIO(raw), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    results = []
    strikes = set()
    for row in rows:
        row = list(row)
        put_s = row[1] if len(row) > 1 else None
        if put_s and isinstance(put_s, (int, float)):
            strikes.add(int(put_s))
            if TARGET in str(row[3] or ""):
                results.append({"date": date_str, "strike": int(put_s), "type": "PUT", "side": "Sell", "qty": int(row[4] or 0)})
            if TARGET in str(row[6] or ""):
                results.append({"date": date_str, "strike": int(put_s), "type": "PUT", "side": "Buy", "qty": int(row[7] or 0)})
        call_s = row[11] if len(row) > 11 else None
        if call_s and isinstance(call_s, (int, float)):
            strikes.add(int(call_s))
            if TARGET in str(row[13] or ""):
                results.append({"date": date_str, "strike": int(call_s), "type": "CALL", "side": "Sell", "qty": int(row[14] or 0)})
            if TARGET in str(row[16] or ""):
                results.append({"date": date_str, "strike": int(call_s), "type": "CALL", "side": "Buy", "qty": int(row[17] or 0)})
    return results, sorted(strikes)
  def main():
    friday = last_friday()
    date_str = friday.strftime("%Y-%m-%d")
    print(f"対象日: {date_str}")
    url = build_url(friday)
    raw = download_excel(url)
    if raw is None:
        friday2 = friday - timedelta(days=7)
        date_str = friday2.strftime("%Y-%m-%d")
        url = build_url(friday2)
        print(f"1週前を試みます: {date_str}")
        raw = download_excel(url)
    if raw is None:
        print("ERROR: Excelを取得できませんでした")
        return
    gs_data, all_strikes = parse_excel(raw, date_str)
    for d in gs_data:
        print(f"  {d['strike']:,}円 {d['type']} {d['side']}: {d['qty']}枚")
    output = {
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "date": date_str,
        "gs": gs_data,
        "strikes": {
            "min": min(all_strikes) if all_strikes else 0,
            "max": max(all_strikes) if all_strikes else 0,
            "all": all_strikes
        }
    }
    with open("gs_data.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print("gs_data.json を出力しました ✅")

if __name__ == "__main__":
    main()
