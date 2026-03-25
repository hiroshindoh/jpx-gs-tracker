import json, requests
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
    return (
        "https://www.jpx.co.jp/automation/markets/derivatives/"
        "open-interest/files/"
        + dt.strftime("%Y") + "/"
        + dt.strftime("%Y%m%d") + "_nk225op_oi_by_tp.xlsx"
    )

def download_excel(url):
    print("DL: " + url)
    r = requests.get(url, headers=HEADERS, timeout=30)
    if r.status_code == 200:
        return r.content
    print("NG: HTTP " + str(r.status_code))
    return None

def parse_excel(raw, date_str):
    wb = openpyxl.load_workbook(BytesIO(raw), read_only=True, data_only=True)
    ws = wb.active
    results = []
    strikes = set()
    for row in ws.iter_rows(values_only=True):
        row = list(row)
        ps = row[1] if len(row) > 1 else None
        if ps and isinstance(ps, (int, float)):
            strikes.add(int(ps))
            if TARGET in str(row[3] or ""):
                results.append({"date": date_str, "strike": int(ps), "type": "PUT", "side": "Sell", "qty": int(row[4] or 0)})
            if TARGET in str(row[6] or ""):
                results.append({"date": date_str, "strike": int(ps), "type": "PUT", "side": "Buy", "qty": int(row[7] or 0)})
        cs = row[11] if len(row) > 11 else None
        if cs and isinstance(cs, (int, float)):
            strikes.add(int(cs))
            if TARGET in str(row[13] or ""):
                results.append({"date": date_str, "strike": int(cs), "type": "CALL", "side": "Sell", "qty": int(row[14] or 0)})
            if TARGET in str(row[16] or ""):
                results.append({"date": date_str, "strike": int(cs), "type": "CALL", "side": "Buy", "qty": int(row[17] or 0)})
    return results, sorted(strikes)

def main():
    friday = last_friday()
    date_str = friday.strftime("%Y-%m-%d")
    print("対象日: " + date_str)
    raw = download_excel(build_url(friday))
    if raw is None:
        friday = friday - timedelta(days=7)
        date_str = friday.strftime("%Y-%m-%d")
        raw = download_excel(build_url(friday))
    if raw is None:
        print("ERROR: Excelを取得できませんでした")
        return
    gs_data, all_strikes = parse_excel(raw, date_str)
    for d in gs_data:
        print(str(d["strike"]) + "円 " + d["type"] + " " + d["side"] + ": " + str(d["qty"]) + "枚")
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
    print("完了 ✅")

if __name__ == "__main__":
    main()
