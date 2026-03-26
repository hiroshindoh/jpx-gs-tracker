import json, re, requests
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl

TARGET = "ゴールドマン"
HEADERS = {"User-Agent": "Mozilla/5.0"}

def get(url):
    print("DL: " + url)
    r = requests.get(url, headers=HEADERS, timeout=30)
    if r.status_code == 200:
        return r.content
    print("NG: HTTP " + str(r.status_code))
    return None

def last_friday():
    d = datetime.today()
    days_back = (d.weekday() + 3) % 7 or 7
    return d - timedelta(days=days_back)

def oi_url(dt):
    return (
        "https://www.jpx.co.jp/automation/markets/derivatives/"
        "open-interest/files/" + dt.strftime("%Y") + "/"
        + dt.strftime("%Y%m%d") + "_nk225op_oi_by_tp.xlsx"
    )

def vol_url(dt, suffix="whole_day"):
    return (
        "https://www.jpx.co.jp/automation/markets/derivatives/"
        "participant-volume/files/daily/" + dt.strftime("%Y%m") + "/"
        + dt.strftime("%Y%m%d") + "_volume_by_participant_" + suffix + ".xlsx"
    )

def parse_oi(raw, date_str):
    wb = openpyxl.load_workbook(BytesIO(raw), read_only=True, data_only=True)
    ws = wb.active
    results, strikes = [], set()
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

def parse_vol(raw, date_str):
    wb = openpyxl.load_workbook(BytesIO(raw), read_only=True, data_only=True)
    ws = wb.active
    results = []
    for row in ws.iter_rows(values_only=True):
        row = list(row)
        if len(row) < 8: continue
        if TARGET not in str(row[5] or ""): continue
        product = str(row[0] or "")
        contract = str(row[2] or "")
        vol = row[7]
        if isinstance(vol, str) and vol.startswith("="):
            vol = float(vol[1:])
        if not vol: continue
        opt_type, strike = None, None
        if product == "NK225E":
            m = re.search(r'([PC])(\d{4})-(\d+)', contract)
            if m:
                opt_type = "CALL" if m.group(1) == "C" else "PUT"
                strike = int(m.group(3))
        results.append({
            "date": date_str,
            "product": product,
            "contract": contract,
            "type": opt_type,
            "strike": strike,
            "volume": int(vol)
        })
    return results

def main():
    today = datetime.today()
    today_str = today.strftime("%Y-%m-%d")
    print("今日: " + today_str)

    # ── 建玉残高（週次・前週金曜日）──
    friday = last_friday()
    date_str = friday.strftime("%Y-%m-%d")
    raw = get(oi_url(friday))
    if raw is None:
        friday2 = friday - timedelta(days=7)
        date_str = friday2.strftime("%Y-%m-%d")
        raw = get(oi_url(friday2))
    oi_data, all_strikes = ([], []) if raw is None else parse_oi(raw, date_str)

    # ── 日次取引量（今日 → 昨日 → 一昨日と遡る）──
    vol_data = []
    vol_date = ""
    for days_back in range(5):
        dt = today - timedelta(days=days_back)
        raw_vol = get(vol_url(dt))
        if raw_vol:
            vol_date = dt.strftime("%Y-%m-%d")
            vol_data = parse_vol(raw_vol, vol_date)
            print("取引量データ取得: " + vol_date)
            break

    # ── 出力 ──
    for d in oi_data:
        print(str(d["strike"]) + "円 " + d["type"] + " " + d["side"] + ": " + str(d["qty"]) + "枚")
    for d in vol_data:
        if d["type"]:
            print("[VOL] " + str(d["type"]) + " " + str(d["strike"]) + "円: " + str(d["volume"]) + "枚")

    output = {
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "oi": {
            "date": date_str,
            "gs": oi_data,
            "strikes": {
                "min": min(all_strikes) if all_strikes else 0,
                "max": max(all_strikes) if all_strikes else 0,
                "all": all_strikes
            }
        },
        "vol": {
            "date": vol_date,
            "gs": vol_data
        }
    }

    with open("gs_data.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print("完了 ✅")

if __name__ == "__main__":
    main()
