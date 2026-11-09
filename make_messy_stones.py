"""產生主示範用的「髒」結石處置資料（全部亂數，不是任何病人）。

刻意做亂：欄名中英混用、大小單位不一（12mm／1.2cm／約1公分／空白）、日期民國西元混用、
合併儲存格當標題、備註寫在數字格、重複病人、打錯值（年齡 245）、
主表右邊放一張月統計小表、下方放一張圖。
固定 seed，重跑結果一樣：python make_messy_stones.py → 假資料-結石處置-髒.xlsx
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

random.seed(20261112)

TX = {"ESWL": 0.62, "URS": 0.85, "RIRS": 0.80, "PCNL": 0.90}
LOCS = ["腎上盞", "腎中盞", "腎下盞", "腎盂", "近端輸尿管", "遠端輸尿管"]


def pick_tx(size, loc):
    if "輸尿管" in loc:
        return random.choice(["URS", "URS", "ESWL"])
    if size >= 20:
        return random.choice(["PCNL", "PCNL", "RIRS"])
    return random.choice(["ESWL", "RIRS", "ESWL", "RIRS", "PCNL"])


def messy_size(mm):
    r = random.random()
    if r < 0.55:
        return f"{mm}mm" if random.random() < 0.5 else mm
    if r < 0.80:
        return f"{mm / 10:.1f}cm"
    if r < 0.90:
        return f"約{round(mm / 10)}公分"
    return None


def messy_date(d):
    r = random.random()
    if r < 0.5:
        return d  # 真日期
    if r < 0.8:
        return f"{d.year - 1911}/{d.month}/{d.day}"  # 民國
    return d.strftime("%Y.%m.%d")


wb = Workbook()
ws = wb.active
ws.title = "結石登錄(2025-26)"

ws.merge_cells("A1:I1")
ws["A1"] = "泌尿科 結石處置登錄表　※ 假資料，示範用"
ws.append([])
ws.append(["編號", "處置日", "Age", "性別", "大小", "location", "處置方式", "1個月 SF?", "住院(天)"])

rows, start = [], date(2025, 1, 6)
for n in range(1, 121):
    loc = random.choice(LOCS)
    size = round(random.uniform(5, 30), 1)
    tx = pick_tx(size, loc)
    p = TX[tx] - (0.12 if loc == "腎下盞" and tx == "ESWL" else 0) - max(0, size - 15) * 0.01
    sf = random.random() < p
    los = {"ESWL": (0, 1), "URS": (1, 2), "RIRS": (1, 3), "PCNL": (3, 6)}[tx]
    d = start + timedelta(days=random.randint(0, 600))
    rows.append([
        f"S{n:03d}", messy_date(d), random.randint(25, 85),
        random.choice(["男", "男", "女", "M", "F"]), messy_size(size), loc,
        tx if random.random() < 0.85 else tx.lower(),
        random.choice(["是", "Y", "yes"]) if sf else random.choice(["否", "N", "no"]),
        random.randint(*los),
    ])

# 髒點：打錯值、備註寫進數字格、重複病人、缺值
rows[17][2] = 245
rows[33][8] = "3(轉ICU)"
rows[51][7] = "待追蹤"
rows[66][2] = None
rows.insert(90, list(rows[40]))  # 重複病人
rows.insert(25, ["", "", "", "", "", "", "", "", ""])  # 空白列
for r in rows:
    ws.append(r)

# 右邊的小表：月統計（手打，故意跟主表對不上）
ws["L3"], ws["M3"], ws["N3"] = "月份", "ESWL", "內視鏡"
for i, (m, a, b) in enumerate([("1月", 5, 7), ("2月", 4, 6), ("3月", 6, 9), ("4月", 3, "8?"), ("5月", 5, 7), ("6月", "", 6)]):
    ws.cell(row=4 + i, column=12, value=m)
    ws.cell(row=4 + i, column=13, value=a)
    ws.cell(row=4 + i, column=14, value=b)
ws["L11"] = "※ 6月ESWL機器維修，資料待補"

# 右下的圖
chart = BarChart()
chart.title = "各月處置量"
data = Reference(ws, min_col=13, min_row=3, max_col=14, max_row=9)
cats = Reference(ws, min_col=12, min_row=4, max_row=9)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "L14")

out = Path(__file__).with_name("假資料-結石處置-髒.xlsx")
wb.save(out)
print(out, len(rows), "列")
