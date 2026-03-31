import calendar, datetime, warnings
warnings.filterwarnings("ignore")
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials/service_account.json", scopes=SCOPES)
svc = build("sheets", "v4", credentials=creds)
SS_ID = "1OMk5vBVvUAJpPCjA4p7bMjlkttJDvN0Lp8bVjvKsIhc"

MONTHS_YM = [
    ("4月",2026,4),("5月",2026,5),("6月",2026,6),
    ("7月",2026,7),("8月",2026,8),("9月",2026,9),
    ("10月",2026,10),("11月",2026,11),("12月",2026,12),
    ("1月",2027,1),("2月",2027,2),("3月",2027,3),
]
WEEKDAYS = ["月","火","水","木","金","土","日"]

def rgb(h):
    h = h.lstrip("#")
    return {"red": int(h[0:2],16)/255, "green": int(h[2:4],16)/255, "blue": int(h[4:6],16)/255}

C = {
    "title":  rgb("1F4E79"), "hdr2": rgb("2E75B6"),
    "input":  rgb("375623"), "calc": rgb("7F7F7F"),
    "total":  rgb("FFF2CC"), "sat":  rgb("DDEEFF"),
    "sun":    rgb("FFE5E5"), "alt":  rgb("F2F7FC"),
    "white":  rgb("FFFFFF"), "fw":   rgb("FFFFFF"),
    "fb":     rgb("000000"), "green_txt": rgb("375623"),
}

def date_serial(d):
    return (d - datetime.date(1899, 12, 30)).days

# ---------- 既存シート取得 ----------
meta = svc.spreadsheets().get(spreadsheetId=SS_ID).execute()
existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
print("既存シート:", list(existing.keys()))

# ---------- 新規シート追加 ----------
add_reqs = []
for name, year, month in MONTHS_YM:
    sheet_title = f"{name}_{year}"
    if sheet_title not in existing:
        add_reqs.append({"addSheet": {"properties": {"title": sheet_title, "gridProperties": {"rowCount": 50, "columnCount": 23}}}})

if add_reqs:
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SS_ID, body={"requests": add_reqs}).execute()
    print(f"{len(add_reqs)}シート追加完了")

# シートID再取得
meta = svc.spreadsheets().get(spreadsheetId=SS_ID).execute()
sheet_ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

# ---------- 各月セットアップ ----------
DATA_START = 4  # 1-indexed row (行4からデータ)
NUM_COLS = 22

# グループヘッダー定義 (start_col 1-indexed, end_col exclusive)
GROUPS = [
    ("基本情報", 3, 5, True),
    ("売　上",   5, 12, True),
    ("客　数",  12, 17, True),
    ("処 分 率",17, 23, True),
]

# 列定義: (header, is_input, number_format_pattern)
COLS = [
    ("日付",           True,  "M/D"),
    ("曜日",           True,  None),
    ("天気",           True,  None),
    ("気温(℃)",       True,  "0.0"),
    ("前年実績",       True,  "#,##0"),
    ("前年累計",       False, "#,##0"),
    ("本年実績(税抜)", True,  "#,##0"),
    ("本年累計",       False, "#,##0"),
    ("前年対比",       False, "0.0%"),
    ("消費税",         False, "#,##0"),
    ("総売上",         False, "#,##0"),
    ("客数(人)",       True,  "#,##0"),
    ("客数累計",       False, "#,##0"),
    ("客単価(円)",     False, "#,##0"),
    ("前年客数",       True,  "#,##0"),
    ("前年客数累計",   False, "#,##0"),
    ("販売残",         True,  "#,##0"),
    ("製造ロス",       True,  "#,##0"),
    ("試食金額",       True,  "#,##0"),
    ("処分金額",       False, "#,##0"),
    ("処分累計",       False, "#,##0"),
    ("処分率",         False, "0.00%"),
]

COL_WIDTHS = [80,45,70,65,90,90,95,90,70,80,90,70,80,80,70,90,70,70,70,70,80,65]

def cell_fmt(bg, bold=False, fg=None, size=9, wrap=False, halign="CENTER", valign="MIDDLE", num_fmt=None):
    fmt = {
        "backgroundColor": bg,
        "textFormat": {
            "bold": bold,
            "fontSize": size,
            "foregroundColor": fg or C["fb"],
            "fontFamily": "Arial",
        },
        "horizontalAlignment": halign,
        "verticalAlignment": valign,
        "wrapStrategy": "WRAP" if wrap else "OVERFLOW_CELL",
    }
    if num_fmt:
        fmt["numberFormat"] = {"type": "NUMBER", "pattern": num_fmt}
    return fmt

def repeat_cell(sheet_id, r1, c1, r2, c2, fmt):
    return {"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2, "startColumnIndex": c1, "endColumnIndex": c2},
        "cell": {"userEnteredFormat": fmt},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy,numberFormat)",
    }}

def merge(sheet_id, r1, c1, r2, c2):
    return {"mergeCells": {
        "range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2, "startColumnIndex": c1, "endColumnIndex": c2},
        "mergeType": "MERGE_ALL",
    }}

def col_width(sheet_id, c1, c2, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": c1, "endIndex": c2},
        "properties": {"pixelSize": px},
        "fields": "pixelSize",
    }}

def row_height(sheet_id, r1, r2, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": r1, "endIndex": r2},
        "properties": {"pixelSize": px},
        "fields": "pixelSize",
    }}

def freeze(sheet_id, rows=3):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": rows}},
        "fields": "gridProperties(frozenRowCount)",
    }}

for month_name, year, month in MONTHS_YM:
    sheet_title = f"{month_name}_{year}"
    sid = sheet_ids[sheet_title]
    _, num_days = calendar.monthrange(year, month)
    total_row = DATA_START + num_days  # 1-indexed
    last_data = DATA_START + num_days - 1  # 1-indexed

    print(f"セットアップ中: {sheet_title}")

    # ===== VALUES =====
    values_data = []

    # 行1: タイトル
    title_row = [f"{year}年 {month}月　日次売上日報"] + [""] * (NUM_COLS - 1)
    values_data.append(title_row)

    # 行2: グループヘッダー
    grp_row = ["", ""]
    grp_map = {}
    for grp, cs, ce, _ in GROUPS:
        for c in range(cs, ce):
            grp_map[c] = grp
    for ci in range(3, NUM_COLS + 1):
        grp_row.append(grp_map.get(ci, ""))
    # deduplicate (only first col of group gets text)
    seen = set()
    grp_row2 = ["", ""]
    for ci in range(3, NUM_COLS + 1):
        g = grp_map.get(ci, "")
        if g and g not in seen:
            grp_row2.append(g)
            seen.add(g)
        else:
            grp_row2.append("")
    values_data.append(grp_row2)

    # 行3: 列ヘッダー
    hdr_row = [col[0] for col in COLS]
    values_data.append(hdr_row)

    # データ行
    for day in range(1, num_days + 1):
        r = DATA_START + day  # 1-indexed (DATA_START=4, day1→row5 in sheets 1-indexed... wait)
        # r is the 1-indexed row number for formulas
        r1 = DATA_START + day - 1  # row index for this day (1-indexed): day1→row4
        # Actually DATA_START=4 means data starts at row 4 (1-indexed)
        # day1 → row 4, day2 → row 5, etc.
        r_formula = DATA_START + day - 1  # this is the 1-indexed row number

        date = datetime.date(year, month, day)
        wd = date.weekday()
        ds = date_serial(date)

        row = [
            ds,                # A: 日付 (serial number)
            WEEKDAYS[wd],      # B: 曜日
            "",                # C: 天気
            "",                # D: 気温
            "",                # E: 前年実績
            f"=SUM(E{DATA_START}:E{r_formula})",  # F: 前年累計
            "",                # G: 本年実績
            f"=SUM(G{DATA_START}:G{r_formula})",  # H: 本年累計
            f'=IF(E{r_formula}=0,"",G{r_formula}/E{r_formula})',  # I: 前年対比
            f"=ROUND(G{r_formula}*0.1,0)",         # J: 消費税
            f"=G{r_formula}+J{r_formula}",          # K: 総売上
            "",                # L: 客数
            f"=SUM(L{DATA_START}:L{r_formula})",   # M: 客数累計
            f'=IF(L{r_formula}=0,"",ROUND(G{r_formula}/L{r_formula},0))',  # N: 客単価
            "",                # O: 前年客数
            f"=SUM(O{DATA_START}:O{r_formula})",   # P: 前年客数累計
            "",                # Q: 販売残
            "",                # R: 製造ロス
            "",                # S: 試食金額
            f"=Q{r_formula}+R{r_formula}+S{r_formula}",  # T: 処分金額
            f"=SUM(T{DATA_START}:T{r_formula})",   # U: 処分累計
            f'=IF(SUM(K{DATA_START}:K{r_formula})=0,"",U{r_formula}/SUM(K{DATA_START}:K{r_formula}))',  # V: 処分率
        ]
        values_data.append(row)

    # 月合計行
    tr = total_row  # 1-indexed
    total_row_data = [
        "月合計", "", "", "",
        f"=SUM(E{DATA_START}:E{last_data})",   # E
        f"=F{last_data}",                        # F
        f"=SUM(G{DATA_START}:G{last_data})",   # G
        f"=H{last_data}",                        # H
        f'=IF(E{tr}=0,"",G{tr}/E{tr})',        # I
        f"=ROUND(G{tr}*0.1,0)",                  # J
        f"=G{tr}+J{tr}",                         # K
        f"=SUM(L{DATA_START}:L{last_data})",   # L
        f"=M{last_data}",                        # M
        f'=IF(L{tr}=0,"",ROUND(G{tr}/L{tr},0))',  # N
        f"=SUM(O{DATA_START}:O{last_data})",   # O
        f"=P{last_data}",                        # P
        f"=SUM(Q{DATA_START}:Q{last_data})",   # Q
        f"=SUM(R{DATA_START}:R{last_data})",   # R
        f"=SUM(S{DATA_START}:S{last_data})",   # S
        f"=SUM(T{DATA_START}:T{last_data})",   # T
        f"=U{last_data}",                        # U
        f'=IF(K{tr}=0,"",T{tr}/K{tr})',        # V
    ]
    values_data.append(total_row_data)

    # 凡例
    values_data.append([])
    values_data.append(["【凡例】 ■青背景:土曜  ■赤背景:日曜  ■緑ヘッダー:入力欄  ■グレーヘッダー:自動計算"])

    # 値書き込み
    range_name = f"'{sheet_title}'!A1"
    svc.spreadsheets().values().update(
        spreadsheetId=SS_ID,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body={"values": values_data}
    ).execute()

    # ===== FORMATTING =====
    reqs = []

    # 列幅
    for ci, w in enumerate(COL_WIDTHS):
        reqs.append(col_width(sid, ci, ci+1, w))

    # 行高
    reqs.append(row_height(sid, 0, 1, 32))   # タイトル
    reqs.append(row_height(sid, 1, 2, 18))   # グループ
    reqs.append(row_height(sid, 2, 3, 36))   # ヘッダー
    reqs.append(row_height(sid, 3, 3+num_days, 20))  # データ
    reqs.append(row_height(sid, 3+num_days, 3+num_days+1, 24))  # 合計

    # タイトル行マージ・書式
    reqs.append(merge(sid, 0, 0, 1, NUM_COLS))
    reqs.append(repeat_cell(sid, 0, 0, 1, NUM_COLS,
        cell_fmt(C["title"], bold=True, fg=C["fw"], size=14)))

    # グループヘッダー: A,B マージ (行1,2)
    reqs.append(merge(sid, 1, 0, 3, 1))
    reqs.append(repeat_cell(sid, 1, 0, 3, 1,
        cell_fmt(C["hdr2"], bold=True, fg=C["fw"], size=9)))
    reqs.append(merge(sid, 1, 1, 3, 2))
    reqs.append(repeat_cell(sid, 1, 1, 3, 2,
        cell_fmt(C["hdr2"], bold=True, fg=C["fw"], size=9)))

    # グループヘッダー行: グループごとにマージ+色
    for grp, cs, ce, _ in GROUPS:
        reqs.append(merge(sid, 1, cs-1, 2, ce-1))
        reqs.append(repeat_cell(sid, 1, cs-1, 2, ce-1,
            cell_fmt(C["input"], bold=True, fg=C["fw"], size=9)))

    # 列ヘッダー行3
    for ci, (_, is_input, _) in enumerate(COLS):
        bg = C["input"] if is_input else C["calc"]
        reqs.append(repeat_cell(sid, 2, ci, 3, ci+1,
            cell_fmt(bg, bold=True, fg=C["fw"], size=8, wrap=True)))

    # データ行の書式（まず全体に白を適用）
    reqs.append(repeat_cell(sid, 3, 0, 3+num_days, NUM_COLS,
        cell_fmt(C["white"], size=9, halign="CENTER")))

    # 右寄せ列（数値列）
    num_cols = [4,5,6,7,9,10,11,12,13,14,15,16,17,18,19,20]  # 0-indexed
    for r_idx in range(num_days):
        row_i = 3 + r_idx
        date = datetime.date(year, month, 1 + r_idx)
        wd = date.weekday()
        bg = C["sat"] if wd == 5 else C["sun"] if wd == 6 else (C["alt"] if r_idx % 2 == 0 else C["white"])
        reqs.append(repeat_cell(sid, row_i, 0, row_i+1, NUM_COLS,
            cell_fmt(bg, size=9)))

    # 日付列: 数値フォーマット (M/D)
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": 3+num_days, "startColumnIndex": 0, "endColumnIndex": 1},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "M/D"}}},
        "fields": "userEnteredFormat.numberFormat",
    }})

    # 気温フォーマット
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": 3+num_days, "startColumnIndex": 3, "endColumnIndex": 4},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.0"}}},
        "fields": "userEnteredFormat.numberFormat",
    }})

    # 数値列フォーマット (#,##0)
    for ci in [4,5,6,7,9,10,11,12,13,14,15,16,17,18,19,20]:
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": 3+num_days+1, "startColumnIndex": ci, "endColumnIndex": ci+1},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}, "horizontalAlignment": "RIGHT"}},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)",
        }})

    # パーセント列 (I=8, V=21)
    for ci, pat in [(8, "0.0%"), (21, "0.00%")]:
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": 3+num_days+1, "startColumnIndex": ci, "endColumnIndex": ci+1},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": pat}}},
            "fields": "userEnteredFormat.numberFormat",
        }})

    # 月合計行
    reqs.append(merge(sid, 3+num_days, 0, 3+num_days+1, 4))
    reqs.append(repeat_cell(sid, 3+num_days, 0, 3+num_days+1, NUM_COLS,
        cell_fmt(C["total"], bold=True, size=10, halign="RIGHT")))
    reqs.append(repeat_cell(sid, 3+num_days, 0, 3+num_days+1, 4,
        cell_fmt(C["total"], bold=True, size=10, halign="CENTER")))

    # フリーズ
    reqs.append(freeze(sid, rows=3))

    # 天気バリデーション
    reqs.append({"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": 3+num_days, "startColumnIndex": 2, "endColumnIndex": 3},
        "rule": {
            "condition": {"type": "ONE_OF_LIST", "values": [
                {"userEnteredValue": v} for v in ["晴れ","曇り","雨","雪","晴れ/曇り","曇り/雨"]
            ]},
            "showCustomUi": True,
            "strict": True,
        }
    }})

    # 数値バリデーション（売上・客数・処分系）
    for ci in [4, 6, 11, 14, 16, 17, 18]:
        reqs.append({"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 3, "endRowIndex": 3+num_days, "startColumnIndex": ci, "endColumnIndex": ci+1},
            "rule": {
                "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0"}]},
                "showCustomUi": False,
                "strict": True,
                "inputMessage": "0以上の整数を入力してください",
            }
        }})

    svc.spreadsheets().batchUpdate(spreadsheetId=SS_ID, body={"requests": reqs}).execute()
    print(f"  → 完了")

print("\n全シートセットアップ完了!")
print(f"URL: https://docs.google.com/spreadsheets/d/{SS_ID}")
