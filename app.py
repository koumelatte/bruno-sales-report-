from flask import Flask, render_template, request, jsonify
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import datetime, warnings
warnings.filterwarnings("ignore")

app = Flask(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SS_ID = "1OMk5vBVvUAJpPCjA4p7bMjlkttJDvN0Lp8bVjvKsIhc"

MONTH_NAMES = {
    1: "1月", 2: "2月", 3: "3月", 4: "4月", 5: "5月", 6: "6月",
    7: "7月", 8: "8月", 9: "9月", 10: "10月", 11: "11月", 12: "12月"
}

def get_sheets_service():
    creds = Credentials.from_service_account_file(
        "credentials/service_account.json", scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)

def find_row_for_date(service, sheet_name, target_date):
    result = service.spreadsheets().values().get(
        spreadsheetId=SS_ID,
        range=f"'{sheet_name}'!A4:A35"
    ).execute()
    values = result.get("values", [])
    for i, row in enumerate(values):
        if not row:
            continue
        try:
            cell_val = str(row[0])
            if cell_val.isdigit():
                origin = datetime.date(1899, 12, 30)
                cell_date = origin + datetime.timedelta(days=int(cell_val))
            elif len(cell_val.split("/")) == 2:
                m, d = cell_val.split("/")
                cell_date = datetime.date(target_date.year, int(m), int(d))
            else:
                cell_date = datetime.date.fromisoformat(cell_val.replace("/", "-"))
            if cell_date == target_date:
                return 4 + i
        except Exception:
            continue
    return None

@app.route("/")
def index():
    today = datetime.date.today().isoformat()
    return render_template("index.html", today=today)

@app.route("/submit", methods=["POST"])
def submit():
    try:
        date_str  = request.form.get("date")
        weather   = request.form.get("weather", "")
        temp      = request.form.get("temp", "")
        prev_sales = request.form.get("prev_sales", "")
        sales     = request.form.get("sales", "")
        guests    = request.form.get("guests", "")
        prev_guests = request.form.get("prev_guests", "")
        unsold    = request.form.get("unsold", "")
        mfg_loss  = request.form.get("mfg_loss", "")
        tasting   = request.form.get("tasting", "")
        memo      = request.form.get("memo", "")

        target_date = datetime.date.fromisoformat(date_str)
        month = target_date.month
        year = target_date.year
        sheet_name = f"{MONTH_NAMES[month]}_{year}"

        svc = get_sheets_service()
        row_num = find_row_for_date(svc, sheet_name, target_date)

        if row_num is None:
            return jsonify({"success": False, "message": f"該当シートに {date_str} の行が見つかりませんでした。シート名: {sheet_name}"})

        def to_num(v):
            try:
                return int(v.replace(",", "")) if v else ""
            except Exception:
                return ""

        def to_float(v):
            try:
                return float(v) if v else ""
            except Exception:
                return ""

        actual_updates = [
            (f"'{sheet_name}'!C{row_num}", [[weather]]),
            (f"'{sheet_name}'!D{row_num}", [[to_float(temp)]]),
            (f"'{sheet_name}'!E{row_num}", [[to_num(prev_sales)]]),
            (f"'{sheet_name}'!G{row_num}", [[to_num(sales)]]),
            (f"'{sheet_name}'!L{row_num}", [[to_num(guests)]]),
            (f"'{sheet_name}'!O{row_num}", [[to_num(prev_guests)]]),
            (f"'{sheet_name}'!Q{row_num}", [[to_num(unsold)]]),
            (f"'{sheet_name}'!R{row_num}", [[to_num(mfg_loss)]]),
            (f"'{sheet_name}'!S{row_num}", [[to_num(tasting)]]),
        ]
        if memo:
            actual_updates.append((f"'{sheet_name}'!W{row_num}", [[memo]]))

        batch_data = [{"range": r, "values": v} for r, v in actual_updates if v[0][0] != ""]

        if batch_data:
            svc.spreadsheets().values().batchUpdate(
                spreadsheetId=SS_ID,
                body={"valueInputOption": "USER_ENTERED", "data": batch_data}
            ).execute()

        return jsonify({"success": True, "message": f"{target_date.strftime('%Y年%m月%d日')} のデータを記録しました！"})

    except Exception as e:
        return jsonify({"success": False, "message": f"エラー: {str(e)}"})

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5001)
