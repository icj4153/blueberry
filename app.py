from flask import Flask, request, send_file, render_template_string
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from io import BytesIO
import datetime
from collections import defaultdict

app = Flask(__name__)

HTML_TEMPLATE = """
<!doctype html>
<html lang=\"ko\">
  <head>
    <meta charset=\"utf-8\">
    <title>배송리스트 → 발주서 변환기</title>
  </head>
  <body style=\"font-family: sans-serif; text-align: center; margin-top: 50px;\">
    <h1>📦 배송리스트 → 발주서 자동 변환</h1>
    <form action=\"/convert\" method=\"post\" enctype=\"multipart/form-data\">
      <input type=\"file\" name=\"file\" accept=\".xlsx\" required><br><br>
      <button type=\"submit\" style=\"font-size: 16px;\">발주서 만들기</button>
    </form>
  </body>
</html>
"""

unit_prices = {
    "14mm이상 400g": 18000,
    "14mm이상 600g": 26000,
    "14mm이상 1kg": 41000,
    "16mm이상 400g": 21000,
    "16mm이상 600g": 30000,
    "18mm이상 400g": 25000,
    "18mm이상 1kg": 59000
}

def format_option_text(text):
    mapping = {
        "대(14mm~16mm)": "14mm이상",
        "특대(16mm~18mm)": "16mm이상",
        "왕특(18mm이상)": "18mm이상"
    }
    for key, val in mapping.items():
        if key in text:
            weight = text.split("_")[-1].strip()
            return f"{val} {weight}"
    return text

@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/convert", methods=["POST"])
def convert():
    file = request.files['file']
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    col_map = {}
    headers = [cell.value for cell in ws[1]]
    for idx, header in enumerate(headers):
        col_map[header] = idx + 1

    dest_ws = wb.create_sheet(title="발주서")
    out_headers = ["주문번호", "주문상품명", "상품모델", "수량", "수취인명", "수취인 우편번호",
                   "수취인 주소", "수취인 전화번호", "수취인 이동통신", "배송메시지", "상품코드", "주문자명"]
    dest_ws.append(out_headers)

    for cell in dest_ws[1]:
        cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    summary = defaultdict(int)

    for row in ws.iter_rows(min_row=2, values_only=True):
        option = row[col_map["등록옵션명"] - 1]
        formatted = format_option_text(option)
        new_row = [
            row[col_map["주문번호"] - 1],
            formatted,
            formatted,
            row[col_map["구매수(수량)"] - 1],
            row[col_map["수취인이름"] - 1],
            row[col_map["우편번호"] - 1],
            row[col_map["수취인 주소"] - 1],
            row[col_map["수취인전화번호"] - 1],
            row[col_map["수취인전화번호"] - 1],
            row[col_map["배송메세지"] - 1],
            row[col_map["주문번호"] - 1],
            row[col_map["구매자"] - 1]
        ]
        dest_ws.append(new_row)
        summary[formatted] += int(row[col_map["구매수(수량)"] - 1])

    for col in dest_ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        dest_ws.column_dimensions[col[0].column_letter].width = max_length + 2

    summary_ws = wb.create_sheet(title="발주 정리표")
    summary_ws.append(["사이즈", "개수", "단가", "합계"])

    total_count = 0
    total_sum = 0
    for key in summary:
        count = summary[key]
        price = unit_prices.get(key, 0)
        total = count * price
        total_count += count
        total_sum += total
        summary_ws.append([key, count, price, total])

    summary_ws.append(["", total_count, "", total_sum])
    for col in summary_ws.columns:
        summary_ws.column_dimensions[col[0].column_letter].auto_size = True

    buffer = BytesIO()
    today = datetime.datetime.today().strftime("%y%m%d")
    filename = f"{today} 하입월드 발주서.xlsx"
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    app.run(debug=True)
