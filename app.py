from flask import Flask, request, send_file, render_template_string
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from io import BytesIO
import datetime

app = Flask(__name__)

# HTML 업로드 UI 템플릿
HTML_TEMPLATE = """
<!doctype html>
<html lang="ko">
  <head>
    <meta charset="utf-8">
    <title>배송리스트 → 발주서 변환기</title>
  </head>
  <body style="font-family: sans-serif; text-align: center; margin-top: 50px;">
    <h1>📦 배송리스트 → 발주서 자동 변환</h1>
    <form action="/convert" method="post" enctype="multipart/form-data">
      <input type="file" name="file" accept=".xlsx" required><br><br>
      <button type="submit" style="font-size: 16px;">발주서 만들기</button>
    </form>
  </body>
</html>
"""

# 옵션 텍스트 포맷 정리 함수
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

    dest_wb = openpyxl.Workbook()
    dest_ws = dest_wb.active
    dest_ws.title = "발주서"

    out_headers = ["주문번호", "주문상품명", "상품모델", "수량", "수취인명", "수취인 우편번호",
                   "수취인 주소", "수취인 전화번호", "수취인 이동통신", "배송메시지", "상품코드", "주문자명"]
    dest_ws.append(out_headers)

    # 헤더 스타일 적용
    for cell in dest_ws[1]:
        cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, values_only=True):
        option = row[col_map["등록옵션명"] - 1]
        formatted = format_option_text(option)
        new_row = [
            row[col_map["주문번호"] - 1],  # 주문번호
            formatted,  # 주문상품명
            formatted,  # 상품모델
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

    for col in dest_ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        dest_ws.column_dimensions[col[0].column_letter].width = max_length + 2

    buffer = BytesIO()
    today = datetime.datetime.today().strftime("%Y%m%d")
    filename = f"발주서_{today}.xlsx"
    dest_wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    app.run(debug=True)
