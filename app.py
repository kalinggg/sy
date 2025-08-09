from flask import Flask, render_template, request, jsonify, send_file
from datetime import datetime
import os
import json
import xlsxwriter
from io import BytesIO

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'quotations'

# 確保儲存目錄存在
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])


@app.route('/')
def index():
    # 載入歷史報價單列表
    quotations = []
    if os.path.exists(app.config['UPLOAD_FOLDER']):
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            if filename.endswith('.json'):
                with open(os.path.join(app.config['UPLOAD_FOLDER'], filename), 'r', encoding='utf-8') as f:
                    try:
                        data = json.load(f)
                        quotations.append({
                            'id': filename.replace('.json', ''),
                            'date': data.get('date', ''),
                            'customer': data.get('customer', ''),
                            'total': data.get('grand_total', 0)
                        })
                    except:
                        continue

    # 產生新的報價單編號
    new_quotation_number = f"QTN-{datetime.now().strftime('%Y%m%d-%H%M%S')}"

    return render_template('index.html',
                           quotation_number=new_quotation_number,
                           current_date=datetime.now().strftime('%Y-%m-%d'),
                           quotations=quotations)


@app.route('/save', methods=['POST'])
def save_quotation():
    data = request.json
    quotation_id = data.get('quotation_number', '')

    if not quotation_id:
        return jsonify({'success': False, 'message': '報價單編號不能為空'})

    # 儲存報價單資料
    filename = os.path.join(app.config['UPLOAD_FOLDER'], f"{quotation_id}.json")
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return jsonify({'success': True, 'message': '報價單儲存成功'})


@app.route('/load/<quotation_id>')
def load_quotation(quotation_id):
    filename = os.path.join(app.config['UPLOAD_FOLDER'], f"{quotation_id}.json")

    if not os.path.exists(filename):
        return jsonify({'success': False, 'message': '報價單不存在'})

    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)

    return jsonify({'success': True, 'data': data})


@app.route('/delete/<quotation_id>')
def delete_quotation(quotation_id):
    filename = os.path.join(app.config['UPLOAD_FOLDER'], f"{quotation_id}.json")

    if os.path.exists(filename):
        os.remove(filename)
        return jsonify({'success': True, 'message': '報價單已刪除'})
    else:
        return jsonify({'success': False, 'message': '報價單不存在'})


@app.route('/export/txt/<quotation_id>')
def export_txt(quotation_id):
    filename = os.path.join(app.config['UPLOAD_FOLDER'], f"{quotation_id}.json")

    if not os.path.exists(filename):
        return jsonify({'success': False, 'message': '報價單不存在'})

    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 產生TXT內容
    txt_content = f"報價單編號: {data.get('quotation_number', '')}\n"
    txt_content += f"日期: {data.get('date', '')}\n"
    txt_content += f"客戶名稱: {data.get('customer', '')}\n"
    txt_content += f"聯絡人: {data.get('contact_person', '')}\n\n"
    txt_content += "項目列表:\n"

    for idx, item in enumerate(data.get('items', []), 1):
        txt_content += f"{idx}. {item.get('description', '')} - 數量: {item.get('quantity', 0)} - 單價: {item.get('unit_price', 0)} - 金額: {item.get('amount', 0)}\n"
        if item.get('notes', ''):
            txt_content += f"   備註: {item.get('notes', '')}\n"

    txt_content += f"\n總金額: ${data.get('grand_total', 0)}\n"
    txt_content += f"地址: {data.get('address', '')}\n"
    txt_content += f"備註說明: {data.get('notes', '')}\n"

    # 建立記憶體中的檔案物件
    txt_file = BytesIO()
    txt_file.write(txt_content.encode('utf-8'))
    txt_file.seek(0)

    return send_file(
        txt_file,
        as_attachment=True,
        download_name=f"{quotation_id}.txt",
        mimetype='text/plain'
    )


@app.route('/export/excel/<quotation_id>')
def export_excel(quotation_id):
    filename = os.path.join(app.config['UPLOAD_FOLDER'], f"{quotation_id}.json")

    if not os.path.exists(filename):
        return jsonify({'success': False, 'message': '報價單不存在'})

    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 建立記憶體中的Excel檔案
    excel_file = BytesIO()
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet('報價單')

    # 設定格式
    bold = workbook.add_format({'bold': True})
    money_format = workbook.add_format({'num_format': '$#,##0.00'})

    # 寫入標題
    worksheet.write(0, 0, '報價單編號:', bold)
    worksheet.write(0, 1, data.get('quotation_number', ''))
    worksheet.write(1, 0, '日期:', bold)
    worksheet.write(1, 1, data.get('date', ''))
    worksheet.write(2, 0, '客戶名稱:', bold)
    worksheet.write(2, 1, data.get('customer', ''))
    worksheet.write(3, 0, '聯絡人:', bold)
    worksheet.write(3, 1, data.get('contact_person', ''))

    # 寫入項目表頭
    worksheet.write(5, 0, '序號', bold)
    worksheet.write(5, 1, '項目描述', bold)
    worksheet.write(5, 2, '數量', bold)
    worksheet.write(5, 3, '單價', bold)
    worksheet.write(5, 4, '金額', bold)
    worksheet.write(5, 5, '備註', bold)

    # 寫入項目內容
    row = 6
    for idx, item in enumerate(data.get('items', []), 1):
        worksheet.write(row, 0, idx)
        worksheet.write(row, 1, item.get('description', ''))
        worksheet.write(row, 2, item.get('quantity', 0))
        worksheet.write(row, 3, item.get('unit_price', 0), money_format)
        worksheet.write(row, 4, item.get('amount', 0), money_format)
        worksheet.write(row, 5, item.get('notes', ''))
        row += 1

    # 寫入總計
    worksheet.write(row + 1, 0, '總金額:', bold)
    worksheet.write(row + 1, 1, data.get('grand_total', 0), money_format)

    # 寫入地址和備註
    worksheet.write(row + 3, 0, '地址:', bold)
    worksheet.write(row + 3, 1, data.get('address', ''))
    worksheet.write(row + 4, 0, '備註說明:', bold)
    worksheet.write(row + 4, 1, data.get('notes', ''))

    workbook.close()
    excel_file.seek(0)

    return send_file(
        excel_file,
        as_attachment=True,
        download_name=f"{quotation_id}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run(debug=True)