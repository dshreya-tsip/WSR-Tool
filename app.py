from flask import Flask, request, jsonify, render_template
import csv
import os
from openpyxl import Workbook
import traceback

# Create Flask app and specify template/static folders
wsr_app = Flask(__name__, template_folder='templates', static_folder='static')

# Local path to save CSV/Excel files (Render compatible)
DATA_DIR = os.path.join(os.getcwd(), "data")
os.makedirs(DATA_DIR, exist_ok=True)

CSV_FILE = os.path.join(DATA_DIR, "WeeklyStatusReport.csv")
EXCEL_FILE = os.path.join(DATA_DIR, "WeeklyStatusReport.xlsx")

@wsr_app.route('/')
def index():
    return render_template('index.html')

@wsr_app.route('/data', methods=['GET'])
def load_data():
    rows = []
    try:
        if os.path.exists(CSV_FILE):
            with open(CSV_FILE, newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader, None)  # Skip header
                rows = [row for row in reader]
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    return jsonify(rows)

@wsr_app.route('/save', methods=['POST'])
def save_data():
    data = request.json.get('tableData')
    header = ["Year", "Month", "Week Days", "Etria", "Solutions"]

    try:
        # Save CSV
        with open(CSV_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerows(data)

        # Save Excel
        wb = Workbook()
        ws = wb.active
        ws.append(header)
        for row in data:
            ws.append(row)
        wb.save(EXCEL_FILE)

        return jsonify({'message': 'Data saved successfully to CSV and Excel.'})

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    wsr_app.run(host='0.0.0.0', port=8080, debug=True)

app = wsr_app