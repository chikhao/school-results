from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)
CORS(app)

# تحميل الملف عند بدء التشغيل
EXCEL_PATH = "chat.xlsx"

# الأقسام المرمّزة وأسماؤها
SECTION_MAP = {
    "3111001": "ج م آ 1",
    "3111002": "ج م آ 2",
    "3112001": "ج م ع ت 1",
    "3112002": "ج م ع ت 2",
    "3112003": "ج م ع ت 3",
}

# تحميل كل الأقسام مرة واحدة في الذاكرة
sections_data = {}

for sheet_id, label in SECTION_MAP.items():
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_id, skiprows=7)
        df = df[['اللقب', 'الاسم', 'تاريخ الميلاد', df.columns[6], df.columns[7]]]
        df.columns = ['last_name', 'first_name', 'birth_date', 'avg_grade', 'exam_grade']
        df['birth_date'] = pd.to_datetime(df['birth_date']).dt.date
        sections_data[label] = df
    except Exception as e:
        print(f"Error loading sheet {sheet_id}: {e}")

@app.route("/api/results", methods=["POST"])
def get_result():
    data = request.get_json()
    section = data.get("section")
    birth_date_str = data.get("birth_date")

    if not section or not birth_date_str:
        return jsonify({"error": "الرجاء اختيار القسم وإدخال تاريخ الميلاد"}), 400

    try:
        birth_date = pd.to_datetime(birth_date_str).date()
    except:
        return jsonify({"error": "تاريخ الميلاد غير صالح"}), 400

    df = sections_data.get(section)
    if df is None:
        return jsonify({"error": "القسم غير موجود"}), 404

    match = df[df["birth_date"] == birth_date]

    if match.empty:
        return jsonify({"error": "لا توجد نتائج لهذا التاريخ في هذا القسم"}), 404

    student = match.iloc[0]
    return jsonify({
        "last_name": student["last_name"],
        "first_name": student["first_name"],
        "avg_grade": student["avg_grade"],
        "exam_grade": student["exam_grade"]
    })

if __name__ == "__main__":
    app.run(debug=True)
