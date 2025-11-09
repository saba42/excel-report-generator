from flask import Flask, render_template_string, request, send_file
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# -------------------- HTML PAGE --------------------
HTML_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Excel Report Generator</title>
<style>
    body {
        font-family: 'Segoe UI', sans-serif;
        background: linear-gradient(135deg, #1a2a5a 0%, #1e3c72 25%, #00c6ff 50%, #1e3c72 75%, #1a2a5a 100%);
        background-attachment: fixed;
        background-size: cover;
        height: 100vh;
        margin: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        overflow: hidden;
    }
    body::before {
        content: "";
        position: absolute;
        inset: 0;
        background: rgba(0, 0, 0, 0.45);
        backdrop-filter: blur(6px);
        z-index: 0;
    }
    .container {
        background-color: #0d1a3a;
        border-radius: 20px;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.4);
        padding: 40px;
        text-align: center;
        width: 400px;
        position: relative;
        z-index: 1;
        backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.3);
    }
    h2 {
        background: linear-gradient(90deg, #00c6ff, #007bff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 26px;
        margin-bottom: 20px;
        font-weight: 700;
    }
    .upload-box {
        border: 2px dashed rgba(255, 255, 255, 0.7);
        padding: 20px;
        border-radius: 12px;
        margin-bottom: 25px;
        transition: 0.3s;
        color: white;
    }
    .upload-box:hover {
        background: rgba(255, 255, 255, 0.1);
        transform: scale(1.02);
    }
    input[type=file] {
        display: none;
    }
    label {
        background: linear-gradient(135deg, #007bff, #00c6ff);
        color: white;
        padding: 10px 20px;
        border-radius: 10px;
        cursor: pointer;
        font-weight: 600;
        transition: 0.3s;
    }
    label:hover {
        background: linear-gradient(135deg, #0056b3, #0099cc);
    }
    .btn {
        background: linear-gradient(135deg, #007bff, #00c6ff);
        color: white;
        border: none;
        padding: 12px 20px;
        margin: 10px 0;
        border-radius: 8px;
        cursor: pointer;
        width: 80%;
        font-size: 15px;
        font-weight: 600;
        transition: 0.3s;
    }
    .btn:hover {
        background: linear-gradient(135deg, #0056b3, #0099cc);
        transform: scale(1.03);
    }
    .footer {
        margin-top: 15px;
        color: rgba(255, 255, 255, 0.8);
        font-size: 13px;
        font-weight: 500;
    }
    .status {
        margin-top: 10px;
        color: #00ffea;
        font-weight: 600;
    }
</style>
<script>
    function showUploading() {
        document.getElementById("status").innerText = "⏳ Uploading...";
    }
</script>
</head>
<body>
<div class="container">
    <h2>Excel Report Generator</h2>
    <form method="POST" enctype="multipart/form-data" onsubmit="showUploading()">
        <div class="upload-box">
            <label for="fileUpload">📂 Choose Excel File</label>
            <input id="fileUpload" type="file" name="file" accept=".xlsx" required>
        </div>
        <button class="btn" name="action" value="participation">Generate Participation Report</button>
        <button class="btn" name="action" value="performance">Generate Performance Report</button>
    </form>
    <div id="status" class="status"></div>
    <div class="footer">Designed with 💙 by Sabapathi</div>
</div>
</body>
</html>
"""

# -------------------- PARTICIPATION REPORT --------------------
def generate_participation_report(input_file, output_file):
    df = pd.read_excel(input_file)
    required_cols = ["Department", "Name", "Test Status"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column: '{col}'")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Data")

    wb = load_workbook(output_file)
    ws = wb["Data"]

    fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Adjust column width
    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 1

    wb.save(output_file)

# -------------------- PERFORMANCE REPORT --------------------
def generate_performance_report(input_file, output_file):
    df = pd.read_excel(input_file)
    generate_participation_report(input_file, output_file)

    pct_col = None
    for col in df.columns:
        if "percentage" in col.lower():
            pct_col = col
            break
    if pct_col is None:
        raise ValueError("Could not find a percentage column in the Excel file.")

    status_index = df.columns.get_loc("Test Status")

    def categorize(p):
        if pd.isna(p):
            return "Not Attended"
        p_str = str(p).strip()
        if p_str in ["-", "", "NA", "n/a"]:
            return "Not Attended"
        try:
            p_val = float(p_str.replace("%", ""))
        except:
            return "Not Attended"
        if p_val > 75:
            return "Good"
        elif p_val > 50:
            return "Satisfactory"
        elif p_val > 25:
            return "Need Attention"
        else:
            return "Intervention"

    df.insert(status_index + 4, "Category", df[pct_col].apply(categorize))
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        pd.pivot_table(df, index="Department", columns="Category", values="Name", aggfunc="count", fill_value=0).reset_index().to_excel(writer, index=False, sheet_name="Performance Summary")

# -------------------- FLASK ROUTE --------------------
@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        try:
            file = request.files.get("file")
            action = request.form.get("action")
            if not file:
                return "⚠️ Please upload an Excel file.", 400

            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)
            output_file = os.path.join(UPLOAD_FOLDER, "Report.xlsx")

            if action == "participation":
                generate_participation_report(filepath, output_file)
            elif action == "performance":
                generate_performance_report(filepath, output_file)
            else:
                return "⚠️ Invalid action.", 400

            return send_file(output_file, as_attachment=True)

        except Exception as e:
            return f"❌ Error processing file: {str(e)}", 500

    return render_template_string(HTML_PAGE)

if __name__ == "__main__":
    app.run(debug=True, port=5002)
