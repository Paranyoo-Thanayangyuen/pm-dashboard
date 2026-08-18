from flask import Flask, jsonify, render_template, request
import pandas as pd
from datetime import date
import os

app = Flask(__name__)
EXCEL_PATH = "PM_Data.xlsx"

def load_and_process_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name="Machines")
    df["Last_PM_Date"] = pd.to_datetime(df["Last_PM_Date"]).dt.date
    today = date.today()
    
    def calc_days_remaining(row):
        days_since_last = (today - row["Last_PM_Date"]).days
        days_remaining = row["PM_Interval_Days"] - days_since_last
        return days_remaining

    df["Days_Remaining"] = df.apply(calc_days_remaining, axis=1)
    
    def classify(days_remaining):
        if days_remaining < 0:
            return "Overdue"
        elif days_remaining <= 7:
            return "Due Soon"
        else:
            return "OK"
            
    df["PM_Status"] = df["Days_Remaining"].apply(classify)
    df["Last_PM_Date"] = df["Last_PM_Date"].astype(str)
    return df

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/machines")
def api_machines():
    df = load_and_process_data()
    return jsonify(df.to_dict(orient="records"))

# API สำหรับรับค่าวันที่ใหม่และบันทึกลง Excel
@app.route("/api/update_pm", methods=["POST"])
def update_pm():
    try:
        data = request.json
        machine_id = data.get("machine_id")
        new_date = data.get("new_date")
        
        df = pd.read_excel(EXCEL_PATH, sheet_name="Machines")
        
        if machine_id in df["Machine_ID"].values:
            df.loc[df["Machine_ID"] == machine_id, "Last_PM_Date"] = new_date
            
            with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Machines", index=False)
                
            return jsonify({"status": "success", "message": "อัปเดตข้อมูลสำเร็จ!"})
        else:
            return jsonify({"status": "error", "message": "ไม่พบ Machine ID"}), 400

    except PermissionError:
        return jsonify({"status": "error", "message": "กรุณาปิดไฟล์ Excel ก่อนกดบันทึก!"}), 500
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)