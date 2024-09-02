from flask import Flask, request, render_template, redirect
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

app = Flask(__name__)
EXCEL_FILE = 'employee_data.xlsx'

# Initialize Excel file if it doesn't exist
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.append(["Employee ID", "Date", "Check-In Time", "Check-Out Time"])  # Adding headers
    wb.save(EXCEL_FILE)

# Route for Check-In Page
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle Check-In form submission
@app.route('/checkin', methods=['POST'])
def checkin():
    employee_id = request.form['employee_id']
    hour = request.form['hour']
    minute = request.form['minute']
    ampm = request.form['ampm']
    checkin_time = f"{hour}:{minute} {ampm}"
    current_date = datetime.now().strftime("%Y-%m-%d")  # Save the current date

    # Save Check-In time to Excel
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([employee_id, current_date, checkin_time, ""])  # Leave check-out time empty for now
    wb.save(EXCEL_FILE)

    return redirect('/')

# Route to handle Check-Out form submission
@app.route('/checkout', methods=['POST'])
def checkout():
    employee_id = request.form['employee_id']
    hour = request.form['hour']
    minute = request.form['minute']
    ampm = request.form['ampm']
    checkout_time = f"{hour}:{minute} {ampm}"
    current_date = datetime.now().strftime("%Y-%m-%d")  # Save the current date

    # Save Check-Out time to Excel
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    # Find the latest entry for this employee ID and update the check-out time
    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[0].value == employee_id and row[1].value == current_date and row[3].value == "":
            row[3].value = checkout_time  # Update check-out time
            break

    wb.save(EXCEL_FILE)
    return redirect('/')

# Route for Status Page
@app.route('/status', methods=['GET', 'POST'])
def status():
    records = []
    if request.method == 'POST':
        employee_id = request.form['employee_id']
        
        # Read from Excel to find records for the given employee ID
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            if row[0] == employee_id:
                records.append({
                    "date": row[1],
                    "checkin": row[2],
                    "checkout": row[3] if row[3] else "Not Checked Out"
                })

    return render_template('status.html', records=records)

if __name__ == '__main__':
    app.run(debug=True)
