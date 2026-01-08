"""
Center of Excellence - Tinnitus and Vestibular Disorders
Patient Evaluation System with Admin Protection
Using MongoDB Atlas for Online Cloud Storage
"""

from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
from functools import wraps
from pymongo import MongoClient
from bson.objectid import ObjectId
from datetime import datetime
import hashlib
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

app = Flask(__name__)
app.secret_key = 'tinnitus_vestibular_center_secure_2024_key'

# ============ MONGODB CONFIGURATION ============
# IMPORTANT: Replace with your MongoDB Atlas connection string
# Go to https://cloud.mongodb.com to create a FREE cluster
# 1. Create account -> Create Free Cluster -> Connect -> Connect your application
# 2. Copy the connection string and paste below
# 3. Replace <password> with your actual password

MONGODB_URI = "mongodb+srv://aiish:4fHsqALLOIzrUq08@cluster0.vzjwbg6.mongodb.net/" # Local MongoDB (for testing)
# For MongoDB Atlas (online), use something like:
# MONGODB_URI = "mongodb+srv://username:password@cluster0.xxxxx.mongodb.net/?retryWrites=true&w=majority"

DATABASE_NAME = "aiish"

# Admin credentials
ADMIN_USERNAME = 'admin'
ADMIN_PASSWORD_HASH = hashlib.sha256('admin123'.encode()).hexdigest()

def get_db():
    """Get MongoDB database connection"""
    client = MongoClient(MONGODB_URI)
    return client[DATABASE_NAME]

def init_db():
    """Initialize database with admin user"""
    db = get_db()
    
    # Create indexes for better performance
    db.patients.create_index("case_number", unique=True)
    db.admins.create_index("username", unique=True)
    
    # Insert default admin if not exists
    existing_admin = db.admins.find_one({"username": ADMIN_USERNAME})
    if not existing_admin:
        db.admins.insert_one({
            "username": ADMIN_USERNAME,
            "password_hash": ADMIN_PASSWORD_HASH,
            "created_at": datetime.now()
        })
        print("Default admin created.")

try:
    init_db()
    print("✓ MongoDB Connected Successfully!")
except Exception as e:
    print(f"⚠ MongoDB Connection Error: {e}")
    print("Make sure MongoDB is running or update MONGODB_URI with your Atlas connection string.")

# Admin authentication decorator
def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('admin_logged_in'):
            flash('Admin login required for this action.', 'error')
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/')
def index():
    return redirect(url_for('new_patient'))

# ============ ADMIN ROUTES ============

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        
        db = get_db()
        admin = db.admins.find_one({
            "username": username,
            "password_hash": password_hash
        })
        
        if admin:
            session['admin_logged_in'] = True
            session['admin_username'] = username
            flash('Logged in successfully!', 'success')
            return redirect(url_for('admin_dashboard'))
        else:
            flash('Invalid username or password!', 'error')
    
    return render_template('admin_login.html')

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_logged_in', None)
    session.pop('admin_username', None)
    flash('Logged out successfully!', 'success')
    return redirect(url_for('index'))

@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    db = get_db()
    patients = list(db.patients.find().sort("created_at", -1))
    total_patients = len(patients)
    return render_template('admin_dashboard.html', patients=patients, total=total_patients)

@app.route('/admin/export-excel')
@admin_required
def export_excel():
    """Export all patient records to Excel"""
    db = get_db()
    patients = list(db.patients.find().sort("created_at", -1))
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Patient Records"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Headers
    headers = [
        "S.No", "Case Number", "Name", "Age", "Gender", "Mobile Number",
        "Date of Evaluation", "Brief History", "Tests Conducted", "Other Tests",
        "Instruments Used", "Other Instruments", "Provisional Diagnosis", "Created At"
    ]
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Data rows
    for row_num, patient in enumerate(patients, 2):
        ws.cell(row=row_num, column=1, value=row_num-1).border = thin_border
        ws.cell(row=row_num, column=2, value=patient.get('case_number', '')).border = thin_border
        ws.cell(row=row_num, column=3, value=patient.get('name', '')).border = thin_border
        ws.cell(row=row_num, column=4, value=patient.get('age', '')).border = thin_border
        ws.cell(row=row_num, column=5, value=patient.get('gender', '')).border = thin_border
        ws.cell(row=row_num, column=6, value=patient.get('mobile_number', '')).border = thin_border
        ws.cell(row=row_num, column=7, value=patient.get('date_of_evaluation', '')).border = thin_border
        ws.cell(row=row_num, column=8, value=patient.get('brief_history', '')).border = thin_border
        ws.cell(row=row_num, column=9, value=patient.get('tests_conducted', '')).border = thin_border
        ws.cell(row=row_num, column=10, value=patient.get('other_tests', '')).border = thin_border
        ws.cell(row=row_num, column=11, value=patient.get('instruments_used', '')).border = thin_border
        ws.cell(row=row_num, column=12, value=patient.get('other_instruments', '')).border = thin_border
        ws.cell(row=row_num, column=13, value=patient.get('provisional_diagnosis', '')).border = thin_border
        created_at = patient.get('created_at', '')
        if created_at:
            created_at = created_at.strftime('%Y-%m-%d %H:%M:%S') if hasattr(created_at, 'strftime') else str(created_at)
        ws.cell(row=row_num, column=14, value=created_at).border = thin_border
    
    # Adjust column widths
    column_widths = [6, 15, 20, 6, 10, 15, 15, 30, 40, 20, 40, 20, 30, 20]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64+i) if i <= 26 else 'A' + chr(64+i-26)].width = width
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    # Generate filename with date
    filename = f"patient_records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

@app.route('/admin/change-password', methods=['GET', 'POST'])
@admin_required
def change_password():
    if request.method == 'POST':
        current_password = request.form['current_password']
        new_password = request.form['new_password']
        confirm_password = request.form['confirm_password']
        
        current_hash = hashlib.sha256(current_password.encode()).hexdigest()
        
        db = get_db()
        admin = db.admins.find_one({
            "username": session['admin_username'],
            "password_hash": current_hash
        })
        
        if not admin:
            flash('Current password is incorrect!', 'error')
        elif new_password != confirm_password:
            flash('New passwords do not match!', 'error')
        elif len(new_password) < 6:
            flash('Password must be at least 6 characters!', 'error')
        else:
            new_hash = hashlib.sha256(new_password.encode()).hexdigest()
            db.admins.update_one(
                {"username": session['admin_username']},
                {"$set": {"password_hash": new_hash}}
            )
            flash('Password changed successfully!', 'success')
            return redirect(url_for('admin_dashboard'))
    
    return render_template('change_password.html')

@app.route('/admin/delete/<id>', methods=['POST'])
@admin_required
def admin_delete_patient(id):
    db = get_db()
    patient = db.patients.find_one({"_id": ObjectId(id)})
    if patient:
        db.patients.delete_one({"_id": ObjectId(id)})
        flash(f'Patient record {patient["case_number"]} deleted by admin.', 'success')
    return redirect(url_for('admin_dashboard'))

# ============ PUBLIC ROUTES ============

@app.route('/new', methods=['GET', 'POST'])
def new_patient():
    if request.method == 'POST':
        # Get form data
        case_number = request.form['case_number']
        name = request.form['name']
        age = request.form['age']
        gender = request.form['gender']
        mobile_number = request.form['mobile_number']
        date_of_evaluation = request.form['date_of_evaluation']
        brief_history = request.form.get('brief_history', '')
        
        # Get tests conducted (checkboxes)
        tests_conducted = ','.join(request.form.getlist('tests_conducted'))
        other_tests = request.form.get('other_tests', '')
        
        # Get instruments used (checkboxes)
        instruments_used = ','.join(request.form.getlist('instruments_used'))
        other_instruments = request.form.get('other_instruments', '')
        
        provisional_diagnosis = request.form.get('provisional_diagnosis', '')
        
        try:
            db = get_db()
            
            # Check if case number already exists
            existing = db.patients.find_one({"case_number": case_number})
            if existing:
                flash('Case Number already exists!', 'error')
                return render_template('form.html')
            
            # Insert patient document
            patient_doc = {
                "case_number": case_number,
                "name": name,
                "age": int(age),
                "gender": gender,
                "mobile_number": mobile_number,
                "date_of_evaluation": date_of_evaluation,
                "brief_history": brief_history,
                "tests_conducted": tests_conducted,
                "other_tests": other_tests,
                "instruments_used": instruments_used,
                "other_instruments": other_instruments,
                "provisional_diagnosis": provisional_diagnosis,
                "created_at": datetime.now(),
                "created_by": "Staff"
            }
            
            db.patients.insert_one(patient_doc)
            flash('Patient record saved successfully! This record is now LOCKED and cannot be edited.', 'success')
            return redirect(url_for('view_patients'))
        except Exception as e:
            flash(f'Error saving record: {str(e)}', 'error')
    
    return render_template('form.html')

@app.route('/patients')
def view_patients():
    db = get_db()
    patients = list(db.patients.find().sort("created_at", -1))
    is_admin = session.get('admin_logged_in', False)
    return render_template('patients.html', patients=patients, is_admin=is_admin)

@app.route('/patient/<id>')
def view_patient(id):
    db = get_db()
    patient = db.patients.find_one({"_id": ObjectId(id)})
    is_admin = session.get('admin_logged_in', False)
    return render_template('patient_detail.html', patient=patient, is_admin=is_admin)

if __name__ == '__main__':
    print("=" * 50)
    print("Center of Excellence - Tinnitus & Vestibular Disorders")
    print("=" * 50)
    print("Access at: http://localhost:5000")
    print("\nDefault Admin Credentials:")
    print("  Username: admin")
    print("  Password: admin123")
    print("\nNote: Patient records CANNOT be edited after submission.")
    print("Only admins can delete records.")
    print("=" * 50)
    # For Render deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
