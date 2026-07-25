import os
import re
import math
import zipfile
import traceback
from datetime import datetime, date
from functools import wraps
from io import BytesIO
from urllib.parse import urlparse

from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

# ============================================
# 1. إعداد التطبيق
# ============================================

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'your-secret-key-here-change-this-in-production')

# ===== إعداد قاعدة البيانات =====
database_url = os.environ.get('DATABASE_URL')

if database_url:
    if database_url.startswith('postgres://'):
        database_url = database_url.replace('postgres://', 'postgresql://', 1)
    if '?' not in database_url:
        database_url += '?sslmode=require'
    else:
        database_url += '&sslmode=require'
    app.config['SQLALCHEMY_DATABASE_URI'] = database_url
    print('✅ Using PostgreSQL database')
else:
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///employees.db'
    print('ℹ️ Using SQLite database (local)')

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

db = SQLAlchemy(app)

# إنشاء المجلدات
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('static', exist_ok=True)


# ============================================
# 2. نماذج قاعدة البيانات (يجب أن تكون قبل db.create_all)
# ============================================

class User(db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    full_name = db.Column(db.String(100))
    is_admin = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<User {self.username}>'


class Employee(db.Model):
    __tablename__ = 'employees'

    id = db.Column(db.Integer, primary_key=True)

    # معلومات أساسية
    payroll_name = db.Column(db.String(100))
    sl_no = db.Column(db.Integer)
    emp_no = db.Column(db.String(50), nullable=False)
    name = db.Column(db.String(200), nullable=False)
    date_of_birth = db.Column(db.String(50))
    hire_date = db.Column(db.String(50))
    end_employment_date = db.Column(db.String(50))
    iqama_number = db.Column(db.String(50))
    gender = db.Column(db.String(20))
    nationality = db.Column(db.String(50))
    designation = db.Column(db.String(100))
    project_name = db.Column(db.String(200))
    sap_id = db.Column(db.String(50))
    location = db.Column(db.String(100))
    sponsor_name = db.Column(db.String(100))
    iban = db.Column(db.String(100))
    bank_name = db.Column(db.String(100))
    division = db.Column(db.String(50))

    # تفاصيل الراتب
    basic_salary = db.Column(db.Float, default=0)
    basic_salary_arrears = db.Column(db.Float, default=0)
    food_allowance = db.Column(db.Float, default=0)
    housing_allowance = db.Column(db.Float, default=0)
    housing_allowance_arrears = db.Column(db.Float, default=0)
    mobile_allowance = db.Column(db.Float, default=0)
    mobile_allowance_arrears = db.Column(db.Float, default=0)
    other_allowance = db.Column(db.Float, default=0)
    other_allowance_arrears = db.Column(db.Float, default=0)
    overtime = db.Column(db.Float, default=0)
    project_fix_allowance = db.Column(db.Float, default=0)
    project_fix_allowance_arrears = db.Column(db.Float, default=0)
    salary_adjustment = db.Column(db.Float, default=0)
    leave_salary_benefits = db.Column(db.Float, default=0)
    transportation_allowance = db.Column(db.Float, default=0)
    transportation_allowance_arrears = db.Column(db.Float, default=0)

    # الخصومات والإجماليات
    total_gross = db.Column(db.Float, default=0)
    absenteeism = db.Column(db.Float, default=0)
    advance_payment_recovery = db.Column(db.Float, default=0)
    gosi_contribution = db.Column(db.Float, default=0)
    other_deductions = db.Column(db.Float, default=0)
    total_deductions = db.Column(db.Float, default=0)
    net_pay = db.Column(db.Float, default=0)

    # معلومات إضافية
    status = db.Column(db.String(50))
    pay_group = db.Column(db.String(50))
    invoice_reference = db.Column(db.String(100))
    salary_released_on = db.Column(db.String(50))

    # الشهر والسنة
    month = db.Column(db.String(20))
    year = db.Column(db.Integer)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class UploadedFile(db.Model):
    __tablename__ = 'uploaded_files'

    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200), nullable=False)
    original_filename = db.Column(db.String(200), nullable=False)
    month = db.Column(db.String(20), nullable=False)
    year = db.Column(db.Integer, nullable=False)
    file_size = db.Column(db.Integer, default=0)
    upload_date = db.Column(db.DateTime, default=datetime.utcnow)
    employee_count = db.Column(db.Integer, default=0)
    total_salary = db.Column(db.Float, default=0)
    file_path = db.Column(db.String(500))
    uploaded_by = db.Column(db.Integer, db.ForeignKey('users.id'))

    user = db.relationship('User', backref=db.backref('uploads', lazy=True))


# ============================================
# 3. إنشاء الجداول والمستخدم المدير (بعد تعريف النماذج)
# ============================================

with app.app_context():
    try:
        db.create_all()
        print('✅ Database tables created/verified')

        # إنشاء المستخدم المدير
        admin = User.query.filter_by(username='admin').first()
        if not admin:
            admin = User(
                username='admin',
                email='admin@example.com',
                full_name='System Administrator',
                is_admin=True
            )
            admin.set_password('admin123')
            db.session.add(admin)
            db.session.commit()
            print('=' * 50)
            print('✅ Admin user created successfully!')
            print('   👤 Username: admin')
            print('   🔑 Password: admin123')
            print('=' * 50)
        else:
            print('ℹ️ Admin user already exists')
    except Exception as e:
        print(f'⚠️ Error initializing database: {e}')
        traceback.print_exc()


# ============================================
# 4. دوال المصادقة (Authentication)
# ============================================

def login_required(f):
    """Decorator لحماية الصفحات التي تتطلب تسجيل دخول"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('⚠️ Please login to access this page', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    """Decorator للصفحات التي تتطلب صلاحيات مدير"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('⚠️ Please login to access this page', 'warning')
            return redirect(url_for('login'))
        user = User.query.get(session['user_id'])
        if not user or not user.is_admin:
            flash('⛔ Access denied. Admin privileges required.', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function


# ============================================
# 5. دوال مساعدة
# ============================================

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


def safe_float(val):
    """تحويل القيمة إلى float بأمان"""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        try:
            cleaned = re.sub(r'[^0-9.]', '', val)
            return float(cleaned) if cleaned else 0.0
        except:
            return 0.0
    return 0.0


def safe_int(val):
    """تحويل القيمة إلى عدد صحيح"""
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return int(float(val))
    if isinstance(val, str):
        try:
            cleaned = re.sub(r'[^0-9.]', '', val)
            return int(float(cleaned)) if cleaned else 0
        except:
            return 0
    return 0


def safe_str(val):
    """تحويل القيمة إلى string بأمان"""
    if pd.isna(val):
        return ''
    return str(val).strip()


def calculate_service_period(hire_date_str, end_date_str=None):
    """حساب مدة الخدمة بالأيام والشهور والسنوات"""
    try:
        date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', str(hire_date_str))
        if not date_match:
            return '---'

        hire_date = date(int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3)))

        if end_date_str and str(end_date_str) != 'nan' and str(end_date_str) != '---':
            date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', str(end_date_str))
            if date_match:
                end_date = date(int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3)))
            else:
                end_date = date.today()
        else:
            end_date = date.today()

        delta = end_date - hire_date
        years = delta.days // 365
        remaining_days = delta.days % 365
        months = remaining_days // 30
        days = remaining_days % 30

        parts = []
        if years > 0:
            parts.append(f"{years} Year{'s' if years > 1 else ''}")
        if months > 0:
            parts.append(f"{months} Month{'s' if months > 1 else ''}")
        if days > 0 or not parts:
            parts.append(f"{days} Day{'s' if days != 1 else ''}")

        return " ".join(parts)

    except Exception:
        return '---'


# ============================================
# 6. إنشاء الـ Slip
# ============================================

def create_professional_slip(employee):
    """إنشاء Slip احترافي مع شعار الشركة"""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    # الشعار
    logo_path = os.path.join('static', 'logo.png')
    logo_para = doc.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if os.path.exists(logo_path):
        try:
            run = logo_para.add_run()
            run.add_picture(logo_path, width=Inches(2.5))
        except:
            logo_para.add_run("🏢").font.size = Pt(48)
    else:
        logo_para.add_run("🏢").font.size = Pt(48)

    doc.add_paragraph()

    # العنوان
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run("STATEMENT OF EARNINGS")
    title_run.bold = True
    title_run.font.size = Pt(22)
    title_run.font.color.rgb = RGBColor(0, 51, 102)

    month_para = doc.add_paragraph()
    month_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    month_display = employee.month or "June"
    year_display = employee.year or 2026
    month_run = month_para.add_run(f"Pay slip for the Month of: {month_display} {year_display}")
    month_run.font.size = Pt(12)
    month_run.italic = True

    doc.add_paragraph("_" * 80)

    # معلومات الموظف
    emp_no_clean = str(employee.emp_no).split('.')[0] if employee.emp_no else ''
    iqama_clean = str(employee.iqama_number).split('.')[0] if employee.iqama_number else '---'
    sap_clean = str(employee.sap_id).split('.')[0] if employee.sap_id else '---'

    hire_date_clean = employee.hire_date or '---'
    if ' ' in hire_date_clean:
        hire_date_clean = hire_date_clean.split(' ')[0]
    elif 'T' in hire_date_clean:
        hire_date_clean = hire_date_clean.split('T')[0]

    end_date_clean = employee.end_employment_date or '---'
    if ' ' in end_date_clean:
        end_date_clean = end_date_clean.split(' ')[0]
    elif 'T' in end_date_clean:
        end_date_clean = end_date_clean.split('T')[0]

    service_period = calculate_service_period(employee.hire_date, employee.end_employment_date)

    info_table = doc.add_table(rows=8, cols=2)
    info_table.style = 'Table Grid'
    info_table.columns[0].width = Inches(3)
    info_table.columns[1].width = Inches(4)

    for row in info_table.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    info_data = [
        ("اﺳﻢ اﻟﻤﻮﻇﻒ / Employee Name", employee.name),
        ("رﻗﻢ اﻟﻤﻮﻇﻒ / Employee Number", emp_no_clean),
        ("رﻗﻢ اﻹﻗﺎﻣﺔ / Iqama Number", iqama_clean),
        ("SAP ID", sap_clean),
        ("اﻟﻤﺴﻤﻰ اﻟﻮﻇﻴﻔﻲ / Job Title", employee.designation or '---'),
        ("اﻟﺠﻨﺴﻴﺔ / Nationality", employee.nationality or '---'),
        ("اسم المشروع / Project Name", employee.project_name or '---'),
        ("اﻟﻤﻮﻗﻊ / Location", employee.location or '---'),
    ]

    for i, (label, value) in enumerate(info_data):
        info_table.cell(i, 0).text = label
        info_table.cell(i, 1).text = str(value)

    doc.add_paragraph()

    # جدول التواريخ
    date_table = doc.add_table(rows=2, cols=3)
    date_table.style = 'Table Grid'
    date_table.columns[0].width = Inches(2.5)
    date_table.columns[1].width = Inches(2.5)
    date_table.columns[2].width = Inches(2.5)

    headers = ["ﺗﺎرﻳﺦ اﻟﺘﻌﻴﻴﻦ / Hire Date", "ﺗﺎرﻳﺦ اﻧﺘﻬﺎء اﻟﺘﻮﻇﻴﻒ / End Date", "ﻣﺪة اﻟﺨﺪﻣﺔ / Service Period"]
    for i, header in enumerate(headers):
        cell = date_table.cell(0, i)
        cell.text = header
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(0, 51, 102)

    date_table.cell(1, 0).text = hire_date_clean
    date_table.cell(1, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_table.cell(1, 1).text = end_date_clean
    date_table.cell(1, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_table.cell(1, 2).text = service_period
    date_table.cell(1, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in date_table.cell(1, 2).paragraphs[0].runs:
        run.bold = True
        run.font.color.rgb = RGBColor(0, 128, 0)

    doc.add_paragraph("_" * 80)

    # جدول الراتب
    salary_table = doc.add_table(rows=7, cols=4)
    salary_table.style = 'Table Grid'
    salary_table.columns[0].width = Inches(2.5)
    salary_table.columns[1].width = Inches(2)
    salary_table.columns[2].width = Inches(2.5)
    salary_table.columns[3].width = Inches(2)

    headers = ["Earnings / الإضافات", "Amount / المبلغ", "Deductions / الخصومات", "Amount / المبلغ"]
    for i, header in enumerate(headers):
        cell = salary_table.cell(0, i)
        cell.text = header
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(0, 51, 102)

    salary_rows = [
        ("Basic Salary / الراتب الأساسي", employee.basic_salary, "Salary Adjustment / تسوية الراتب",
         employee.salary_adjustment),
        ("Housing Allowance / بدل سكن", employee.housing_allowance, "", ""),
        ("Transportation Allowance / بدل النقل", employee.transportation_allowance, "", ""),
        ("Other Allowance / بدلات أخرى", employee.other_allowance, "", ""),
        ("", "", "", ""),
        ("Total Earnings / الأرباح الكلية", employee.total_gross, "Total Deductions / إجمالي الخصومات",
         employee.total_deductions),
    ]

    for i, row_data in enumerate(salary_rows):
        row_idx = i + 1
        for j, value in enumerate(row_data):
            if j == 0 or j == 2:
                salary_table.cell(row_idx, j).text = str(value)
                salary_table.cell(row_idx, j).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                salary_table.cell(row_idx, j).text = f"{value:,.2f}" if value else "0.00"
                salary_table.cell(row_idx, j).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("_" * 80)

    # صافي الراتب
    net_table = doc.add_table(rows=1, cols=2)
    net_table.style = 'Table Grid'
    net_table.columns[0].width = Inches(4)
    net_table.columns[1].width = Inches(3)

    net_cell = net_table.cell(0, 0)
    net_cell.text = "Net Salary / صافي الراتب"
    net_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in net_cell.paragraphs[0].runs:
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0, 51, 102)

    amount_cell = net_table.cell(0, 1)
    amount_cell.text = f"{employee.net_pay:,.2f}"
    amount_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in amount_cell.paragraphs[0].runs:
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0, 128, 0)

    doc.add_paragraph("_" * 80)

    # البنك
    bank_table = doc.add_table(rows=1, cols=4)
    bank_table.style = 'Table Grid'

    bank_data = [
        ("Bank Name / اسم البنك", employee.bank_name or '---', "Account Number / رقم الحساب", employee.iban or '---')
    ]

    for i, row_data in enumerate(bank_data):
        for j, value in enumerate(row_data):
            bank_table.cell(i, j).text = str(value)
            bank_table.cell(i, j).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("_" * 80)

    # إخلاء المسؤولية
    disclaimer_para = doc.add_paragraph()
    disclaimer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    disclaimer_text = (
        "This document is automatically generated by the system and does not require signatures.\n"
        "هذا المستند تم توليده تلقائياً من النظام ولا يحتاج إلى توقيعات."
    )

    disclaimer_run = disclaimer_para.add_run(disclaimer_text)
    disclaimer_run.font.size = Pt(11)
    disclaimer_run.italic = True
    disclaimer_run.font.color.rgb = RGBColor(100, 100, 100)

    doc.add_paragraph("_" * 80)

    # التذييل
    footer_para = doc.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_para.add_run(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    footer_run.font.size = Pt(8)
    footer_run.italic = True

    return doc


# ============================================
# 7. Routes - المصادقة (Authentication)
# ============================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    """صفحة تسجيل الدخول"""
    if 'user_id' in session:
        return redirect(url_for('index'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        if not username or not password:
            flash('⚠️ Please enter both username and password', 'warning')
            return render_template('login.html')

        user = User.query.filter_by(username=username).first()

        if not user:
            flash('❌ Invalid username or password', 'error')
            return render_template('login.html')

        if user.check_password(password):
            session['user_id'] = user.id
            session['username'] = user.username
            session['full_name'] = user.full_name
            session['is_admin'] = user.is_admin

            user.last_login = datetime.utcnow()
            db.session.commit()

            flash(f'✅ Welcome back, {user.full_name or user.username}!', 'success')

            next_page = request.args.get('next')
            if next_page:
                return redirect(next_page)
            return redirect(url_for('index'))
        else:
            flash('❌ Invalid username or password', 'error')

    return render_template('login.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    """صفحة تسجيل حساب جديد"""
    if 'user_id' in session:
        return redirect(url_for('index'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        email = request.form.get('email', '').strip()
        full_name = request.form.get('full_name', '').strip()
        password = request.form.get('password', '')
        confirm_password = request.form.get('confirm_password', '')

        if not username or not email or not password:
            flash('⚠️ Please fill in all required fields', 'warning')
            return render_template('register.html')

        if len(username) < 3:
            flash('⚠️ Username must be at least 3 characters', 'warning')
            return render_template('register.html')

        if len(password) < 6:
            flash('⚠️ Password must be at least 6 characters', 'warning')
            return render_template('register.html')

        if password != confirm_password:
            flash('⚠️ Passwords do not match', 'warning')
            return render_template('register.html')

        existing_user = User.query.filter_by(username=username).first()
        if existing_user:
            flash('❌ Username already exists', 'error')
            return render_template('register.html')

        existing_email = User.query.filter_by(email=email).first()
        if existing_email:
            flash('❌ Email already registered', 'error')
            return render_template('register.html')

        user = User(
            username=username,
            email=email,
            full_name=full_name or username,
            is_admin=False
        )
        user.set_password(password)

        db.session.add(user)
        db.session.commit()

        flash('✅ Account created successfully! Please login.', 'success')
        return redirect(url_for('login'))

    return render_template('register.html')


@app.route('/logout')
def logout():
    """تسجيل الخروج"""
    session.clear()
    flash('👋 You have been logged out', 'success')
    return redirect(url_for('login'))


@app.route('/profile')
@login_required
def profile():
    """صفحة الملف الشخصي"""
    user = User.query.get(session['user_id'])
    return render_template('profile.html', user=user)


# ============================================
# 8. Routes - الصفحات الرئيسية (محمية)
# ============================================

@app.route('/')
@login_required
def index():
    """الصفحة الرئيسية"""
    return render_template('index.html')


@app.route('/upload', methods=['GET', 'POST'])
@login_required
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('❌ No file uploaded', 'error')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('❌ No file selected', 'error')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            unique_filename = f"{timestamp}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(filepath)

            try:
                df = pd.read_excel(filepath, sheet_name='Payroll')

                month = request.form.get('month', 'June')
                year = int(request.form.get('year', datetime.now().year))

                count = 0
                total_salary = 0

                for _, row in df.iterrows():
                    if pd.isna(row.get('Emp No.')) or pd.isna(row.get('Name')):
                        continue

                    net_pay = safe_float(row.get('Net Pay'))
                    total_salary += net_pay

                    employee = Employee(
                        payroll_name=safe_str(row.get('Payroll Name')),
                        sl_no=safe_int(row.get('SL #')),
                        emp_no=str(safe_int(row.get('Emp No.'))),
                        name=safe_str(row.get('Name')),
                        date_of_birth=safe_str(row.get('Date of Birth')),
                        hire_date=safe_str(row.get('Hire Date')),
                        end_employment_date=safe_str(row.get('End Employment Date')),
                        iqama_number=str(safe_int(row.get('Iqama Number'))),
                        gender=safe_str(row.get('Gender')),
                        nationality=safe_str(row.get('Nationality')),
                        designation=safe_str(row.get('Designation')),
                        project_name=safe_str(row.get('Project Name')),
                        sap_id=str(safe_int(row.get('SAP ID'))),
                        location=safe_str(row.get('Location')),
                        sponsor_name=safe_str(row.get('Sponsor Name')),
                        iban=safe_str(row.get('IBAN')),
                        bank_name=safe_str(row.get('BANK NAME')),
                        division=safe_str(row.get('Division')),
                        basic_salary=safe_float(row.get('Basic Salary')),
                        basic_salary_arrears=safe_float(row.get('Basic Salary Arrears')),
                        food_allowance=safe_float(row.get('Food Allowance')),
                        housing_allowance=safe_float(row.get('Housing Allowance')),
                        housing_allowance_arrears=safe_float(row.get('Housing Allowance Arrears')),
                        mobile_allowance=safe_float(row.get('Mobile Allowance')),
                        mobile_allowance_arrears=safe_float(row.get('Mobile Allowance Arrears')),
                        other_allowance=safe_float(row.get('Other Allowance')),
                        other_allowance_arrears=safe_float(row.get('Other Allowance Arrears')),
                        overtime=safe_float(row.get('Overtime')),
                        project_fix_allowance=safe_float(row.get('Project Fix Allowance')),
                        project_fix_allowance_arrears=safe_float(row.get('Project Fix Allowance Arrears')),
                        salary_adjustment=safe_float(row.get('Salary Adjustment')),
                        leave_salary_benefits=safe_float(row.get('The Leave Salary Benefits')),
                        transportation_allowance=safe_float(row.get('Transportation Allowance')),
                        transportation_allowance_arrears=safe_float(row.get('Transportation Allowance Arrears')),
                        total_gross=safe_float(row.get('Total Gross')),
                        absenteeism=safe_float(row.get('Absenteeism')),
                        advance_payment_recovery=safe_float(row.get('Advance Payment Recovery')),
                        gosi_contribution=safe_float(row.get('GOSI EE Contribution')),
                        other_deductions=safe_float(row.get('Other Deductions')),
                        total_deductions=safe_float(row.get('Total Deductions')),
                        net_pay=net_pay,
                        status=safe_str(row.get('Status')),
                        pay_group=safe_str(row.get('Pay Group')),
                        invoice_reference=safe_str(row.get('Invoice Reference')),
                        salary_released_on=safe_str(row.get('Salary released on')),
                        month=month,
                        year=year
                    )
                    db.session.add(employee)
                    count += 1

                db.session.commit()

                file_size = os.path.getsize(filepath)
                uploaded_file = UploadedFile(
                    filename=unique_filename,
                    original_filename=filename,
                    month=month,
                    year=year,
                    file_size=file_size,
                    employee_count=count,
                    total_salary=total_salary,
                    file_path=filepath,
                    uploaded_by=session['user_id']
                )
                db.session.add(uploaded_file)
                db.session.commit()

                flash(f'✅ Successfully uploaded {count} employees for {month} {year}', 'success')
                return redirect(url_for('uploaded_files'))

            except Exception as e:
                db.session.rollback()
                if os.path.exists(filepath):
                    os.remove(filepath)
                flash(f'❌ Error uploading file: {str(e)}', 'error')
                return redirect(request.url)

        flash('❌ File type not supported. Please upload an Excel file (.xlsx, .xls)', 'error')
        return redirect(request.url)

    return render_template('upload.html')


@app.route('/uploaded-files')
@login_required
def uploaded_files():
    """عرض جميع الملفات المرفوعة"""
    files = UploadedFile.query.order_by(UploadedFile.upload_date.desc()).all()
    return render_template('uploaded_files.html', files=files)


@app.route('/download-file/<int:file_id>')
@login_required
def download_file(file_id):
    """تحميل ملف Excel"""
    uploaded_file = UploadedFile.query.get_or_404(file_id)
    file_path = uploaded_file.file_path

    if not os.path.exists(file_path):
        flash('❌ File not found', 'error')
        return redirect(url_for('uploaded_files'))

    return send_file(
        file_path,
        as_attachment=True,
        download_name=uploaded_file.original_filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/file-stats/<int:file_id>')
@login_required
def file_stats(file_id):
    """عرض إحصائيات ملف معين"""
    uploaded_file = UploadedFile.query.get_or_404(file_id)

    employees = Employee.query.filter_by(
        month=uploaded_file.month,
        year=uploaded_file.year
    ).all()

    employees_json = []
    for emp in employees:
        employees_json.append({
            'id': emp.id,
            'emp_no': emp.emp_no,
            'name': emp.name,
            'division': emp.division or 'No Division',
            'project_name': emp.project_name or 'No Client',
            'designation': emp.designation or '-',
            'net_pay': float(emp.net_pay or 0),
            'month': emp.month,
            'year': emp.year
        })

    stats = {
        'total_employees': len(employees),
        'total_salary': sum(emp.net_pay for emp in employees),
        'avg_salary': sum(emp.net_pay for emp in employees) / len(employees) if employees else 0,
        'max_salary': max((emp.net_pay for emp in employees), default=0),
        'min_salary': min((emp.net_pay for emp in employees), default=0),
        'divisions': len(set(emp.division for emp in employees if emp.division)),
        'clients': len(set(emp.project_name for emp in employees if emp.project_name)),
        'month': uploaded_file.month,
        'year': uploaded_file.year,
        'filename': uploaded_file.original_filename,
        'upload_date': uploaded_file.upload_date,
        'file_id': uploaded_file.id
    }

    return render_template(
        'file_stats.html',
        stats=stats,
        employees=employees,
        employees_json=employees_json
    )


@app.route('/delete-file/<int:file_id>')
@login_required
def delete_file(file_id):
    """حذف ملف مرفوع وبياناته"""
    uploaded_file = UploadedFile.query.get_or_404(file_id)

    try:
        if os.path.exists(uploaded_file.file_path):
            os.remove(uploaded_file.file_path)

        Employee.query.filter_by(
            month=uploaded_file.month,
            year=uploaded_file.year
        ).delete()

        db.session.delete(uploaded_file)
        db.session.commit()

        flash(f'✅ Deleted file: {uploaded_file.original_filename}', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'❌ Error deleting file: {str(e)}', 'error')

    return redirect(url_for('uploaded_files'))


@app.route('/employees')
@login_required
def view_employees():
    """عرض الموظفين مع فلتر و Pagination"""
    page = request.args.get('page', 1, type=int)
    search_query = request.args.get('q', '')
    filter_division = request.args.get('filter_division', '')
    filter_client = request.args.get('filter_client', '')
    per_page = 15

    query = Employee.query

    if search_query:
        search_term = f"%{search_query}%"
        query = query.filter(
            db.or_(
                Employee.name.like(search_term),
                Employee.emp_no.like(search_term),
                Employee.designation.like(search_term)
            )
        )

    if filter_division:
        query = query.filter(Employee.division == filter_division)

    if filter_client:
        query = query.filter(Employee.project_name == filter_client)

    query = query.order_by(Employee.emp_no)

    all_divisions = db.session.query(Employee.division).distinct().order_by(Employee.division).all()
    all_clients = db.session.query(Employee.project_name).distinct().order_by(Employee.project_name).all()

    divisions = [d[0] for d in all_divisions if d[0]]
    clients = [c[0] for c in all_clients if c[0]]

    total = query.count()
    all_employees = query.all()
    total_net_pay_all = sum(emp.net_pay for emp in all_employees)

    employees = query.offset((page - 1) * per_page).limit(per_page).all()
    total_pages = math.ceil(total / per_page)

    return render_template(
        'employees.html',
        employees=employees,
        page=page,
        total_pages=total_pages,
        total=total,
        total_net_pay_all=total_net_pay_all,
        search_query=search_query,
        filter_division=filter_division,
        filter_client=filter_client,
        divisions=divisions,
        clients=clients
    )


@app.route('/employee/<int:emp_id>')
@login_required
def employee_detail(emp_id):
    """عرض تفاصيل الموظف"""
    employee = Employee.query.get_or_404(emp_id)
    return render_template('employee_detail.html', employee=employee)


@app.route('/generate_slip/<int:emp_id>')
@login_required
def generate_slip(emp_id):
    """إنشاء وتحميل Slip فردي"""
    employee = Employee.query.get_or_404(emp_id)

    doc = create_professional_slip(employee)

    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    filename = f"slip_{employee.emp_no}_{employee.month}_{employee.year}.docx"

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@app.route('/all_slips')
@login_required
def all_slips():
    """عرض جميع الـ Slips مع فلاتر و Pagination"""
    filter_division = request.args.get('division', '')
    filter_client = request.args.get('client', '')
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 12)

    query = Employee.query

    if filter_division:
        query = query.filter(Employee.division == filter_division)

    if filter_client:
        query = query.filter(Employee.project_name == filter_client)

    query = query.order_by(Employee.emp_no)

    total = query.count()
    all_employees_for_sum = query.all()
    total_net_pay_all = sum(emp.net_pay for emp in all_employees_for_sum)

    if per_page == 'all':
        employees = all_employees_for_sum
        total_pages = 1
        page = 1
    else:
        per_page_int = int(per_page)
        employees = query.offset((page - 1) * per_page_int).limit(per_page_int).all()
        total_pages = math.ceil(total / per_page_int)

    all_divisions = db.session.query(Employee.division).distinct().order_by(Employee.division).all()
    all_clients = db.session.query(Employee.project_name).distinct().order_by(Employee.project_name).all()

    divisions = [d[0] for d in all_divisions if d[0]]
    clients = [c[0] for c in all_clients if c[0]]

    return render_template(
        'all_slips.html',
        employees=employees,
        divisions=divisions,
        clients=clients,
        filter_division=filter_division,
        filter_client=filter_client,
        page=page,
        total_pages=total_pages,
        total=total,
        total_net_pay_all=total_net_pay_all,
        per_page=per_page
    )


@app.route('/generate_selected_slips', methods=['POST'])
@login_required
def generate_selected_slips():
    """إنشاء Slips للموظفين المحددين (ملف ZIP)"""
    employee_ids = request.form.getlist('employee_ids')

    if not employee_ids:
        flash('⚠️ No employees selected', 'warning')
        return redirect(url_for('all_slips'))

    try:
        ids = [int(id) for id in employee_ids if id.isdigit()]
    except ValueError:
        flash('⚠️ Invalid employee IDs', 'warning')
        return redirect(url_for('all_slips'))

    if not ids:
        flash('⚠️ Invalid employee IDs', 'warning')
        return redirect(url_for('all_slips'))

    employees = Employee.query.filter(Employee.id.in_(ids)).all()

    if not employees:
        flash('⚠️ No employees found', 'warning')
        return redirect(url_for('all_slips'))

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for emp in employees:
            doc = create_professional_slip(emp)
            doc_buffer = BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)

            emp_no_clean = str(emp.emp_no).replace('/', '_').replace('\\', '_')
            filename = f"slip_{emp_no_clean}_{emp.month}_{emp.year}.docx"
            zip_file.writestr(filename, doc_buffer.getvalue())

    zip_buffer.seek(0)

    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=f'slips_{datetime.now().strftime("%Y%m%d_%H%M")}.zip',
        mimetype='application/zip'
    )


@app.route('/delete_all')
@admin_required
def delete_all():
    """حذف جميع الموظفين"""
    try:
        count = db.session.query(Employee).delete()
        db.session.commit()
        flash(f'✅ Deleted {count} employees', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'❌ Error: {str(e)}', 'error')
    return redirect(url_for('view_employees'))


@app.route('/search')
@login_required
def search_employees():
    """البحث عن الموظفين"""
    search_query = request.args.get('q', '')
    page = request.args.get('page', 1, type=int)
    per_page = 15

    query = Employee.query

    if search_query:
        search_term = f"%{search_query}%"
        query = query.filter(
            db.or_(
                Employee.division.like(search_term),
                Employee.project_name.like(search_term)
            )
        )

    query = query.order_by(Employee.emp_no)

    total = query.count()
    all_employees = query.all()
    total_net_pay_all = sum(e.net_pay for e in all_employees)

    employees = query.offset((page - 1) * per_page).limit(per_page).all()
    total_pages = math.ceil(total / per_page)

    return render_template(
        'employees.html',
        employees=employees,
        page=page,
        total_pages=total_pages,
        total=total,
        total_net_pay_all=total_net_pay_all,
        search_query=search_query
    )


# ============================================
# 9. تشغيل التطبيق
# ============================================

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True, host='0.0.0.0', port=5000)