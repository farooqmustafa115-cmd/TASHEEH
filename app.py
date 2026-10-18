import os
import sys
import json
import base64
import io
import re
import sqlite3
import threading
import random
import ctypes
import unicodedata
import traceback
import secrets
import hashlib
import hmac
import math
import webview
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify, send_file

# دعم التصدير لمستند Word
try:
    import docx
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.text.paragraph import Paragraph as DocxParagraph
    from docx.table import Table as DocxTable
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

try:
    import qrcode
    HAS_QRCODE = True
except ImportError:
    HAS_QRCODE = False

try:
    import pymupdf as fitz
    import cv2
    import numpy as np
    HAS_SCAN_LIBS = True
except ImportError:
    HAS_SCAN_LIBS = False

# محرك OCR الاحتياطي (Tesseract) للتعرف على الرقم العسكري المطبوع عند تعذر قراءة QR
try:
    import pytesseract
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False

# محرك QR احتياطي مستقل يعزز قراءة الرموز من المسح الضوئي الرديء
try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    from PIL import Image as PILImage
    HAS_PYZBAR = True
except ImportError:
    HAS_PYZBAR = False

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 512 * 1024 * 1024

# مسارات التخزين بجانب ملف البرنامج أو الملف التنفيذي لتسهيل نقله معه
if getattr(sys, 'frozen', False):
  APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
  search_dir = APP_DIR
  while True:
    if os.path.isfile(os.path.join(search_dir, 'app.py')) and os.path.isfile(os.path.join(search_dir, 'exam_bank.db')):
      APP_DIR = search_dir
      break
    parent_dir = os.path.dirname(search_dir)
    if parent_dir == search_dir:
      break
    search_dir = parent_dir
else:
  APP_DIR = os.path.dirname(os.path.abspath(__file__))

# تطبيق PyInstaller ذي الواجهة لا يوفّر stdout/stderr؛ يحتاج Flask مسارًا صالحًا لرسائل التشغيل.
RUNTIME_LOG = None
RUNTIME_LOG_PATH = os.path.join(APP_DIR, 'app.log')
if getattr(sys, 'frozen', False) or sys.stdout is None or sys.stderr is None:
  RUNTIME_LOG = open(RUNTIME_LOG_PATH, 'a', encoding='utf-8', buffering=1)
  sys.stdout = RUNTIME_LOG
  sys.stderr = RUNTIME_LOG

DB_PATH = os.path.join(APP_DIR, 'exam_bank.db')
QR_DATABASE_PATH = os.path.join(APP_DIR, 'qr_database', 'qr_tokens.db')
BACKUP_DIR = os.path.join(APP_DIR, 'backups')
BACKUP_PATH = os.path.join(BACKUP_DIR, 'backup_exam_bank.db')
EXPORT_DIR = os.path.join(APP_DIR, 'exports')
RESOURCE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
LOGO_PATH = os.path.join(RESOURCE_DIR, 'شعار_الكلية_العسكرية_الأولى.png')
WATERMARK_LOGO_PATH = os.path.join(RESOURCE_DIR, 'logo_watermark_transparent.gif')
BACKUP_LOCK = threading.RLock()
os.makedirs(BACKUP_DIR, exist_ok=True)

# حروف الأركان الصغرى الأبجدية العسكرية
ARABIC_ABJAD = ['أ', 'ب', 'جـ', 'د', 'هـ', 'و', 'ز', 'ح', 'ط', 'ي', 'ك', 'ل', 'م', 'ن', 'س', 'ع', 'ف', 'ص', 'ق', 'ر', 'ش', 'ت', 'ث', 'خ', 'ذ', 'ض', 'ظ', 'غ']
ALLOWED_QUESTION_TYPES = {'mc', 'tf', 'fill', 'essay', 'multi_select', 'multi'}
COHORTS = {
  'A': 'الفوج الأول - فوج الإمام علي بن أبي طالب (عليه السلام)',
  'K': 'الفوج الثاني - فوج الإمام موسى الكاظم (عليه السلام)',
  'O': 'فوج عمر بن عبد العزيز'
}

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def validate_cohort(value):
  cohort = str(value or '').strip().upper()
  if cohort not in COHORTS:
    raise ValueError('الفوج المحدد غير صحيح')
  return cohort

def jwt_b64(value):
  return base64.urlsafe_b64encode(
    json.dumps(value, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
  ).decode('ascii').rstrip('=')

def student_jwt_secret(conn):
  row = conn.execute(
    "SELECT value FROM app_state WHERE key = 'student_qr_jwt_secret'"
  ).fetchone()
  if row:
    return row['value']
  secret = secrets.token_urlsafe(48)
  conn.execute(
    "INSERT INTO app_state (key, value) VALUES ('student_qr_jwt_secret', ?)",
    (secret,)
  )
  return secret

def make_student_jwt(conn, student):
  """Creates a permanent, deterministic identity token for one student."""
  header = {'alg': 'HS256', 'typ': 'JWT'}
  payload = {
    'iss': 'military-college',
    'v': 2,
    'sub': str(student['military_number']),
    'name': str(student['student_name']),
    'company': int(student['company']),
    'platoon': str(student['platoon']),
    'sequence': int(student['student_sequence'])
  }
  signing_input = f'{jwt_b64(header)}.{jwt_b64(payload)}'
  signature = base64.urlsafe_b64encode(
    hmac.new(
      student_jwt_secret(conn).encode('utf-8'),
      signing_input.encode('ascii'),
      hashlib.sha256
    ).digest()
  ).decode('ascii').rstrip('=')
  return f'{signing_input}.{signature}'

def decode_student_jwt(token, secret):
  try:
    header_encoded, payload_encoded, signature = token.split('.')
    signing_input = f'{header_encoded}.{payload_encoded}'
    expected = base64.urlsafe_b64encode(
      hmac.new(secret.encode('utf-8'), signing_input.encode('ascii'), hashlib.sha256).digest()
    ).decode('ascii').rstrip('=')
    if not hmac.compare_digest(signature, expected):
      return None
    header = json.loads(base64.urlsafe_b64decode(
      header_encoded + '=' * (-len(header_encoded) % 4)
    ).decode('utf-8'))
    payload = json.loads(base64.urlsafe_b64decode(
      payload_encoded + '=' * (-len(payload_encoded) % 4)
    ).decode('utf-8'))
  except (ValueError, TypeError, UnicodeDecodeError, json.JSONDecodeError,
          base64.binascii.Error):
    return None
  if header != {'alg': 'HS256', 'typ': 'JWT'} or not isinstance(payload, dict):
    return None
  if payload.get('iss') != 'military-college' or payload.get('v') not in (1, 2):
    return None
  required = ('sub', 'name', 'company', 'platoon')
  if any(not str(payload.get(field) or '').strip() for field in required):
    return None
  if payload.get('v') == 2:
    try:
      if int(payload.get('company')) not in (1, 2, 3, 4) or int(payload.get('sequence')) <= 0:
        return None
    except (TypeError, ValueError):
      return None
  return payload

def normalize_image_data(image_data):
  if not image_data:
    return None
  if not isinstance(image_data, str):
    raise ValueError('صيغة الصورة غير صحيحة')
  header, separator, encoded = image_data.partition(',')
  mime_type = header.removeprefix('data:').removesuffix(';base64')
  if not separator or not header.endswith(';base64') or mime_type not in ('image/jpeg', 'image/png'):
    raise ValueError('يسمح بصور PNG أو JPEG فقط')
  try:
    image_bytes = base64.b64decode(encoded, validate=True)
  except (ValueError, base64.binascii.Error):
    raise ValueError('بيانات الصورة غير صالحة')
  if not image_bytes or len(image_bytes) > 5 * 1024 * 1024:
    raise ValueError('الحد الأقصى لحجم الصورة 5 ميغابايت')
  return image_data

def ensure_image_column(conn):
  columns = {row[1] for row in conn.execute('PRAGMA table_info(questions)')}
  if 'image' not in columns:
    conn.execute('ALTER TABLE questions ADD COLUMN image TEXT')
  if 'image_size' not in columns:
    conn.execute('ALTER TABLE questions ADD COLUMN image_size INTEGER NOT NULL DEFAULT 100')

def normalize_image_size(value):
  try:
    size = int(value)
  except (TypeError, ValueError):
    return 100
  return min(100, max(20, size))

def validate_question_payload(data):
  if not isinstance(data, dict):
    raise ValueError('بيانات السؤال غير صحيحة')

  required_fields = ('subject', 'topic', 'year', 'semester', 'type', 'text')
  if any(not str(data.get(field) or '').strip() for field in required_fields):
    raise ValueError('أكمل جميع بيانات السؤال المطلوبة')

  try:
    year = int(data['year'])
    semester = int(data['semester'])
  except (TypeError, ValueError):
    raise ValueError('السنة أو النصف الدراسي غير صحيح')
  if year not in (1, 2, 3) or semester not in (1, 2):
    raise ValueError('السنة أو النصف الدراسي غير صحيح')

  question_type = str(data['type']).strip()
  if question_type not in ALLOWED_QUESTION_TYPES:
    raise ValueError('نوع السؤال غير صحيح')

  options = data.get('options') or []
  if not isinstance(options, list):
    raise ValueError('خيارات السؤال غير صحيحة')
  options = [str(option).strip() for option in options if str(option).strip()]

  return {
    'subject': str(data['subject']).strip(),
    'topic': str(data['topic']).strip(),
    'year': year,
    'semester': semester,
    'type': 'multi_select' if question_type == 'multi' else question_type,
    'text': str(data['text']).strip(),
    'options': options,
    'answer': str(data.get('answer') or '').strip(),
  }

def create_database_backup():
  os.makedirs(BACKUP_DIR, exist_ok=True)
  temporary_path = BACKUP_PATH + '.tmp'
  with BACKUP_LOCK:
    source = sqlite3.connect(DB_PATH)
    target = None
    try:
      target = sqlite3.connect(temporary_path)
      source.backup(target)
      target.close()
      target = None
      os.replace(temporary_path, BACKUP_PATH)
      for name in os.listdir(BACKUP_DIR):
        if name.endswith('.db') and (
          name.startswith('backup_exam_bank_') or name.startswith('before_restore_')
        ):
          os.remove(os.path.join(BACKUP_DIR, name))
    except Exception:
      if target is not None:
        target.close()
      if os.path.exists(temporary_path):
        os.remove(temporary_path)
      raise
    finally:
      source.close()
  return BACKUP_PATH

def get_latest_backup_path():
  return BACKUP_PATH if os.path.isfile(BACKUP_PATH) else None

def auto_backup():
    try:
        if os.path.exists(DB_PATH):
                  create_database_backup()
    except Exception as e:
        print(f"Backup Error: {e}")

def parse_bundled_roster(file_stream, filename):
  """Reads the bundled name lists, which use a serial number instead of a military ID."""
  if not HAS_OPENPYXL:
    raise RuntimeError('مكتبة قراءة ملفات XLSX وXLSM غير مثبتة')
  workbook = openpyxl.load_workbook(file_stream, read_only=True, data_only=True)
  try:
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
  finally:
    workbook.close()
  if not rows:
    raise ValueError(f'ملف Excel فارغ: {filename}')
  headers = [normalize_roster_header(value) for value in rows[0]]
  name_index = next((i for i, value in enumerate(headers) if value.startswith('الاسمالكامل') or value == 'اسمالطالب'), None)
  company_index = next((i for i, value in enumerate(headers) if value in ('السرية', 'رقم السرية')), None)
  platoon_index = next((i for i, value in enumerate(headers) if value in ('الفصيل', 'رقم الفصيل')), None)
  if name_index is None or company_index is None or platoon_index is None:
    raise ValueError(f'عناوين ملف Excel غير مدعومة: {filename}')
  companies = {}
  company_sequences = {}
  seen_sequences = set()
  students = []
  for row_number, row in enumerate(rows[1:], start=2):
    values = list(row)
    if not any(value is not None and str(value).strip() for value in values):
      continue
    name = normalize_roster_cell(values[name_index] if name_index < len(values) else None)
    company_name = normalize_roster_cell(values[company_index] if company_index < len(values) else None)
    platoon = normalize_roster_cell(values[platoon_index] if platoon_index < len(values) else None)
    if not name or not company_name or not platoon:
      raise ValueError(f'بيانات الطالب في ملف "{filename}"، صف {row_number} ناقصة')
    company_key = normalize_import_text(company_name)
    if company_key not in companies:
      companies[company_key] = len(companies) + 1
    serial = values[0] if values else None
    student_sequence = parse_student_sequence(serial)
    company_sequences[company_key] = company_sequences.get(company_key, 0) + 1
    if student_sequence is None:
      student_sequence = company_sequences[company_key]
    identity = (companies[company_key], student_sequence)
    if identity in seen_sequences:
      raise ValueError(
        f'تسلسل الطالب مكرر في السرية {companies[company_key]}: {student_sequence}'
      )
    seen_sequences.add(identity)
    students.append({
      'military_number': make_roster_identifier(companies[company_key], student_sequence),
      'student_name': name,
      'company': companies[company_key],
      'platoon': platoon,
      'student_sequence': student_sequence
    })
  if not students:
    raise ValueError(f'لم توجد أسماء طلاب في ملف Excel: {filename}')
  return students

def ensure_student_qr_tokens(conn):
  rows = conn.execute(
    'SELECT military_number, student_name, company, platoon, student_sequence, qr_jwt FROM exam_roster'
  ).fetchall()
  for row in rows:
    payload = decode_student_jwt(row['qr_jwt'], student_jwt_secret(conn)) if row['qr_jwt'] else None
    token = row['qr_jwt']
    if (
      not payload or payload.get('v') != 2 or
      payload.get('sub') != row['military_number'] or
      payload.get('name') != row['student_name'] or
      int(payload.get('company', 0)) != row['company'] or
      payload.get('platoon') != row['platoon'] or
      int(payload.get('sequence', 0)) != row['student_sequence']
    ):
      token = make_student_jwt(conn, row)
    conn.execute(
      'UPDATE exam_roster SET qr_jwt = ? WHERE military_number = ?',
      (token, row['military_number'])
    )

  if DB_PATH != os.path.join(APP_DIR, 'exam_bank.db'):
    return
  source_dir = os.path.join(APP_DIR, 'qr_database')
  if not os.path.isdir(source_dir):
    return
  os.makedirs(source_dir, exist_ok=True)
  export = sqlite3.connect(QR_DATABASE_PATH)
  try:
    export.execute('''
      CREATE TABLE IF NOT EXISTS student_qr_tokens (
        military_number TEXT PRIMARY KEY,
        student_name TEXT NOT NULL,
        company INTEGER NOT NULL,
        platoon TEXT NOT NULL,
        student_sequence INTEGER NOT NULL,
        jwt TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
      )
    ''')
    export_columns = {row[1] for row in export.execute('PRAGMA table_info(student_qr_tokens)')}
    if 'student_sequence' not in export_columns:
      export.execute('ALTER TABLE student_qr_tokens ADD COLUMN student_sequence INTEGER')
    export.execute('DELETE FROM student_qr_tokens')
    for row in conn.execute(
      'SELECT military_number, student_name, company, platoon, student_sequence, qr_jwt FROM exam_roster'
    ):
      export.execute(
        '''INSERT INTO student_qr_tokens
           (military_number, student_name, company, platoon, student_sequence, jwt)
           VALUES (?, ?, ?, ?, ?, ?)''',
        (row['military_number'], row['student_name'], row['company'],
         row['platoon'], row['student_sequence'], row['qr_jwt'])
      )
    export.commit()
  finally:
    export.close()

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subject TEXT NOT NULL,
            topic TEXT NOT NULL,
            year INTEGER NOT NULL,
            semester INTEGER NOT NULL,
            type TEXT NOT NULL,
            text TEXT NOT NULL,
            options TEXT,
            answer TEXT,
            image TEXT,
            image_size INTEGER NOT NULL DEFAULT 100,
            used INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    ensure_image_column(conn)
    cursor.execute('CREATE TABLE IF NOT EXISTS app_state (key TEXT PRIMARY KEY, value TEXT NOT NULL)')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cohort_question_usage (
            cohort TEXT NOT NULL,
            question_id INTEGER NOT NULL,
            used_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            source_mode TEXT NOT NULL DEFAULT 'exam',
            PRIMARY KEY(cohort, question_id),
            FOREIGN KEY(question_id) REFERENCES questions(id)
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS question_exam_history (
            id TEXT PRIMARY KEY,
            cohort TEXT NOT NULL,
            source_mode TEXT NOT NULL,
            subject TEXT NOT NULL,
            year TEXT,
            semester TEXT,
            question_ids_json TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS exam_roster (
            military_number TEXT PRIMARY KEY,
            student_name TEXT NOT NULL,
            company INTEGER NOT NULL CHECK(company BETWEEN 1 AND 4),
            platoon TEXT NOT NULL,
            student_sequence INTEGER NOT NULL,
            qr_jwt TEXT UNIQUE,
            imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS exam_issues (
            id TEXT PRIMARY KEY,
            course TEXT NOT NULL,
            cohort TEXT NOT NULL,
            subject TEXT NOT NULL,
            platoon_filter TEXT NOT NULL DEFAULT '',
            year INTEGER NOT NULL,
            semester INTEGER NOT NULL,
            exam_date TEXT NOT NULL,
            exam_time TEXT NOT NULL,
            model_a_json TEXT NOT NULL,
            model_b_json TEXT NOT NULL,
            grading_json TEXT NOT NULL DEFAULT '{}',
            print_confirmed INTEGER NOT NULL DEFAULT 0,
            key_transferred INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS exam_booklets (
            qr_token TEXT PRIMARY KEY,
            issue_id TEXT NOT NULL,
            military_number TEXT NOT NULL,
            qr_jwt TEXT,
            company INTEGER NOT NULL CHECK(company BETWEEN 1 AND 4),
            platoon TEXT NOT NULL,
            serial_number INTEGER NOT NULL CHECK(serial_number > 0),
            model TEXT NOT NULL CHECK(model IN ('A', 'B')),
            responses_json TEXT NOT NULL DEFAULT '{}',
            review_json TEXT NOT NULL DEFAULT '[]',
            manual_scores_json TEXT NOT NULL DEFAULT '{}',
            score REAL,
            status TEXT NOT NULL DEFAULT 'issued',
            UNIQUE(issue_id, company, serial_number)
        )
    ''')
    issue_columns = {row[1] for row in cursor.execute('PRAGMA table_info(exam_issues)')}
    if 'platoon_filter' not in issue_columns:
      cursor.execute("ALTER TABLE exam_issues ADD COLUMN platoon_filter TEXT NOT NULL DEFAULT ''")
    roster_columns = {row[1] for row in cursor.execute('PRAGMA table_info(exam_roster)')}
    if 'student_sequence' not in roster_columns:
      cursor.execute('ALTER TABLE exam_roster ADD COLUMN student_sequence INTEGER')
    missing_sequences = cursor.execute(
      '''SELECT military_number, company FROM exam_roster
         WHERE student_sequence IS NULL OR student_sequence <= 0
         ORDER BY company, imported_at, military_number'''
    ).fetchall()
    sequence_by_company = {}
    for row in missing_sequences:
      company = row['company']
      sequence_by_company[company] = sequence_by_company.get(company, 0) + 1
      cursor.execute(
        'UPDATE exam_roster SET student_sequence = ? WHERE military_number = ?',
        (sequence_by_company[company], row['military_number'])
      )
    if 'qr_jwt' not in roster_columns:
      cursor.execute('ALTER TABLE exam_roster ADD COLUMN qr_jwt TEXT')
    booklet_columns = {row[1] for row in cursor.execute('PRAGMA table_info(exam_booklets)')}
    if 'qr_jwt' not in booklet_columns:
      cursor.execute('ALTER TABLE exam_booklets ADD COLUMN qr_jwt TEXT')
    conn.commit()
    ensure_student_qr_tokens(conn)
    conn.close()

# فئة الجسر (API Bridge) بين واجهة WebView وبايثون لفتح نافذة Save As
class ApiBridge:
    def export_word_save_as(self, data_json):
        if not HAS_DOCX:
            return {"status": "error", "message": "مكتبة python-docx غير مثبتة"}
            
        data = json.loads(data_json)
        
        file_path = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=f"امتحان_{data.get('subject')}_النموذجين.docx",
            file_types=('Word Documents (*.docx)', 'All files (*.*)')
        )
        
        if not file_path:
            return {"status": "cancelled"}
            
        target_path = file_path if isinstance(file_path, str) else file_path[0]
        ApiBridge.create_exam_document(data).save(target_path)
        return {"status": "success", "path": target_path}

    @staticmethod
    def create_exam_document(data):
      doc = docx.Document()
      duration = data.get('duration', 'ساعتان')
      exam_date = data.get('examDate') or datetime.now().strftime('%Y/%m/%d')
      models = [('نموذج أ', data.get('modelA', [])), ('نموذج ب', data.get('modelB', []))]

      for model_name, blocks in models:
        heading = doc.add_paragraph()
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading_run = heading.add_run('جمهورية العراق ـــــ وزارة الدفاع')
        heading_run.bold = True
        heading_run.font.name = 'Traditional Arabic'
        heading_run.font.size = Pt(16)

        school = doc.add_paragraph()
        school.alignment = WD_ALIGN_PARAGRAPH.CENTER
        school_run = school.add_run('/ جامعة الدفاع للدراسات العسكرية / الكلية العسكرية / جناح التدريس /')
        school_run.bold = True
        school_run.font.name = 'Traditional Arabic'
        school_run.font.size = Pt(14)

        subject = doc.add_paragraph()
        subject.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subject_run = subject.add_run(f"المادة: {data.get('subject', '')}     ({model_name})")
        subject_run.bold = True
        subject_run.font.name = 'Traditional Arabic'
        subject_run.font.size = Pt(14)

        metadata = doc.add_paragraph()
        metadata.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        metadata_run = metadata.add_run(f'التاريخ: {exam_date}     الوقت: {duration}')
        metadata_run.bold = True
        metadata_run.font.name = 'Traditional Arabic'
        metadata_run.font.size = Pt(12)
        doc.add_paragraph('----------------------------------------------------------------------------------------------------')

        for question_block in blocks:
          doc.add_paragraph(str(question_block.get('title', '')), style='Heading 2')
          for index, item in enumerate(question_block.get('items', [])):
            abjad_symbol = ARABIC_ABJAD[index % len(ARABIC_ABJAD)]
            doc.add_paragraph(f"  {abjad_symbol}- {item.get('text', '')}")

            if item.get('image'):
              _, encoded_image = item['image'].split(',', 1)
              image_paragraph = doc.add_paragraph()
              image_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
              image_paragraph.add_run().add_picture(
                io.BytesIO(base64.b64decode(encoded_image)),
                width=Inches(5.5 * normalize_image_size(item.get('image_size', 100)) / 100)
              )

            if item.get('type') == 'mc' and item.get('options'):
              options = '    '.join(
                f"({ARABIC_ABJAD[i % len(ARABIC_ABJAD)]}) {option}"
                for i, option in enumerate(item.get('options'))
              )
              doc.add_paragraph(f'      {options}')
            elif item.get('type') in ['multi_select', 'multi'] and item.get('options'):
              doc.add_paragraph('      الخيارات المتاحة للترقيم:')
              for i, option in enumerate(item.get('options')):
                doc.add_paragraph(f"        [{ARABIC_ABJAD[i % len(ARABIC_ABJAD)]}] {option}")
              doc.add_paragraph('      إجابة الطالب (اكتب رموز الخيارات الصحيحة فقط): [ ________________________ ]')
            elif item.get('type') == 'essay':
              doc.add_paragraph('      الإجابة: _______________________________________________________________________\n      _______________________________________________________________________________')

        doc.add_page_break()

      return doc

def normalize_import_text(value):
  value = unicodedata.normalize('NFKC', str(value or ''))
  value = ''.join(char for char in value if not unicodedata.combining(char) and unicodedata.category(char) != 'Cf')
  value = value.replace('ـ', '').translate(str.maketrans({'أ': 'ا', 'إ': 'ا', 'آ': 'ا', 'ٱ': 'ا', 'ى': 'ي'}))
  return re.sub(r'\s+', ' ', value).strip().casefold()

def canonicalize_import_subject(value):
  subject = str(value or '').strip()
  if normalize_import_text(subject) == normalize_import_text('الحاسبات'):
    return 'الحاسبات الالكترونية'
  return subject

def parse_import_year(value):
  normalized = normalize_import_text(value)
  digits = normalized.translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
  match = re.search(r'(?<!\d)([123])(?!\d)', digits)
  if match:
    return int(match.group(1))
  for year, words in ((1, ('الاولى', 'الاول', 'اولى', 'اول')), (2, ('الثانية', 'الثاني', 'ثانية', 'ثاني')), (3, ('الثالثة', 'الثالث', 'ثالثة', 'ثالث'))):
    if any(word in normalized for word in words):
      return year
  return None

def parse_import_semester(value):
  normalized = normalize_import_text(value)
  digits = normalized.translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
  match = re.search(r'(?<!\d)([12])(?!\d)', digits)
  if match:
    return int(match.group(1))
  if any(word in normalized for word in ('الثاني', 'الثانية', 'ثاني', 'ثانية')):
    return 2
  if any(word in normalized for word in ('الاول', 'الاولى', 'اول', 'اولى')):
    return 1
  return None

def parse_import_type(value):
  normalized = normalize_import_text(value)
  multiple_answer_phrases = (
    'اختيار الاجابات الصحيحة', 'اختر الاجابات الصحيحة', 'الاجابات الصحيحة فقط',
    'حدد الاجابات الصحيحة', 'حدد الخيارات الصحيحة', 'الاجابات المتعددة',
    'اكثر من اجابة', 'اكثر من خيار'
  )
  selection_verbs = ('اختر', 'حدد', 'انتق', 'اختار')
  multiple_answer_cues = (
    'جميع', 'كل ', 'وسيلتين', 'وسيلتان', 'اجابتين', 'اجابتان',
    'خيارين', 'خياران', 'اثنين', 'اثنتين', 'اثنان', 'ثلاث', 'ارب',
    'اكثر من', 'متعددة'
  )
  if any(phrase in normalized for phrase in multiple_answer_phrases) or (
    any(verb in normalized for verb in selection_verbs) and
    any(cue in normalized for cue in multiple_answer_cues)
  ):
    return 'multi_select'
  if any(phrase in normalized for phrase in ('صح وخطا', 'صح او خطا', 'صح ام خطا')):
    return 'tf'
  if any(phrase in normalized for phrase in ('ملء الفراغ', 'املأ الفراغ', 'املئ الفراغ', 'اكمل الفراغ', 'اكمل الجملة', 'اكمل العبارة', 'املأ التالي')):
    return 'fill'
  if any(phrase in normalized for phrase in ('اختيار من متعدد', 'اختر الاجابة الصحيحة', 'اختر الاجابة المناسبة', 'اختر الصحيح', 'اي مما يلي')):
    return 'mc'
  if any(phrase in normalized for phrase in ('شرح', 'مقالي')):
    return 'essay'
  if normalized.startswith(('اشرح ', 'علل ', 'فسر ', 'وضح ', 'قارن ', 'ناقش ', 'عرف ', 'عدد ', 'اذكر ', 'تحدث عن ', 'تكلم عن ', 'ما اهمية ', 'ما المقصود ')):
    return 'essay'
  return None

def infer_import_question_type(question):
  explicit_type = question.get('type')
  if explicit_type in ('mc', 'tf', 'fill', 'essay', 'multi_select'):
    return explicit_type

  text = str(question.get('text') or '')
  stem_type = parse_import_type(text)
  if stem_type == 'multi_select':
    return 'multi_select'
  if stem_type in ('tf', 'fill'):
    return stem_type

  answer = normalize_import_text(question.get('answer'))
  if answer in ('صح', 'خطا', 'صحيح', 'غير صحيح'):
    return 'tf'
  options = [normalize_import_text(option).strip(' .،,') for option in question.get('options', [])]
  if options and all(option in ('صح', 'خطا', 'صحيح', 'غير صحيح') for option in options):
    return 'tf'
  if re.search(r'(?:_{2,}|ـ{2,}|\.{3,}|…)', text) or any(phrase in normalize_import_text(text) for phrase in ('املأ الفراغ', 'املئ الفراغ', 'اكمل الفراغ')):
    return 'fill'
  if options:
    return 'mc'
  if stem_type == 'mc':
    return 'mc'
  if stem_type == 'essay':
    return 'essay'

  normalized_text = normalize_import_text(text)
  if any(normalized_text.startswith(prefix) for prefix in (
    'اشرح ', 'عرف ', 'عرّف ', 'عدد ', 'اذكر ', 'قارن ', 'ناقش ',
    'علل ', 'فسر ', 'وضح ', 'بين ', 'تحدث عن ', 'تكلم عن ',
    'ما اهمية ', 'ما المقصود ', 'ما تعريف '
  )):
    return 'essay'
  return 'essay'

def clean_import_value(value):
  value = str(value or '').strip()
  if value.startswith('[') and value.endswith(']'):
    value = value[1:-1].strip()
  normalized = normalize_import_text(value)
  placeholders = ('تملأ من قبل', 'تملا من قبل', 'اكتب ', 'اختر ', 'عند الحاجة', 'كما يظهر في البرنامج', 'اسم المادة', 'اسم المفردة')
  if not normalized or any(placeholder in normalized for placeholder in placeholders):
    return ''
  if ' / ' in value or ' او ' in normalized or ' أو ' in value:
    return ''
  return value

def iter_import_document_lines(document):
  for block in document.iter_inner_content():
    if isinstance(block, DocxParagraph):
      if block.text.strip():
        yield block.text.strip()
    elif isinstance(block, DocxTable):
      for row in block.rows:
        cells = [' '.join(paragraph.text.strip() for paragraph in cell.paragraphs if paragraph.text.strip()) for cell in row.cells]
        line = ' | '.join(cell for cell in cells if cell)
        if line:
          yield line

def parse_import_field(line):
  parts = re.split(r'[:：|]', str(line), maxsplit=1)
  if len(parts) != 2:
    return None, None
  label = normalize_import_text(parts[0]).removeprefix('اسم ')
  value = clean_import_value(parts[1])
  if label in ('المادة', 'المادة الدراسية', 'اسم المادة'):
    return 'subject', canonicalize_import_subject(value)
  if label in ('المفردة', 'اسم المفردة', 'الموضوع'):
    return 'topic', value
  if label in ('السنة', 'السنة الدراسية'):
    return 'year', parse_import_year(value)
  if label in ('النصف', 'النصف الدراسي'):
    return 'semester', parse_import_semester(value)
  if label == 'نوع السؤال':
    return 'type', parse_import_type(value)
  return None, None

def parse_import_answer(line):
  match = re.match(r'^\s*(?:الإجابة الصحيحة|الاجابة الصحيحة|الإجابات الصحيحة|الاجابات الصحيحة|الإجابة النموذجية|الاجابة النموذجية|الإجابة|الاجابة)\s*[:：\-]\s*(.*)$', line)
  if not match:
    return None
  return clean_import_value(match.group(1))

def parse_import_options(line):
  marker = re.compile(r'[\(\[]\s*(?P<label>جـ|[أابجدهوزحطيكلمنسفصقرشثخذضظغ]|[A-Za-z]|\d+)\s*[\)\]]\s*')
  matches = list(marker.finditer(line))
  bracketed = bool(matches and line[matches[0].start()] == '[')
  if not matches:
    marker = re.compile(r'(?:^|\s)(?:جـ|[أابجدهوزحطيكلمنسفصقرشثخذضظغ])\s*[-.)]\s*')
    matches = list(marker.finditer(line))
  if not matches:
    return None
  options = []
  for index, match in enumerate(matches):
    end = matches[index + 1].start() if index + 1 < len(matches) else len(line)
    value = clean_import_value(line[match.end():end])
    if value and 'الخيار ' not in normalize_import_text(value):
      options.append(value)
  return line[:matches[0].start()].strip(), options, bracketed

def looks_like_import_question(value, question_type=None):
  normalized = normalize_import_text(value)
  if normalized.endswith(('؟', '?')):
    return True
  if normalized.startswith(('ما ', 'ماذا ', 'كيف ', 'لماذا ', 'متى ', 'أين ', 'اين ', 'اذكر ', 'عدد ', 'اكتب ', 'قارن ', 'اشرح ', 'عرف ', 'بين ', 'حدد ', 'هل ', 'أكمل ', 'اكمل ', 'املأ ', 'ضع ')):
    return True
  if re.search(r'(?:_{2,}|ـ{2,}|\.{3,}|…)', str(value)):
    return True
  return question_type == 'tf' and normalized.endswith(('.', '؛', ';'))

def parse_import_word_document(file_stream):
  if not HAS_DOCX:
    raise ValueError('مكتبة python-docx غير مثبتة')
  document = docx.Document(file_stream)
  lines = list(iter_import_document_lines(document))
  start_index = None
  question_number = re.compile(r'^\s*(?:سؤال\s*)?[0-9٠-٩۰-۹]+\s*[-.)：:|]\s*(.*)$')
  for index, line in enumerate(lines):
    marker = normalize_import_text(line).strip(' .:：-–—')
    if marker.startswith(('قالب التعبئة', 'قالب تعبئة الأسئلة', 'قالب تعبئة الاسئلة')):
      start_index = index + 1
      break
  if start_index is None:
    for index, line in enumerate(lines):
      field, value = parse_import_field(line)
      if field == 'subject' and value:
        start_index = index
        break
  if start_index is None:
    for index, line in enumerate(lines):
      if question_number.match(line):
        start_index = index
        break
  if start_index is None:
    raise ValueError('لم يتم العثور على حقول التصنيف أو سؤال مرقم في المستند')
  metadata = {'subject': '', 'topic': '', 'year': None, 'semester': None, 'type': None}
  questions = []
  current_question = None
  last_line_was_answer = False
  in_question_section = False
  option_marker = re.compile(r'[\(\[]\s*(?:جـ|[أابجدهوزحطيكلمنسفصقرشثخذضظغ]|[A-Za-z]|\d+)\s*[\)\]]')

  def add_question(text):
    nonlocal current_question, last_line_was_answer
    text = str(text or '').strip()
    if text.startswith('[') and text.endswith(']'):
      text = text[1:-1].strip()
    normalized_text = normalize_import_text(text)
    if (not normalized_text or
        normalized_text in ('اكتب نص السؤال هنا', 'اكتب السؤال هنا', 'نص السؤال هنا', 'question text here') or
        normalized_text.startswith('قبل ارسال الملف')):
      current_question = None
      return
    current_question = {
      'subject': metadata['subject'], 'topic': metadata['topic'],
      'year': metadata['year'], 'semester': metadata['semester'],
      'type': metadata['type'], 'text': text, 'options': [], 'answer': ''
    }
    questions.append(current_question)
    last_line_was_answer = False

  for raw_line in lines[start_index:]:
    line = raw_line.strip()
    normalized = normalize_import_text(line)
    if not normalized:
      continue
    if normalized in ('الأسئلة', 'الاسئلة', 'بيانات المجموعة'):
      in_question_section = normalized != 'بيانات المجموعة'
      current_question = None
      continue
    field, value = parse_import_field(line)
    if field:
      if field == 'subject':
        metadata = {'subject': value, 'topic': '', 'year': None, 'semester': None, 'type': None}
      else:
        metadata[field] = value
      current_question = None
      last_line_was_answer = False
      continue
    if normalized.startswith(('الأسئلة:', 'الاسئلة:')):
      in_question_section = True
      continue
    answer = parse_import_answer(line)
    if answer is not None:
      if current_question:
        current_question['answer'] = answer
        last_line_was_answer = True
      continue
    option_data = parse_import_options(line)
    number_match = question_number.match(line)
    if number_match:
      text = number_match.group(1)
      if option_data:
        marker = option_marker.search(text)
        if marker:
          text = text[:marker.start()].strip()
      add_question(text)
      if current_question and option_data:
        current_question['options'].extend(option_data[1])
      continue
    if option_data:
      prefix, options, bracketed = option_data
      if prefix:
        prefix_number = question_number.match(prefix)
        if prefix_number:
          add_question(prefix_number.group(1))
        elif not current_question:
          add_question(prefix)
      if current_question:
        current_question['options'].extend(options)
        if bracketed and current_question['type'] is None:
          current_question['type'] = 'multi_select'
        last_line_was_answer = False
      continue
    if question_number.match(line):
      add_question(question_number.match(line).group(1))
    elif current_question and (last_line_was_answer or looks_like_import_question(line, metadata['type']) or (metadata['type'] == 'tf' and line.endswith('.'))):
      add_question(line)
    elif current_question:
      current_question['text'] = (current_question['text'] + '\n' + line).strip()
    elif in_question_section or metadata['subject']:
      add_question(line)

  for question in questions:
    question['type'] = infer_import_question_type(question)

  unique_questions = []
  seen = {}
  for question in questions:
    key = (question['subject'].casefold(), question['topic'].casefold(), question['year'], question['semester'], question['type'], normalize_import_text(question['text']))
    existing = seen.get(key)
    if existing is None:
      seen[key] = question
      unique_questions.append(question)
    else:
      for option in question['options']:
        if option not in existing['options']:
          existing['options'].append(option)
      if not existing['answer']:
        existing['answer'] = question['answer']
  return unique_questions

def import_question_key(question):
  return (
    normalize_import_text(question.get('subject')),
    normalize_import_text(question.get('topic')),
    question.get('year'), question.get('semester'), question.get('type'),
    normalize_import_text(question.get('text'))
  )

# ==========================================
# APIs الخلفية
# ==========================================

def question_rows_with_cohort_usage(conn, cohort=None, query=None):
    params = []
    where = ''
    if query:
      where = ' WHERE text LIKE ? OR answer LIKE ?'
      params = [f'%{query}%', f'%{query}%']
    rows = conn.execute(
      f'''SELECT q.*, CASE WHEN u.question_id IS NULL THEN 0 ELSE 1 END AS cohort_used
          FROM questions q
          LEFT JOIN cohort_question_usage u
            ON u.question_id = q.id AND u.cohort = ?
          {where}
          ORDER BY q.id DESC''',
      [cohort or ''] + params
    ).fetchall()
    result = []
    for row in rows:
      question = dict(row)
      question['used'] = bool(question['cohort_used']) if cohort else bool(question['used'])
      question['cohort_used'] = bool(question['cohort_used'])
      if question.get('options'):
        try:
          question['options'] = json.loads(question['options'])
        except (TypeError, ValueError):
          pass
      result.append(question)
    return result

def record_question_history(conn, cohort, source_mode, subject, year, semester, question_ids):
  cohort = validate_cohort(cohort)
  ids = sorted({int(question_id) for question_id in question_ids if int(question_id) > 0})
  if not ids:
    return None
  history_id = secrets.token_urlsafe(18)
  conn.execute(
    '''INSERT INTO question_exam_history
       (id, cohort, source_mode, subject, year, semester, question_ids_json)
       VALUES (?, ?, ?, ?, ?, ?, ?)''',
    (history_id, cohort, str(source_mode or 'generated'), str(subject or ''),
     str(year or ''), str(semester or ''),
     json.dumps(ids, ensure_ascii=False))
  )
  return history_id

@app.route('/api/questions', methods=['GET'])
def get_questions():
    conn = get_db()
    try:
      cohort = validate_cohort(request.args.get('cohort')) if request.args.get('cohort') else None
    except ValueError as error:
      conn.close()
      return jsonify({'status': 'error', 'message': str(error)}), 400
    questions = question_rows_with_cohort_usage(conn, cohort)
    conn.close()
    return jsonify(questions)

@app.route('/api/search', methods=['GET'])
def search_questions():
    query = request.args.get('q', '').strip()
    if not query:
        return jsonify([])
    conn = get_db()
    try:
      cohort = validate_cohort(request.args.get('cohort')) if request.args.get('cohort') else None
    except ValueError as error:
      conn.close()
      return jsonify({'status': 'error', 'message': str(error)}), 400
    results = question_rows_with_cohort_usage(conn, cohort, query)
    conn.close()
    return jsonify(results)

@app.route('/api/import-word/preview', methods=['POST'])
def preview_word_import():
  if not HAS_DOCX:
    return jsonify({"status": "error", "message": "مكتبة python-docx غير مثبتة"}), 503
  uploaded_file = request.files.get('file')
  if not uploaded_file or not uploaded_file.filename:
    return jsonify({"status": "error", "message": "اختر مستند Word أولاً"}), 400
  if not uploaded_file.filename.lower().endswith('.docx'):
    return jsonify({"status": "error", "message": "يدعم الاستيراد مستندات Word بصيغة DOCX فقط"}), 415
  try:
    questions = parse_import_word_document(uploaded_file.stream)
  except Exception as error:
    return jsonify({"status": "error", "message": "تعذر قراءة المستند: " + str(error)}), 400
  if not questions:
    return jsonify({"status": "error", "message": "لم يتم العثور على أسئلة مرقمة بعد قسم قالب التعبئة"}), 422
  if len(questions) > 1000:
    return jsonify({"status": "error", "message": "يتجاوز المستند الحد الأقصى البالغ 1000 سؤال"}), 413

  conn = get_db()
  existing_keys = {import_question_key(dict(row)) for row in conn.execute(
    'SELECT subject, topic, year, semester, type, text FROM questions'
  ).fetchall()}
  conn.close()
  for question in questions:
    question['duplicate'] = import_question_key(question) in existing_keys
  return jsonify({"status": "success", "questions": questions, "count": len(questions)})

@app.route('/api/import-word', methods=['POST'])
def import_word_questions():
  data = request.get_json(silent=True) or {}
  raw_questions = data.get('questions')
  if not isinstance(raw_questions, list) or not raw_questions:
    return jsonify({"status": "error", "message": "لا توجد أسئلة محددة للاستيراد"}), 400
  if len(raw_questions) > 1000:
    return jsonify({"status": "error", "message": "يتجاوز الاستيراد الحد الأقصى البالغ 1000 سؤال"}), 413

  allowed_types = {'mc', 'tf', 'fill', 'essay', 'multi_select'}
  questions = []
  for index, question in enumerate(raw_questions, start=1):
    if not isinstance(question, dict):
      return jsonify({"status": "error", "message": f"بيانات السؤال {index} غير صحيحة"}), 400
    subject = str(question.get('subject') or '').strip()
    topic = str(question.get('topic') or '').strip()
    text = str(question.get('text') or '').strip()
    question_type = 'multi_select' if question.get('type') == 'multi' else question.get('type')
    try:
      year = int(question.get('year'))
      semester = int(question.get('semester'))
    except (TypeError, ValueError):
      return jsonify({"status": "error", "message": f"أكمل السنة والنصف للسؤال {index}"}), 400
    if not subject or not topic or not text:
      return jsonify({"status": "error", "message": f"أكمل المادة والمفردة ونص السؤال {index}"}), 400
    if year not in (1, 2, 3) or semester not in (1, 2):
      return jsonify({"status": "error", "message": f"السنة أو النصف غير صحيح للسؤال {index}"}), 400
    if question_type not in allowed_types:
      return jsonify({"status": "error", "message": f"اختر نوعًا صحيحًا للسؤال {index}"}), 400
    options = question.get('options', [])
    if not isinstance(options, list):
      return jsonify({"status": "error", "message": f"خيارات السؤال {index} غير صحيحة"}), 400
    questions.append({
      'subject': subject, 'topic': topic, 'year': year, 'semester': semester,
      'type': question_type, 'text': text,
      'options': [str(option).strip() for option in options if str(option).strip()],
      'answer': str(question.get('answer') or '').strip()
    })

  conn = get_db()
  inserted = 0
  skipped_duplicates = 0
  try:
    conn.execute('BEGIN IMMEDIATE')
    existing_keys = {import_question_key(dict(row)) for row in conn.execute(
      'SELECT subject, topic, year, semester, type, text FROM questions'
    ).fetchall()}
    for question in questions:
      key = import_question_key(question)
      if key in existing_keys:
        skipped_duplicates += 1
        continue
      options_json = json.dumps(question['options'], ensure_ascii=False) if question['options'] else None
      conn.execute(
        '''INSERT INTO questions (subject, topic, year, semester, type, text, options, answer, image, image_size, used)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL, 100, 0)''',
        (question['subject'], question['topic'], question['year'], question['semester'],
         question['type'], question['text'], options_json, question['answer'])
      )
      existing_keys.add(key)
      inserted += 1
    conn.commit()
  except sqlite3.Error as error:
    conn.rollback()
    conn.close()
    return jsonify({"status": "error", "message": "تعذر حفظ الأسئلة المستوردة: " + str(error)}), 500
  conn.close()
  if inserted:
    auto_backup()
  return jsonify({"status": "success", "imported": inserted, "skippedDuplicates": skipped_duplicates})

@app.route('/api/questions', methods=['POST'])
def add_question():
    data = request.get_json(silent=True)
    try:
        question = validate_question_payload(data)
        image_data = normalize_image_data(data.get('image'))
    except (TypeError, ValueError) as error:
        return jsonify({"status": "error", "message": str(error)}), 400
    options_str = json.dumps(question['options'], ensure_ascii=False) if question['options'] else None
    image_size = normalize_image_size(data.get('image_size', 100))
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO questions (subject, topic, year, semester, type, text, options, answer, image, image_size, used)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
    ''', (
        question['subject'], question['topic'], question['year'],
        question['semester'], question['type'], question['text'],
        options_str, question['answer'], image_data, image_size
    ))
    conn.commit()
    q_id = cursor.lastrowid
    conn.close()
    auto_backup()
    return jsonify({"status": "success", "id": q_id})

@app.route('/api/questions/<int:q_id>', methods=['PUT'])
def update_question(q_id):
    data = request.get_json(silent=True)
    try:
        question = validate_question_payload(data)
        image_data = normalize_image_data(data.get('image'))
    except (TypeError, ValueError) as error:
        return jsonify({"status": "error", "message": str(error)}), 400
    options_str = json.dumps(question['options'], ensure_ascii=False) if question['options'] else None
    image_size = normalize_image_size(data.get('image_size', 100))
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE questions 
        SET subject=?, topic=?, year=?, semester=?, type=?, text=?, options=?, answer=?, image=?, image_size=?
        WHERE id=?
    ''', (
        question['subject'], question['topic'], question['year'],
        question['semester'], question['type'], question['text'],
        options_str, question['answer'], image_data, image_size, q_id
    ))
    if cursor.rowcount == 0:
        conn.close()
        return jsonify({"status": "error", "message": "السؤال غير موجود"}), 404
    conn.commit()
    conn.close()
    auto_backup()
    return jsonify({"status": "success"})

@app.route('/api/questions/<int:q_id>', methods=['DELETE'])
def delete_question(q_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM questions WHERE id = ?", (q_id,))
    if cursor.rowcount == 0:
        conn.close()
        return jsonify({"status": "error", "message": "السؤال غير موجود"}), 404
    conn.commit()
    conn.close()
    auto_backup()
    return jsonify({"status": "success"})

@app.route('/api/questions/reset-used', methods=['POST'])
def reset_used():
    data = request.get_json(silent=True) or {}
    conn = get_db()
    try:
      cohort = validate_cohort(data.get('cohort'))
    except ValueError as error:
      conn.close()
      return jsonify({"status": "error", "message": str(error)}), 400
    conn.execute("DELETE FROM cohort_question_usage WHERE cohort = ?", (cohort,))
    conn.execute(
      '''UPDATE questions SET used = CASE WHEN EXISTS (
           SELECT 1 FROM cohort_question_usage u WHERE u.question_id = questions.id
         ) THEN 1 ELSE 0 END'''
    )
    conn.execute("DELETE FROM question_exam_history WHERE cohort = ?", (cohort,))
    conn.commit()
    conn.close()
    auto_backup()
    return jsonify({"status": "success", "cohort": cohort})

@app.route('/api/questions/mark-used', methods=['POST'])
def mark_used():
    data = request.get_json(silent=True) or {}
    ids = data.get('ids', [])
    if not isinstance(ids, list):
        return jsonify({"status": "error", "message": "قائمة الأسئلة غير صحيحة"}), 400
    try:
      cohort = validate_cohort(data.get('cohort'))
    except ValueError as error:
      return jsonify({"status": "error", "message": str(error)}), 400
    ids = list(dict.fromkeys(q_id for q_id in ids if type(q_id) is int and q_id > 0))
    if ids:
        conn = get_db()
        cursor = conn.cursor()
        cursor.executemany(
          '''INSERT OR IGNORE INTO cohort_question_usage
             (cohort, question_id, source_mode) VALUES (?, ?, ?)''',
          [(cohort, q_id, str(data.get('source_mode') or 'exam')) for q_id in ids]
        )
        cursor.executemany("UPDATE questions SET used = 1 WHERE id = ?", [(q_id,) for q_id in ids])
        conn.commit()
        conn.close()
        auto_backup()
    return jsonify({"status": "success", "cohort": cohort, "used": len(ids)})

@app.route('/api/questions/mark-available', methods=['POST'])
def mark_available():
    data = request.get_json(silent=True) or {}
    raw_ids = data.get('ids', [])
    if not isinstance(raw_ids, list):
        return jsonify({"status": "error", "message": "قائمة الأسئلة غير صحيحة"}), 400
    ids = list(dict.fromkeys(q_id for q_id in raw_ids if type(q_id) is int and q_id > 0))
    if not ids:
        return jsonify({"status": "error", "message": "لا توجد أسئلة للاسترجاع"}), 400
    try:
      cohort = validate_cohort(data.get('cohort'))
    except ValueError as error:
      return jsonify({"status": "error", "message": str(error)}), 400

    conn = get_db()
    cursor = conn.cursor()
    cursor.executemany(
      "DELETE FROM cohort_question_usage WHERE cohort = ? AND question_id = ?",
      [(cohort, q_id) for q_id in ids]
    )
    cursor.execute(
      '''UPDATE questions SET used = CASE WHEN EXISTS (
           SELECT 1 FROM cohort_question_usage u WHERE u.question_id = questions.id
         ) THEN 1 ELSE 0 END WHERE id IN ({})'''.format(','.join('?' for _ in ids)),
      ids
    )
    restored = cursor.rowcount
    conn.commit()
    conn.close()
    auto_backup()
    return jsonify({"status": "success", "restored": restored, "cohort": cohort})

@app.route('/api/issued/last', methods=['GET', 'POST'])
def last_issued_exam():
    conn = get_db()
    conn.execute('CREATE TABLE IF NOT EXISTS app_state (key TEXT PRIMARY KEY, value TEXT NOT NULL)')
    if request.method == 'GET':
        row = conn.execute("SELECT value FROM app_state WHERE key = 'last_issued_exam'").fetchone()
        conn.close()
        return jsonify({"issue": json.loads(row['value']) if row else None})

    data = request.get_json(silent=True) or {}
    raw_ids = data.get('ids', [])
    if not isinstance(data.get('subject'), str) or not isinstance(raw_ids, list):
        conn.close()
        return jsonify({"status": "error", "message": "بيانات الإصدار غير صحيحة"}), 400
    ids = list(dict.fromkeys(q_id for q_id in raw_ids if type(q_id) is int and q_id > 0))
    if not ids:
        conn.close()
        return jsonify({"status": "error", "message": "لا توجد أسئلة في الإصدار"}), 400

    try:
      cohort = validate_cohort(data.get('cohort'))
    except ValueError as error:
      conn.close()
      return jsonify({"status": "error", "message": str(error)}), 400
    issue = {key: data.get(key) for key in ('subject', 'year', 'semester', 'pool', 'order', 'duration', 'durationValue', 'examDate', 'sourceMode')}
    issue['cohort'] = cohort
    issue['ids'] = ids
    conn.execute(
        "INSERT OR REPLACE INTO app_state (key, value) VALUES ('last_issued_exam', ?)",
        (json.dumps(issue, ensure_ascii=False),)
    )
    record_question_history(
      conn, cohort, issue.get('sourceMode'), issue.get('subject'),
      issue.get('year'), issue.get('semester'), ids
    )
    conn.commit()
    conn.close()
    auto_backup()
    return jsonify({"status": "success"})

@app.route('/api/cohorts', methods=['GET'])
def get_cohorts():
    return jsonify([
      {'code': code, 'name': name} for code, name in COHORTS.items()
    ])

@app.route('/api/jwt-database', methods=['GET'])
def browse_jwt_database():
    query = request.args.get('q', '').strip()
    try:
      page = max(1, int(request.args.get('page', 1)))
      page_size = min(200, max(1, int(request.args.get('page_size', 50))))
    except (TypeError, ValueError):
      return jsonify({'status': 'error', 'message': 'ترقيم صفحات قاعدة JWT غير صحيح'}), 400
    if not os.path.isfile(QR_DATABASE_PATH):
      return jsonify({'status': 'success', 'total': 0, 'page': page, 'page_size': page_size, 'rows': []})
    conn = sqlite3.connect(QR_DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    where = ''
    params = []
    if query:
      where = ' WHERE student_sequence LIKE ? OR student_name LIKE ? OR jwt LIKE ?'
      params = [f'%{query}%', f'%{query}%', f'%{query}%']
    total = conn.execute(
      f'SELECT COUNT(*) FROM student_qr_tokens{where}', params
    ).fetchone()[0]
    rows = conn.execute(
      f'''SELECT student_sequence, student_name, company, platoon, jwt, created_at
          FROM student_qr_tokens{where}
          ORDER BY company, student_sequence LIMIT ? OFFSET ?''',
      params + [page_size, (page - 1) * page_size]
    ).fetchall()
    conn.close()
    return jsonify({
      'status': 'success', 'total': total, 'page': page,
      'page_size': page_size, 'rows': [dict(row) for row in rows]
    })

@app.route('/api/question-memory', methods=['GET'])
def get_question_memory():
    conn = get_db()
    counts = {}
    for code in COHORTS:
      counts[code] = conn.execute(
        'SELECT COUNT(*) FROM cohort_question_usage WHERE cohort = ?', (code,)
      ).fetchone()[0]
    rows = conn.execute(
      '''SELECT id, cohort, source_mode, subject, year, semester,
                question_ids_json, created_at
         FROM question_exam_history ORDER BY created_at DESC LIMIT 200'''
    ).fetchall()
    conn.close()
    history = []
    for row in rows:
      item = dict(row)
      item['cohort_name'] = COHORTS.get(item['cohort'], item['cohort'])
      item['question_count'] = len(json.loads(item.pop('question_ids_json') or '[]'))
      history.append(item)
    return jsonify({'status': 'success', 'cohorts': COHORTS, 'used_counts': counts, 'history': history})

@app.route('/api/export-word', methods=['POST'])
def export_word_document():
    if not HAS_DOCX:
        return jsonify({"status": "error", "message": "مكتبة python-docx غير مثبتة"}), 503

    data = request.get_json(silent=True) or {}
    if not isinstance(data.get('modelA'), list) or not isinstance(data.get('modelB'), list):
        return jsonify({"status": "error", "message": "بيانات النموذجين غير صحيحة"}), 400

    try:
        document = ApiBridge.create_exam_document(data)
        os.makedirs(EXPORT_DIR, exist_ok=True)
        filename = f"exam_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.docx"
        target_path = os.path.join(EXPORT_DIR, filename)
        document.save(target_path)
        return jsonify({"status": "success", "file": filename, "path": target_path})
    except Exception as error:
        return jsonify({"status": "error", "message": str(error)}), 500

@app.route('/api/backup', methods=['POST'])
def create_backup():
  try:
    backup_path = create_database_backup()
    return jsonify({"status": "success", "file": os.path.basename(backup_path)})
  except (OSError, sqlite3.Error) as error:
    return jsonify({"status": "error", "message": str(error)}), 500

@app.route('/api/backup/restore', methods=['POST'])
def restore_latest_backup():
  backup_path = get_latest_backup_path()
  if not backup_path:
    return jsonify({"status": "error", "message": "لا توجد نسخة احتياطية للاستيراد"}), 404

  source = None
  snapshot = None
  target = None
  try:
    with BACKUP_LOCK:
      source = sqlite3.connect(backup_path)
      integrity = source.execute('PRAGMA integrity_check').fetchone()[0]
      if integrity != 'ok':
        return jsonify({"status": "error", "message": "النسخة الاحتياطية تالفة"}), 400

      question_count = source.execute('SELECT COUNT(*) FROM questions').fetchone()[0]
      snapshot = sqlite3.connect(':memory:')
      source.backup(snapshot)
      source.close()
      source = None
      target = get_db()
      snapshot.backup(target)
      ensure_image_column(target)
      target.commit()
      target.close()
      target = None
      init_db()
      create_database_backup()
      return jsonify({"status": "success", "questions": question_count})
  except (OSError, sqlite3.Error) as error:
    return jsonify({"status": "error", "message": str(error)}), 500
  finally:
    if target is not None:
      target.close()
    if source is not None:
      source.close()
    if snapshot is not None:
      snapshot.close()

def exam_questions(blocks):
  questions = []
  for block_index, block in enumerate(blocks):
    if not isinstance(block, dict) or not isinstance(block.get('items'), list):
      continue
    for item_index, item in enumerate(block['items']):
      if isinstance(item, dict):
        questions.append({
          'key': f'{block_index}:{item_index}',
          'block_index': block_index,
          'item_index': item_index,
          'block_title': str(block.get('title') or ''),
          'item': item
        })
  return questions

def exam_answer_indexes(item):
  options = item.get('options')
  if item.get('type') == 'tf' and (not isinstance(options, list) or not options):
    options = ['صح', 'خطأ']
  answer = str(item.get('answer') or '').strip()
  if not isinstance(options, list) or not answer:
    return []
  tokens = [token.strip() for token in re.split(r'[،,/\\\n]+', answer) if token.strip()]
  indexes = []
  for index, option in enumerate(options):
    symbol = ARABIC_ABJAD[index % len(ARABIC_ABJAD)]
    if any(
      token == symbol or token.replace('ـ', '') == symbol.replace('ـ', '') or
      token.casefold() == str(option).strip().casefold()
      for token in tokens
    ):
      indexes.append(index)
  return indexes

def qr_signing_secret(conn):
  row = conn.execute("SELECT value FROM app_state WHERE key = 'exam_qr_signing_secret'").fetchone()
  if row:
    return row['value']
  secret = secrets.token_urlsafe(48)
  conn.execute(
    "INSERT INTO app_state (key, value) VALUES ('exam_qr_signing_secret', ?)",
    (secret,)
  )
  return secret

def make_exam_qr(conn, issue_id, cohort, company, serial_number, model, qr_token, student_jwt=None):
  if not HAS_QRCODE:
    raise RuntimeError('مكتبة توليد QR غير مثبتة')
  payload = {
    'v': 1, 'i': issue_id, 'f': cohort, 'c': company, 's': serial_number,
    'm': model, 'k': qr_token
  }
  encoded = base64.urlsafe_b64encode(
    json.dumps(payload, separators=(',', ':')).encode('utf-8')
  ).decode('ascii').rstrip('=')
  signature = hmac.new(
    qr_signing_secret(conn).encode('utf-8'), encoded.encode('ascii'), hashlib.sha256
  ).hexdigest()[:24]
  code = f'EXM1.{encoded}.{signature}'
  image = qrcode.make(code)
  buffer = io.BytesIO()
  image.save(buffer, format='PNG')
  return 'data:image/png;base64,' + base64.b64encode(buffer.getvalue()).decode('ascii')

def make_student_qr(student_jwt):
  if not HAS_QRCODE:
    raise RuntimeError('مكتبة توليد QR غير مثبتة')
  image = qrcode.make(student_jwt)
  buffer = io.BytesIO()
  image.save(buffer, format='PNG')
  return 'data:image/png;base64,' + base64.b64encode(buffer.getvalue()).decode('ascii')

def decode_exam_qr(code, secret):
  try:
    prefix, encoded, signature = code.split('.', 2)
    expected = hmac.new(secret.encode('utf-8'), encoded.encode('ascii'), hashlib.sha256).hexdigest()[:24]
    if prefix != 'EXM1' or not hmac.compare_digest(signature, expected):
      return None
    padded = encoded + '=' * (-len(encoded) % 4)
    payload = json.loads(base64.urlsafe_b64decode(padded.encode('ascii')))
  except (ValueError, TypeError, json.JSONDecodeError):
    return None
  if not isinstance(payload, dict) or payload.get('v') != 1:
    return None
  return payload

def find_booklet_by_student_jwt(conn, token):
  rows = conn.execute(
    '''SELECT b.*, i.model_a_json, i.model_b_json, i.grading_json,
              i.key_transferred, i.cohort
       FROM exam_booklets b JOIN exam_issues i ON i.id = b.issue_id
       WHERE b.qr_jwt = ? AND b.status != 'corrected'
       ORDER BY b.rowid DESC''',
    (token,)
  ).fetchall()
  return rows[0] if len(rows) == 1 else None

def normalize_company_number(value):
  text = str(value or '').strip().translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
  match = re.search(r'(?<!\d)([1-4])(?!\d)', text)
  if match:
    return int(match.group(1))
  normalized = normalize_import_text(text)
  for number, word in ((1, 'الاول'), (2, 'الثاني'), (3, 'الثالث'), (4, 'الرابع')):
    if word in normalized:
      return number
  return None

ROSTER_HEADER_ALIASES = {
  'student_sequence': (
    'التسلسل', 'تسلسل الطالب', 'رقم التسلسل', 'ت', 'serial', 'sequence'
  ),
  'student_name': (
    'اسم الطالب', 'الاسم الكامل', 'اسم الطالب الكامل', 'الاسم الثلاثي',
    'student name', 'student full name', 'name'
  ),
  'company': ('السرية', 'رقم السرية', 'company', 'company number'),
  'platoon': ('الفصيل', 'رقم الفصيل', 'platoon', 'platoon number')
}

def normalize_roster_header(value):
  normalized = normalize_import_text(value)
  return ''.join(character for character in normalized if character.isalnum())

def find_roster_header(headers):
  normalized_headers = [normalize_roster_header(value) for value in headers]
  indexes = {}
  for field, aliases in ROSTER_HEADER_ALIASES.items():
    normalized_aliases = [normalize_roster_header(alias) for alias in aliases]
    match = next((
      index for index, header in enumerate(normalized_headers)
      if header and any(header == alias or header.startswith(alias) for alias in normalized_aliases)
    ), None)
    if match is None and field in ('student_name', 'company', 'platoon'):
      return None
    if match is not None:
      indexes[field] = match
  return indexes

def normalize_roster_cell(value, number_format=''):
  if value is None:
    return ''
  if isinstance(value, float) and value.is_integer():
    value = int(value)
  text = str(value).strip()
  if isinstance(value, int) and re.fullmatch(r'0+', number_format or ''):
    text = text.zfill(len(number_format))
  return text

def parse_student_sequence(value):
  text = normalize_roster_cell(value).translate(
    str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789')
  )
  if not text:
    return None
  if not re.fullmatch(r'\d+', text) or int(text) <= 0:
    raise ValueError('تسلسل الطالب يجب أن يكون رقمًا صحيحًا موجبًا')
  return int(text)

def make_roster_identifier(company, student_sequence):
  return f'C{int(company)}-S{int(student_sequence):04d}'

def parse_roster_xlsx(file_stream):
  if not HAS_OPENPYXL:
    raise RuntimeError('مكتبة قراءة ملفات XLSX وXLSM غير مثبتة')
  workbook = openpyxl.load_workbook(file_stream, read_only=True, data_only=True)
  empty_roster_sheet = None
  try:
    for sheet in workbook.worksheets:
      header_row = None
      indexes = None
      for row_number, cells in enumerate(sheet.iter_rows(min_row=1, max_row=20), start=1):
        indexes = find_roster_header([cell.value for cell in cells])
        if indexes:
          header_row = row_number
          break
      if indexes is None:
        continue

      students = []
      seen_sequences = set()
      next_sequences = {}
      for line_number, row in enumerate(sheet.iter_rows(min_row=header_row + 1), start=header_row + 1):
        cells_by_field = {
          field: row[index] if index < len(row) else None
          for field, index in indexes.items()
        }
        if not any(cell is not None and cell.value is not None and str(cell.value).strip()
                   for cell in cells_by_field.values()):
          continue
        student_name = normalize_roster_cell(
          cells_by_field['student_name'].value if cells_by_field['student_name'] else None
        )
        company_value = cells_by_field['company'].value if cells_by_field['company'] else None
        platoon = normalize_roster_cell(
          cells_by_field['platoon'].value if cells_by_field['platoon'] else None
        )
        company = normalize_company_number(company_value)
        if not student_name or company is None or not platoon:
          raise ValueError(
            f'بيانات الطالب في ورقة "{sheet.title}"، صف {line_number} ناقصة أو غير صحيحة'
          )
        sequence_cell = cells_by_field.get('student_sequence')
        student_sequence = parse_student_sequence(
          sequence_cell.value if sequence_cell else None
        )
        next_sequences[company] = next_sequences.get(company, 0) + 1
        if student_sequence is None:
          student_sequence = next_sequences[company]
        identity = (company, student_sequence)
        if identity in seen_sequences:
          raise ValueError(f'تسلسل الطالب مكرر في السرية {company}: {student_sequence}')
        seen_sequences.add(identity)
        students.append({
          'military_number': make_roster_identifier(company, student_sequence),
          'student_name': student_name,
          'company': company,
          'platoon': platoon,
          'student_sequence': student_sequence
        })
      if not students:
        empty_roster_sheet = sheet.title
        continue
      return students
  finally:
    workbook.close()
  if empty_roster_sheet:
    raise ValueError(f'عُثر على عناوين الأعمدة في ورقة "{empty_roster_sheet}" لكن لا توجد بيانات طلاب بعدها')
  raise ValueError(
    'لم يتم العثور على صف عناوين يتضمن اسم الطالب والسرية والفصيل '
    'في أول 20 صفاً من أي ورقة'
  )

def parse_roster_xls(file_stream):
  if not HAS_XLRD:
    raise RuntimeError('لقراءة ملفات XLS القديمة، ثبّت الاعتمادية xlrd')
  workbook = xlrd.open_workbook(file_contents=file_stream.read(), formatting_info=True)
  empty_roster_sheet = None
  for sheet in workbook.sheets():
    header_row = None
    indexes = None
    for row_number in range(min(sheet.nrows, 20)):
      indexes = find_roster_header(sheet.row_values(row_number))
      if indexes:
        header_row = row_number
        break
    if indexes is None:
      continue
    students = []
    seen_sequences = set()
    next_sequences = {}
    for row_number in range(header_row + 1, sheet.nrows):
      row = sheet.row(row_number)
      cells_by_field = {
        field: row[index] if index < len(row) else None
        for field, index in indexes.items()
      }
      if not any(cell is not None and cell.ctype != xlrd.XL_CELL_EMPTY and str(cell.value).strip()
                 for cell in cells_by_field.values()):
        continue
      student_name = normalize_roster_cell(
        cells_by_field['student_name'].value if cells_by_field['student_name'] else None
      )
      company_value = cells_by_field['company'].value if cells_by_field['company'] else None
      platoon = normalize_roster_cell(
        cells_by_field['platoon'].value if cells_by_field['platoon'] else None
      )
      company = normalize_company_number(company_value)
      excel_row = row_number + 1
      if not student_name or company is None or not platoon:
        raise ValueError(
          f'بيانات الطالب في ورقة "{sheet.name}"، صف {excel_row} ناقصة أو غير صحيحة'
        )
      sequence_cell = cells_by_field.get('student_sequence')
      student_sequence = parse_student_sequence(
        sequence_cell.value if sequence_cell else None
      )
      next_sequences[company] = next_sequences.get(company, 0) + 1
      if student_sequence is None:
        student_sequence = next_sequences[company]
      identity = (company, student_sequence)
      if identity in seen_sequences:
        raise ValueError(f'تسلسل الطالب مكرر في السرية {company}: {student_sequence}')
      seen_sequences.add(identity)
      students.append({
        'military_number': make_roster_identifier(company, student_sequence),
        'student_name': student_name,
        'company': company,
        'platoon': platoon,
        'student_sequence': student_sequence
      })
    if not students:
      empty_roster_sheet = sheet.name
      continue
    return students
  if empty_roster_sheet:
    raise ValueError(f'عُثر على عناوين الأعمدة في ورقة "{empty_roster_sheet}" لكن لا توجد بيانات طلاب بعدها')
  raise ValueError(
    'لم يتم العثور على صف عناوين يتضمن اسم الطالب والسرية والفصيل '
    'في أول 20 صفاً من أي ورقة'
  )

def parse_roster_workbook(file_stream, filename):
  extension = os.path.splitext(str(filename or '').lower())[1]
  if extension in ('.xlsx', '.xlsm'):
    file_data = file_stream.read()
    try:
      return parse_roster_xlsx(io.BytesIO(file_data))
    except ValueError as standard_error:
      try:
        return parse_bundled_roster(io.BytesIO(file_data), os.path.basename(filename))
      except ValueError:
        raise standard_error
  if extension == '.xls':
    return parse_roster_xls(file_stream)
  raise ValueError('الصيغ المدعومة لملف Excel هي XLSX وXLSM وXLS')

@app.route('/api/exams/roster', methods=['GET', 'POST'])
def exam_roster():
  if request.method == 'GET':
    conn = get_db()
    counts = {str(row['company']): row['count'] for row in conn.execute(
      'SELECT company, COUNT(*) AS count FROM exam_roster GROUP BY company'
    ).fetchall()}
    total = conn.execute('SELECT COUNT(*) FROM exam_roster').fetchone()[0]
    platoons = [row['platoon'] for row in conn.execute(
      'SELECT DISTINCT platoon FROM exam_roster ORDER BY platoon'
    ).fetchall()]
    conn.close()
    return jsonify({
      'counts': counts, 'total': total, 'platoons': platoons,
      'excel_available': HAS_OPENPYXL or HAS_XLRD
    })

  uploads = request.files.getlist('files')
  if not uploads:
    single_upload = request.files.get('file')
    if single_upload:
      uploads = [single_upload]
  uploads = [upload for upload in uploads if upload and upload.filename]
  if not uploads:
    return jsonify({'status': 'error', 'message': 'اختر ملف Excel بصيغة XLSX أو XLSM أو XLS'}), 400
  students = []
  seen_identifiers = set()
  try:
    for uploaded_file in uploads:
      extension = os.path.splitext(uploaded_file.filename.lower())[1]
      if extension not in ('.xlsx', '.xlsm', '.xls'):
        raise ValueError(f'صيغة الملف غير مدعومة: {uploaded_file.filename}')
      file_students = parse_roster_workbook(uploaded_file.stream, uploaded_file.filename)
      for student in file_students:
        if student['military_number'] in seen_identifiers:
          raise ValueError(
            f'تسلسل الطالب مكرر بين الملفات: السرية {student["company"]}، '
            f'التسلسل {student["student_sequence"]}'
          )
        seen_identifiers.add(student['military_number'])
        students.append(student)
  except (ValueError, RuntimeError, OSError) as error:
    return jsonify({'status': 'error', 'message': str(error)}), 400
  except Exception as error:
    return jsonify({'status': 'error', 'message': 'تعذرت قراءة ملفات Excel: ' + str(error)}), 400
  if not students:
    return jsonify({'status': 'error', 'message': 'لم يتم العثور على طلاب في الملفات المحددة'}), 422

  conn = get_db()
  try:
    with conn:
      conn.execute('DELETE FROM exam_roster')
      conn.executemany(
        '''INSERT INTO exam_roster
           (military_number, student_name, company, platoon, student_sequence)
           VALUES (?, ?, ?, ?, ?)''',
        [
          (s['military_number'], s['student_name'], s['company'], s['platoon'], s['student_sequence'])
          for s in students
        ]
      )
      ensure_student_qr_tokens(conn)
  except sqlite3.Error as error:
    conn.close()
    return jsonify({'status': 'error', 'message': 'تعذر حفظ قائمة الطلاب: ' + str(error)}), 500
  counts = {str(row['company']): row['count'] for row in conn.execute(
    'SELECT company, COUNT(*) AS count FROM exam_roster GROUP BY company'
  ).fetchall()}
  conn.close()
  auto_backup()
  return jsonify({
    'status': 'success', 'total': len(students), 'counts': counts,
    'files': [os.path.basename(upload.filename) for upload in uploads]
  })

@app.route('/api/exams/issues', methods=['GET', 'POST'])
def exam_issues():
  if request.method == 'GET':
    conn = get_db()
    rows = conn.execute(
      '''SELECT id, course, cohort, subject, platoon_filter, year, semester, exam_date, exam_time,
                print_confirmed, key_transferred, created_at
         FROM exam_issues ORDER BY created_at DESC'''
    ).fetchall()
    conn.close()
    return jsonify({'issues': [dict(row) for row in rows], 'qr_available': HAS_QRCODE})

  if not HAS_QRCODE:
    return jsonify({'status': 'error', 'message': 'مكتبة توليد QR غير مثبتة'}), 503
  data = request.get_json(silent=True) or {}
  required_text = ('course', 'subject', 'exam_date', 'exam_time')
  if any(not isinstance(data.get(field), str) or not data[field].strip() for field in required_text):
    return jsonify({'status': 'error', 'message': 'أكمل اسم الدورة والمادة والتاريخ ووقت الامتحان'}), 400
  try:
    datetime.strptime(data['exam_date'], '%Y-%m-%d')
    datetime.strptime(data['exam_time'], '%H:%M')
  except ValueError:
    return jsonify({'status': 'error', 'message': 'صيغة التاريخ أو وقت الامتحان غير صحيحة'}), 400
  cohort = str(data.get('cohort') or '')
  if cohort not in ('A', 'K', 'O'):
    return jsonify({'status': 'error', 'message': 'رمز الفوج يجب أن يكون A أو K أو O'}), 400
  try:
    year = int(data.get('year'))
    semester = int(data.get('semester'))
  except (TypeError, ValueError):
    return jsonify({'status': 'error', 'message': 'السنة أو النصف الدراسي غير صحيح'}), 400
  if year not in (1, 2, 3) or semester not in (1, 2):
    return jsonify({'status': 'error', 'message': 'السنة أو النصف الدراسي غير صحيح'}), 400
  counts = data.get('counts')
  model_a = data.get('modelA')
  model_b = data.get('modelB')
  if not isinstance(counts, dict) or not isinstance(model_a, list) or not model_a or not isinstance(model_b, list) or not model_b:
    return jsonify({'status': 'error', 'message': 'أصدر نموذجي الأسئلة A وB وأدخل أعداد السرايا'}), 400
  normalized_counts = {}
  for company in range(1, 5):
    count = counts.get(str(company), 0)
    if type(count) is int:
      normalized_counts[company] = count
    elif isinstance(count, str) and count.isdigit():
      normalized_counts[company] = int(count)
    else:
      return jsonify({'status': 'error', 'message': 'أعداد الطلاب يجب أن تكون أعداداً صحيحة'}), 400
  counts = normalized_counts
  if any(value < 0 for value in counts.values()) or sum(counts.values()) < 1 or sum(counts.values()) > 2000:
    return jsonify({'status': 'error', 'message': 'يجب إدخال عدد صحيح بين 0 و2000 دفتر'}), 400
  platoon_filter = str(data.get('platoon') or '').strip()

  conn = get_db()
  try:
    roster = conn.execute(
      '''SELECT military_number, student_name, company, platoon, student_sequence, qr_jwt
         FROM exam_roster ORDER BY company, student_sequence'''
    ).fetchall()
    if platoon_filter:
      roster = [student for student in roster if student['platoon'] == platoon_filter]
    if not roster:
      conn.close()
      return jsonify({'status': 'error', 'message': 'استورد قائمة الطلاب من Excel أولاً'}), 400
    by_company = {company: [] for company in range(1, 5)}
    for student in roster:
      by_company[student['company']].append(dict(student))
    for company, requested in counts.items():
      available = len(by_company[company])
      if requested > available:
        conn.close()
        return jsonify({
          'status': 'error',
          'message': f'السرية {company}: العدد المطلوب {requested} أكبر من عدد الطلاب في Excel ({available})'
        }), 400

    issue_id = secrets.token_urlsafe(18)
    with conn:
      conn.execute(
        '''INSERT INTO exam_issues
           (id, course, cohort, subject, platoon_filter, year, semester, exam_date, exam_time, model_a_json, model_b_json)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        (issue_id, data['course'].strip(), cohort, data['subject'].strip(), platoon_filter, year, semester,
         data['exam_date'].strip(), data['exam_time'].strip(),
         json.dumps(model_a, ensure_ascii=False), json.dumps(model_b, ensure_ascii=False))
      )
      booklets = []
      for company, count in counts.items():
        for position, student in enumerate(by_company[company][:count], start=1):
          student_sequence = student['student_sequence']
          model = 'A' if position % 2 else 'B'
          qr_token = secrets.token_urlsafe(24)
          conn.execute(
            '''INSERT INTO exam_booklets
               (qr_token, issue_id, military_number, qr_jwt, company, platoon, serial_number, model)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
            (qr_token, issue_id, student['military_number'], student['qr_jwt'],
             company, student['platoon'], student_sequence, model)
          )
          qr_data = make_exam_qr(
            conn, issue_id, cohort, company, student_sequence, model, qr_token,
            student['qr_jwt']
          )
          booklets.append({
            'qr_token': qr_token, 'qr_data': qr_data, 'military_number': student['military_number'],
            'company': company, 'platoon': student['platoon'],
            'student_sequence': student_sequence, 'serial_number': student_sequence, 'model': model
          })
  except sqlite3.Error as error:
    conn.close()
    return jsonify({'status': 'error', 'message': 'تعذر حفظ إصدار الامتحان: ' + str(error)}), 500
  except Exception as error:
    traceback.print_exc()
    conn.close()
    return jsonify({'status': 'error', 'message': 'تعذر إنشاء QR أو دفتر الامتحان: ' + str(error)}), 500
  conn.close()
  auto_backup()
  return jsonify({'status': 'success', 'issue_id': issue_id, 'booklets': booklets})

@app.route('/api/exams/issues/<issue_id>', methods=['GET'])
def get_exam_issue(issue_id):
  conn = get_db()
  issue = conn.execute('SELECT * FROM exam_issues WHERE id = ?', (issue_id,)).fetchone()
  if not issue:
    conn.close()
    return jsonify({'status': 'error', 'message': 'إصدار الامتحان غير موجود'}), 404
  booklets = conn.execute(
    '''SELECT qr_token, military_number, company, platoon, serial_number, model, status, score,
              review_json, responses_json, manual_scores_json
       FROM exam_booklets WHERE issue_id = ? ORDER BY company, serial_number''',
    (issue_id,)
  ).fetchall()
  conn.close()
  result = dict(issue)
  result['modelA'] = json.loads(result.pop('model_a_json'))
  result['modelB'] = json.loads(result.pop('model_b_json'))
  result['grading'] = json.loads(result.pop('grading_json'))
  result['booklets'] = []
  for row in booklets:
    booklet = dict(row)
    booklet['review'] = json.loads(booklet.pop('review_json') or '[]')
    booklet['responses'] = json.loads(booklet.pop('responses_json') or '{}')
    booklet['manual_scores'] = json.loads(booklet.pop('manual_scores_json') or '{}')
    result['booklets'].append(booklet)
  return jsonify({'issue': result})

@app.route('/api/exams/issues/<issue_id>/transfer-key', methods=['POST'])
def transfer_exam_answer_key(issue_id):
  data = request.get_json(silent=True) or {}
  if type(data.get('printed')) is not bool or type(data.get('transfer')) is not bool:
    return jsonify({'status': 'error', 'message': 'أكد الطباعة واختيار ترحيل المفتاح صراحةً'}), 400
  conn = get_db()
  issue = conn.execute('SELECT id FROM exam_issues WHERE id = ?', (issue_id,)).fetchone()
  if not issue:
    conn.close()
    return jsonify({'status': 'error', 'message': 'إصدار الامتحان غير موجود'}), 404
  conn.execute(
    'UPDATE exam_issues SET print_confirmed = ?, key_transferred = ? WHERE id = ?',
    (int(data['printed']), int(data['printed'] and data['transfer']), issue_id)
  )
  conn.commit()
  conn.close()
  return jsonify({'status': 'success', 'key_transferred': bool(data['printed'] and data['transfer'])})

@app.route('/api/exams/issues/<issue_id>/answer-key', methods=['POST'])
def import_exam_answer_key(issue_id):
  data = request.get_json(silent=True) or {}
  conn = get_db()
  issue = conn.execute('SELECT * FROM exam_issues WHERE id = ?', (issue_id,)).fetchone()
  if not issue:
    conn.close()
    return jsonify({'status': 'error', 'message': 'إصدار الامتحان غير موجود'}), 404
  corrected = conn.execute(
    "SELECT COUNT(*) FROM exam_booklets WHERE issue_id = ? AND status != 'issued'", (issue_id,)
  ).fetchone()[0]
  if corrected:
    conn.close()
    return jsonify({'status': 'error', 'message': 'لا يمكن تغيير المفتاح بعد بدء تصحيح أي دفتر'}), 409
  models = {}
  for model, column in (('A', 'model_a_json'), ('B', 'model_b_json')):
    raw_answers = data.get('model' + model)
    blocks = json.loads(issue[column])
    if not isinstance(raw_answers, list) or len(raw_answers) != len(blocks):
      conn.close()
      return jsonify({'status': 'error', 'message': f'مفتاح النموذج {model} لا يطابق مجموعات الأسئلة'}), 400
    for block_index, block_answers in enumerate(raw_answers):
      items = blocks[block_index].get('items') if isinstance(blocks[block_index], dict) else None
      if not isinstance(block_answers, list) or not isinstance(items, list) or len(block_answers) != len(items):
        conn.close()
        return jsonify({'status': 'error', 'message': f'عدد الإجابات في مجموعة {block_index + 1} من النموذج {model} غير مطابق'}), 400
      for item_index, answer in enumerate(block_answers):
        if not isinstance(answer, (str, int, float)) or not str(answer).strip():
          conn.close()
          return jsonify({'status': 'error', 'message': f'الإجابة {block_index + 1}:{item_index + 1} في النموذج {model} فارغة أو غير صحيحة'}), 400
        items[item_index]['answer'] = str(answer).strip()
    models[model] = json.dumps(blocks, ensure_ascii=False)
  conn.execute(
    'UPDATE exam_issues SET model_a_json = ?, model_b_json = ?, key_transferred = 1 WHERE id = ?',
    (models['A'], models['B'], issue_id)
  )
  conn.commit()
  conn.close()
  auto_backup()
  return jsonify({'status': 'success'})

@app.route('/api/exams/issues/<issue_id>/grading', methods=['POST'])
def save_exam_grading(issue_id):
  data = request.get_json(silent=True) or {}
  grading = data.get('grading')
  if not isinstance(grading, dict) or set(grading) != {'A', 'B'}:
    return jsonify({'status': 'error', 'message': 'بيانات توزيع الدرجات للنموذجين غير صحيحة'}), 400
  conn = get_db()
  issue = conn.execute('SELECT * FROM exam_issues WHERE id = ?', (issue_id,)).fetchone()
  if not issue:
    conn.close()
    return jsonify({'status': 'error', 'message': 'إصدار الامتحان غير موجود'}), 404
  scanned = conn.execute(
    "SELECT COUNT(*) FROM exam_booklets WHERE issue_id = ? AND status != 'issued'", (issue_id,)
  ).fetchone()[0]
  if scanned:
    conn.close()
    return jsonify({'status': 'error', 'message': 'لا يمكن تعديل توزيع الدرجات بعد بدء تصحيح أي دفتر'}), 409
  normalized = {}
  for model in ('A', 'B'):
    blocks = json.loads(issue['model_a_json'] if model == 'A' else issue['model_b_json'])
    raw_blocks = grading.get(model)
    if not isinstance(raw_blocks, list) or len(raw_blocks) != len(blocks):
      conn.close()
      return jsonify({'status': 'error', 'message': f'عدد مجموعات الدرجات في النموذج {model} لا يطابق الأسئلة'}), 400
    normalized_blocks = []
    for index, (raw, block) in enumerate(zip(raw_blocks, blocks)):
      if not isinstance(raw, dict) or not isinstance(raw.get('manual'), list):
        conn.close()
        return jsonify({'status': 'error', 'message': f'بيانات المجموعة {index + 1} غير صحيحة'}), 400
      try:
        points = float(raw.get('points'))
      except (TypeError, ValueError):
        conn.close()
        return jsonify({'status': 'error', 'message': f'أدخل درجة صحيحة للمجموعة {index + 1} في النموذج {model}'}), 400
      items = block.get('items') if isinstance(block, dict) else None
      if not math.isfinite(points) or points < 0 or not isinstance(items, list):
        conn.close()
        return jsonify({'status': 'error', 'message': 'الدرجات أو بنية الأسئلة غير صحيحة'}), 400
      manual = []
      for item_index in raw['manual']:
        if type(item_index) is not int or not 0 <= item_index < len(items):
          conn.close()
          return jsonify({'status': 'error', 'message': 'تحديد سؤال التصحيح اليدوي غير صحيح'}), 400
        manual.append(item_index)
      normalized_blocks.append({'points': points, 'manual': sorted(set(manual))})
    normalized[model] = normalized_blocks
  conn.execute('UPDATE exam_issues SET grading_json = ? WHERE id = ?',
               (json.dumps(normalized, ensure_ascii=False), issue_id))
  conn.commit()
  conn.close()
  auto_backup()
  return jsonify({'status': 'success'})

# صيغ ملفات المسح المقبولة: PDF أو صور بجميع الصيغ الشائعة
SCAN_PDF_EXTENSIONS = {'.pdf'}
SCAN_IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp', '.gif'}
SCAN_MAX_PAGES = 2000

def iter_scan_pages(filename, data):
  """يحوّل ملف المسح (PDF أو صورة بأي صيغة شائعة) إلى صفحات صور جاهزة للتحليل."""
  extension = os.path.splitext(filename or '')[1].lower()
  if extension in SCAN_PDF_EXTENSIONS:
    document = fitz.open(stream=data, filetype='pdf')
    try:
      if document.page_count < 1:
        raise ValueError('ملف PDF لا يحتوي على صفحات')
      for page_index in range(document.page_count):
        pixmap = document.load_page(page_index).get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width, pixmap.n)
        yield page_index + 1, image
    finally:
      document.close()
    return
  if extension in SCAN_IMAGE_EXTENSIONS:
    image = cv2.imdecode(np.frombuffer(data, dtype=np.uint8), cv2.IMREAD_COLOR)
    if image is None or image.size == 0:
      # احتياط: بعض الصيغ (مثل GIF القديمة) تُقرأ عبر محرك المستندات
      try:
        document = fitz.open(stream=data, filetype=extension.lstrip('.'))
      except (RuntimeError, ValueError, TypeError):
        document = None
      if document is None or document.page_count < 1:
        if document is not None:
          document.close()
        raise ValueError('تعذر فك ترميز الصورة؛ الملف تالف أو الصيغة غير مدعومة')
      pixmap = document.load_page(0).get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
      image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width, pixmap.n)
      document.close()
    yield 1, image
    return
  raise ValueError('صيغة غير مدعومة؛ المقبول: PDF أو صور (PNG, JPG, JPEG, BMP, TIFF, WEBP, GIF)')

def read_qr_from_image(image, detector=None):
  """محرك قراءة QR متعدد المراحل: معالجات وتكبيرات وتدويرات مختلفة حتى ينجح."""
  if detector is None:
    detector = cv2.QRCodeDetector()
  gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if len(image.shape) == 3 else image
  height, width = gray.shape[:2]
  variants = [gray]
  try:
    variants.append(cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1])
  except cv2.error:
    pass
  try:
    variants.append(cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 11))
  except cv2.error:
    pass
  # رمز الدفتر مطبوع أعلى يسار الورقة؛ الاقتصاص يركّز البحث ويحسّن دقته
  header_crop = gray[0:max(1, int(height * 0.4)), 0:max(1, int(width * 0.5))]
  variants.append(header_crop)
  variants.append(cv2.resize(gray, None, fx=1.6, fy=1.6, interpolation=cv2.INTER_CUBIC))
  variants.append(cv2.resize(header_crop, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC))

  for variant in variants:
    try:
      code, _, _ = detector.detectAndDecode(variant)
    except cv2.error:
      code = ''
    if code:
      return code
    try:
      multi = detector.detectAndDecodeMulti(variant)
    except cv2.error:
      multi = None
    if multi and multi[0] and len(multi) > 1 and multi[1] is not None:
      for candidate in multi[1]:
        if candidate:
          return candidate
  # محاولات أخيرة بالتدوير لالتقاط الأوراق الممسوحة مقلوبة أو مائلة 90 درجة
  for angle in (cv2.ROTATE_90_CLOCKWISE, cv2.ROTATE_180, cv2.ROTATE_90_COUNTERCLOCKWISE):
    try:
      code, _, _ = detector.detectAndDecode(cv2.rotate(gray, angle))
    except cv2.error:
      code = ''
    if code:
      return code
  # محرك احتياطي مستقل (pyzbar) يقوّي القراءة من المسح الرديء الجودة
  if HAS_PYZBAR:
    for variant in (gray, header_crop):
      try:
        results = pyzbar_decode(PILImage.fromarray(variant))
      except Exception:
        continue
      for result in results:
        text = result.data.decode('utf-8', 'ignore') if result.data else ''
        if text:
          return text
  return ''

_tesseract_cmd_cache = None

def tesseract_cmd():
  """يحدد مسار مشغّل Tesseract إن وُجد (مثبت نظامياً أو بجانب البرنامج)."""
  global _tesseract_cmd_cache
  if not HAS_TESSERACT:
    return None
  if _tesseract_cmd_cache is not None:
    return _tesseract_cmd_cache or None
  candidates = [
    os.environ.get('TESSERACT_CMD'),
    os.path.join(APP_DIR, 'Tesseract-OCR', 'tesseract.exe'),
    os.path.join(APP_DIR, 'tesseract', 'tesseract.exe'),
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
  ]
  for candidate in candidates:
    if candidate and os.path.isfile(candidate):
      pytesseract.pytesseract.tesseract_cmd = candidate
      _tesseract_cmd_cache = candidate
      return candidate
  try:
    pytesseract.get_tesseract_version()
    _tesseract_cmd_cache = pytesseract.pytesseract.tesseract_cmd
  except Exception:
    _tesseract_cmd_cache = ''
  return _tesseract_cmd_cache or None

ARABIC_INDIC_DIGITS = str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789')

def ocr_read_military_number(image):
  """قراءة احتياطية: يستخرج الرقم العسكري المطبوع أعلى الدفتر عند تعذر قراءة QR."""
  if tesseract_cmd() is None:
    return None
  gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if len(image.shape) == 3 else image
  height, width = gray.shape[:2]
  # سطر بيانات الطالب يقع ضمن ترويسة الدفتر أسفل رمز QR مباشرة ويتكرر مع كل صفحة
  header = gray[int(height * 0.06):int(height * 0.32), :]
  if header.size == 0:
    return None
  header = cv2.resize(header, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
  variants = [header]
  try:
    variants.append(cv2.threshold(header, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1])
  except cv2.error:
    pass
  config = '--psm 6 -c tessedit_char_whitelist=0123456789'
  found = []
  for variant in variants:
    try:
      text = pytesseract.image_to_string(variant, lang='eng', config=config) or ''
    except Exception:
      continue
    for run in re.findall(r'\d{3,}', text.translate(ARABIC_INDIC_DIGITS)):
      # استبعاد السنة المطبوعة ضمن التاريخ حتى لا تُقرأ رقماً عسكرياً
      if len(run) == 4 and run[:2] in ('19', '20'):
        continue
      found.append(run)
  if not found:
    return None
  return max(found, key=len)

def find_booklet_by_military_number(conn, military_number):
  """يطابق رقماً عسكرياً مقروءاً بالـ OCR مع دفتر وحيد غير مصحح."""
  rows = conn.execute(
    '''SELECT b.*, i.model_a_json, i.model_b_json, i.grading_json, i.key_transferred, i.cohort
       FROM exam_booklets b JOIN exam_issues i ON i.id = b.issue_id
       WHERE b.military_number = ? AND b.status != 'corrected'
       ORDER BY b.rowid DESC''',
    (military_number,)
  ).fetchall()
  return rows[0] if len(rows) == 1 else None

def measure_bubble_darkness(gray, x, y, radius):
  """يقيس امتلاء الفقاعة نسبةً إلى محيطها الورقي، فيصمد مع اختلاف أجهزة المسح."""
  height, width = gray.shape[:2]
  outer = int(radius * 1.8) + 2
  x0, x1 = max(0, x - outer), min(width, x + outer + 1)
  y0, y1 = max(0, y - outer), min(height, y + outer + 1)
  window = gray[y0:y1, x0:x1]
  if window.size == 0:
    return 0.0
  ys, xs = np.mgrid[y0:y1, x0:x1]
  distance = np.sqrt((xs - x) ** 2 + (ys - y) ** 2)
  inner_mask = distance <= max(2.0, radius * 0.65)
  ring_mask = (distance >= radius * 1.2) & (distance <= radius * 1.7)
  if not inner_mask.any() or not ring_mask.any():
    return 0.0
  inner_pixels = window[inner_mask].astype(np.float32)
  ring_mean = float(window[ring_mask].mean())
  if ring_mean < 40.0:
    return 0.0
  inner_mean = float(inner_pixels.mean())
  fill = float((inner_pixels < max(70.0, ring_mean - 55.0)).mean())
  contrast = min(1.0, max(0.0, (ring_mean - inner_mean) / 95.0))
  return 0.65 * fill + 0.35 * contrast

def collect_hough_circles(gray):
  """يكشف الدوائر بعدة حساسيات ويدمج النتائج لالتقاط الفقاعات الباهتة والواضحة."""
  found = []
  for param2 in (24, 19, 15):
    try:
      circles = cv2.HoughCircles(
        gray, cv2.HOUGH_GRADIENT, dp=1.2, minDist=16, param1=100,
        param2=param2, minRadius=5, maxRadius=26
      )
    except cv2.error:
      circles = None
    if circles is not None:
      found.extend(np.round(circles[0]).astype(int).tolist())
  return found

def collect_template_circles(gray):
  """احتياط عند ضعف كشف Hough: مطابقة قالب حلقة دائرية في منطقة عمود الفقاعات."""
  height, width = gray.shape[:2]
  offset_x = int(width * 0.55)
  region = 255 - gray[:, offset_x:]
  if region.size == 0:
    return []
  detections = []
  for radius in (12, 15, 18, 21):
    size = radius * 2 + 9
    template = np.zeros((size, size), dtype=np.uint8)
    cv2.circle(template, (size // 2, size // 2), radius, 255, 2)
    try:
      result = cv2.matchTemplate(region, template, cv2.TM_CCOEFF_NORMED)
    except cv2.error:
      continue
    locations = np.where(result >= 0.55)
    for py, px in zip(*locations):
      detections.append((int(px + offset_x + size // 2), int(py + size // 2), radius, float(result[py, px])))
  detections.sort(key=lambda item: -item[3])
  kept = []
  for x, y, radius, _ in detections[:600]:
    if all((x - kx) ** 2 + (y - ky) ** 2 > radius ** 2 for kx, ky, _ in kept):
      kept.append((x, y, radius))
  return kept

def merge_circle_detections(detections):
  """يدمج الكشوفات المتكررة للفقاعة الواحدة القادمة من أكثر من كاشف."""
  kept = []
  for x, y, radius in sorted(detections, key=lambda item: -item[2]):
    if all((x - kx) ** 2 + (y - ky) ** 2 > max(9.0, radius * 0.7) ** 2 for kx, ky, _ in kept):
      kept.append((x, y, radius))
  return kept

def detect_page_bubbles(image):
  gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) if len(image.shape) == 3 else image
  blurred = cv2.medianBlur(gray, 5)
  circles = collect_hough_circles(blurred)
  if len(circles) < 6:
    circles = circles + collect_template_circles(gray)
  circles = merge_circle_detections(circles)
  if not circles:
    return []
  height, width = gray.shape[:2]
  candidates = []
  for x, y, radius in circles:
    if x < width * 0.55 or x > width * 0.995 or y < height * 0.02 or y > height * 0.99:
      continue
    darkness = measure_bubble_darkness(gray, x, y, radius)
    candidates.append((int(y), int(x), int(radius), darkness))
  if not candidates:
    return []
  # عمود الفقاعات قد ينحرف مع ميلان الورقة؛ نطابق خطاً مستقيماً بدل موضع أفقي ثابت
  buckets = {}
  for entry in candidates:
    key = int(round(entry[1] / 12.0))
    buckets.setdefault(key, []).append(entry)
  best = max(buckets.values(), key=len)
  column_ys = np.array([entry[0] for entry in best], dtype=np.float64)
  column_xs = np.array([entry[1] for entry in best], dtype=np.float64)
  if len(best) >= 3:
    slope, intercept = np.polyfit(column_ys, column_xs, 1)
  else:
    slope, intercept = 0.0, float(column_xs.mean())
  aligned = [entry for entry in candidates if abs(entry[1] - (slope * entry[0] + intercept)) <= 16.0]
  if len(aligned) < max(2, len(best) // 2):
    aligned = list(best)
  aligned.sort(key=lambda entry: entry[0])
  merged = []
  for entry in aligned:
    if merged and entry[0] - merged[-1][0] <= max(8.0, entry[2] * 0.6):
      if entry[3] > merged[-1][3]:
        merged[-1] = entry
    else:
      merged.append(entry)
  return [
    {'marked': darkness >= 0.45, 'uncertain': 0.20 <= darkness < 0.45, 'y': y}
    for y, _, _, darkness in merged
  ]

def split_bubble_groups(detected, option_counts):
  """يوزع الفقاعات المقروءة على أسئلة النموذج؛ يستعين بفواصل الأسئلة عند اختلاف العدد."""
  total = sum(option_counts)
  if len(detected) == total:
    groups = []
    cursor = 0
    for count in option_counts:
      groups.append(detected[cursor:cursor + count])
      cursor += count
    return groups
  question_indexes = [index for index, count in enumerate(option_counts) if count > 0]
  if not question_indexes or len(detected) < len(question_indexes):
    return None
  ys = [bubble['y'] for bubble in detected]
  gaps = sorted(
    ((ys[index + 1] - ys[index], index) for index in range(len(ys) - 1)),
    reverse=True
  )
  split_points = sorted(index + 1 for _, index in gaps[:len(question_indexes) - 1])
  segments = []
  start = 0
  for point in split_points + [len(detected)]:
    segments.append(detected[start:point])
    start = point
  if len(segments) != len(question_indexes):
    return None
  if any(len(segment) != option_counts[index] for segment, index in zip(segments, question_indexes)):
    return None
  groups = [[] for _ in option_counts]
  for segment, index in zip(segments, question_indexes):
    groups[index] = segment
  return groups

@app.route('/api/exams/scan', methods=['POST'])
def scan_exam_pdfs():
  if not HAS_SCAN_LIBS:
    return jsonify({'status': 'error', 'message': 'مكتبات قراءة ملفات المسح غير مثبتة'}), 503
  uploads = request.files.getlist('files')
  single_upload = request.files.get('file')
  if single_upload and single_upload.filename:
    uploads.append(single_upload)
  uploads = [upload for upload in uploads if upload and upload.filename]
  if not uploads:
    return jsonify({'status': 'error', 'message': 'اختر ملفات المسح: PDF أو صور بأي صيغة (PNG, JPG, BMP, TIFF, WEBP)'}), 400

  conn = get_db()
  signing_secret = qr_signing_secret(conn)
  conn.commit()
  detector = cv2.QRCodeDetector()
  pages_by_token = {}
  problems = []
  page_index = 0
  limit_reached = False
  multiple_sources = len(uploads) > 1
  for uploaded in uploads:
    if limit_reached:
      break
    source = os.path.basename(uploaded.filename)
    prefix = f'{source} - ' if multiple_sources else ''
    try:
      for page_number, image in iter_scan_pages(uploaded.filename, uploaded.read()):
        page_index += 1
        if page_index > SCAN_MAX_PAGES:
          problems.append({'page': page_index, 'message': f'{prefix}تجاوز الحد الأقصى للمسح ({SCAN_MAX_PAGES} صفحة)؛ أوقف بقية الملفات'})
          limit_reached = True
          break
        try:
          identified_by = 'qr'
          code = read_qr_from_image(image, detector)
          student_payload = decode_student_jwt(code, student_jwt_secret(conn)) if code else None
          payload = decode_exam_qr(code, signing_secret) if code and not student_payload else None
          if student_payload:
            row = find_booklet_by_student_jwt(conn, code)
            if not row or (
              row['military_number'] != student_payload.get('sub') or
              row['company'] != student_payload.get('company') or
              row['platoon'] != student_payload.get('platoon')
            ):
              problems.append({'page': page_index, 'message': prefix + 'رمز JWT لا يطابق دفترًا مسجلاً'})
              continue
          elif payload:
            row = conn.execute(
              '''SELECT b.*, i.model_a_json, i.model_b_json, i.grading_json, i.key_transferred, i.cohort
                 FROM exam_booklets b JOIN exam_issues i ON i.id = b.issue_id
                 WHERE b.qr_token = ? AND b.issue_id = ?''',
              (payload.get('k'), payload.get('i'))
            ).fetchone()
            if not row or (
              row['company'] != payload.get('c') or row['serial_number'] != payload.get('s') or
              row['model'] != payload.get('m') or row['cohort'] != payload.get('f') or
              (payload.get('j') and row['qr_jwt'] != payload.get('j'))
            ):
              problems.append({'page': page_index, 'message': prefix + 'رمز QR لا يطابق دفترًا مسجلاً'})
              continue
          else:
            # تعذر QR: محاولة التعرف على الدفتر من الرقم العسكري المطبوع عبر OCR
            row = None
            military_number = ocr_read_military_number(image)
            if military_number:
              row = find_booklet_by_military_number(conn, military_number)
            if row is None:
              problems.append({'page': page_index, 'message': prefix + 'لم يُقرأ QR صالح من الصفحة ولم يتعرف OCR على رقم عسكري مطابق'})
              continue
            identified_by = 'ocr'
          if row['status'] == 'corrected':
            problems.append({'page': page_index, 'message': prefix + 'هذا الدفتر مصحح ومعتمد مسبقاً'})
            continue
          if not row['key_transferred']:
            problems.append({'page': page_index, 'message': prefix + 'لم يعتمد المستخدم مفتاح الإجابة لهذا الإصدار'})
            continue
          page_bubbles = detect_page_bubbles(image)
          preview = cv2.resize(image, (720, int(image.shape[0] * 720 / image.shape[1])))
          encoded_preview, preview_bytes = cv2.imencode('.jpg', preview, [cv2.IMWRITE_JPEG_QUALITY, 45])
          pages_by_token.setdefault(row['qr_token'], []).append({
            'page': page_index, 'bubbles': page_bubbles, 'booklet': dict(row),
            'identified_by': identified_by,
            'preview': base64.b64encode(preview_bytes).decode('ascii') if encoded_preview else ''
          })
        except (RuntimeError, ValueError, TypeError, cv2.error) as error:
          problems.append({'page': page_index, 'message': prefix + 'تعذر تحليل الصفحة: ' + str(error)})
    except (RuntimeError, ValueError, TypeError) as error:
      problems.append({'page': page_index, 'message': f'تعذر فتح الملف {source}: {error}'})

  results = []
  preview_pages_left = 100
  for token, pages in pages_by_token.items():
    pages.sort(key=lambda page: page['page'])
    booklet = pages[0]['booklet']
    model = booklet['model']
    blocks = json.loads(booklet['model_a_json'] if model == 'A' else booklet['model_b_json'])
    grading = json.loads(booklet['grading_json'] or '{}').get(model, [])
    if len(grading) != len(blocks):
      problems.append({'page': pages[0]['page'], 'message': f'لم تُحفظ درجات النموذج {model} كاملة؛ لم يُعتمد هذا الدفتر'})
      continue
    detected = [bubble for page in pages for bubble in page['bubbles']]
    required = []
    for question in exam_questions(blocks):
      item = question['item']
      block_index = question['block_index']
      item_index = question['item_index']
      block_grading = grading[block_index] if block_index < len(grading) else {}
      manual = item_index in block_grading.get('manual', []) or item.get('type') == 'essay'
      options = item.get('options')
      if item.get('type') == 'tf' and (not isinstance(options, list) or not options):
        options = ['صح', 'خطأ']
      option_count = len(options) if isinstance(options, list) else 0
      if item.get('type') == 'tf' and option_count == 0:
        option_count = 2
      manual = manual or option_count == 0
      required.append({
        **question, 'manual': manual,
        'points': float(block_grading.get('points', 0)) / max(1, len(blocks[block_index].get('items', []))),
        'option_count': option_count
      })

    option_counts = [question['option_count'] for question in required]
    groups = split_bubble_groups(detected, option_counts)
    bubble_stream_valid = groups is not None
    if groups is None:
      groups = [[] for _ in required]
      problems.append({
        'page': pages[0]['page'],
        'message': (
          f'الدفتر تسلسل {booklet["serial_number"]} (السرية {booklet["company"]}): '
          f'عدد الفقاعات المقروءة {len(detected)} لا يطابق المتوقع {sum(option_counts)}؛ أُحيلت إجاباته للمراجعة اليدوية'
        )
      })
    responses = {}
    review = []
    for question, group in zip(required, groups):
      if question['manual']:
        review.append({
          'question_key': question['key'], 'prompt': str(question['item'].get('text') or ''),
          'reason': 'يتطلب هذا السؤال تصحيحاً يدوياً', 'manual_score': True,
          'points': question['points']
        })
        continue
      key = question['key']
      selected = [index for index, bubble in enumerate(group) if bubble['marked']] if bubble_stream_valid else []
      multi = question['item'].get('type') in ('multi_select', 'multi')
      if not bubble_stream_valid:
        reason = 'تعذر مطابقة عدد فقاعات الصفحة مع أسئلة النموذج؛ راجع الإجابة'
      elif any(bubble.get('uncertain') for bubble in group):
        reason = 'درجة التظليل غير واضحة'
      elif len(selected) == 0:
        reason = 'لم يظهر تظليل واضح'
      elif not multi and len(selected) > 1:
        reason = 'تم تظليل أكثر من اختيار'
      else:
        expected = exam_answer_indexes(question['item'])
        if not expected:
          review.append({
            'question_key': key, 'prompt': str(question['item'].get('text') or ''),
            'reason': 'الإجابة النموذجية غير محددة أو لا تطابق الخيارات',
            'manual_score': False, 'points': question['points'], 'selected': selected
          })
          responses[key] = {'selected': selected, 'correct': False, 'points': question['points']}
          continue
        correct = bool(expected) and selected == expected
        responses[key] = {'selected': selected, 'correct': correct, 'points': question['points']}
        continue
      review.append({
        'question_key': key, 'prompt': str(question['item'].get('text') or ''),
        'reason': reason, 'manual_score': False, 'points': question['points'],
        'selected': selected
      })
      responses[key] = {'selected': selected, 'correct': False, 'points': question['points']}

    conn.execute(
      "UPDATE exam_booklets SET responses_json = ?, review_json = ?, score = NULL, status = 'review' WHERE qr_token = ?",
      (json.dumps(responses, ensure_ascii=False), json.dumps(review, ensure_ascii=False), token)
    )
    page_images = []
    for page in pages:
      if not page['preview'] or preview_pages_left <= 0:
        continue
      page_images.append({'page': page['page'], 'data': 'data:image/jpeg;base64,' + page['preview']})
      preview_pages_left -= 1
      if len(page_images) >= 10:
        break
    results.append({
      'qr_token': token, 'military_number': booklet['military_number'],
      'company': booklet['company'], 'serial_number': booklet['serial_number'],
      'model': model, 'review': review, 'status': 'review',
      'identified_by': 'ocr' if any(page.get('identified_by') == 'ocr' for page in pages) else 'qr',
      'page_images': page_images
    })
  conn.commit()
  conn.close()
  auto_backup()
  return jsonify({
    'status': 'success', 'booklets': results, 'page_problems': problems,
    'engines': {'qr': True, 'qr_backup': HAS_PYZBAR, 'ocr': tesseract_cmd() is not None}
  })

@app.route('/api/exams/booklets/<qr_token>/finalize', methods=['POST'])
def finalize_exam_booklet(qr_token):
  data = request.get_json(silent=True) or {}
  decisions = data.get('decisions')
  manual_scores = data.get('manual_scores')
  if not isinstance(decisions, dict) or not isinstance(manual_scores, dict):
    return jsonify({'status': 'error', 'message': 'أرسل قرارات المراجعة ودرجات الأسئلة اليدوية'}), 400
  conn = get_db()
  booklet = conn.execute(
    '''SELECT b.*, i.model_a_json, i.model_b_json, i.grading_json
       FROM exam_booklets b JOIN exam_issues i ON i.id = b.issue_id WHERE b.qr_token = ?''',
    (qr_token,)
  ).fetchone()
  if not booklet:
    conn.close()
    return jsonify({'status': 'error', 'message': 'الدفتر غير موجود'}), 404
  if booklet['status'] != 'review':
    conn.close()
    return jsonify({'status': 'error', 'message': 'امسح الدفتر وراجع إجاباته قبل اعتماد الدرجة'}), 409
  model = booklet['model']
  blocks = json.loads(booklet['model_a_json'] if model == 'A' else booklet['model_b_json'])
  grading = json.loads(booklet['grading_json'] or '{}').get(model, [])
  responses = json.loads(booklet['responses_json'] or '{}')
  review = json.loads(booklet['review_json'] or '[]')
  score = 0.0
  for question in exam_questions(blocks):
    key = question['key']
    block_index = question['block_index']
    item_index = question['item_index']
    block_grading = grading[block_index] if block_index < len(grading) else {}
    points = float(block_grading.get('points', 0)) / max(1, len(blocks[block_index].get('items', [])))
    item = question['item']
    is_manual = item_index in block_grading.get('manual', []) or item.get('type') == 'essay'
    if is_manual:
      try:
        awarded = float(manual_scores[key])
      except (KeyError, TypeError, ValueError):
        conn.close()
        return jsonify({'status': 'error', 'message': f'أدخل درجة السؤال اليدوي {key}'}), 400
      if not math.isfinite(awarded) or not 0 <= awarded <= points:
        conn.close()
        return jsonify({'status': 'error', 'message': f'درجة السؤال {key} يجب أن تكون بين 0 و{points:g}'}), 400
      score += awarded
    elif key in {entry['question_key'] for entry in review}:
      if type(decisions.get(key)) is not bool:
        conn.close()
        return jsonify({'status': 'error', 'message': f'حدد السؤال {key} صحيحاً أو خاطئاً في المراجعة'}), 400
      if decisions[key]:
        score += points
    elif responses.get(key, {}).get('correct'):
      score += points
  conn.execute(
    "UPDATE exam_booklets SET manual_scores_json = ?, score = ?, status = 'corrected' WHERE qr_token = ?",
    (json.dumps(manual_scores, ensure_ascii=False), score, qr_token)
  )
  conn.commit()
  conn.close()
  auto_backup()
  return jsonify({'status': 'success', 'score': score})

@app.route('/api/exams/issues/<issue_id>/grades.xlsx', methods=['GET'])
def export_exam_grades(issue_id):
  if not HAS_OPENPYXL:
    return jsonify({'status': 'error', 'message': 'مكتبة تصدير Excel غير مثبتة'}), 503
  conn = get_db()
  issue = conn.execute('SELECT course, subject, exam_date FROM exam_issues WHERE id = ?', (issue_id,)).fetchone()
  if not issue:
    conn.close()
    return jsonify({'status': 'error', 'message': 'إصدار الامتحان غير موجود'}), 404
  booklets = conn.execute(
    '''SELECT military_number, company, platoon, serial_number, model, score, status,
              responses_json, manual_scores_json
       FROM exam_booklets WHERE issue_id = ? ORDER BY company, serial_number''',
    (issue_id,)
  ).fetchall()
  conn.close()
  incomplete = sum(row['status'] != 'corrected' for row in booklets)
  if incomplete:
    return jsonify({'status': 'error', 'message': f'لم يكتمل تصحيح {incomplete} دفتر؛ لا يمكن إصدار سجل الدرجات'}), 409
  workbook = openpyxl.Workbook()
  sheet = workbook.active
  sheet.title = 'الدرجات'
  sheet.append([
    'الدورة', 'المادة', 'التاريخ', 'الرقم العسكري', 'السرية', 'الفصيل',
    'تسلسل الدفتر', 'النموذج', 'الدرجة النهائية', 'تفاصيل إجابات الأسئلة',
    'درجات التصحيح اليدوي'
  ])
  for row in booklets:
    sheet.append([
      issue['course'], issue['subject'], issue['exam_date'], row['military_number'],
      row['company'], row['platoon'], row['serial_number'], row['model'], row['score'],
      row['responses_json'], row['manual_scores_json']
    ])
  sheet.freeze_panes = 'A2'
  sheet.auto_filter.ref = sheet.dimensions
  for column in ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'):
    sheet.column_dimensions[column].width = 18
  sheet.column_dimensions['J'].width = 50
  sheet.column_dimensions['K'].width = 30
  output = io.BytesIO()
  workbook.save(output)
  output.seek(0)
  filename = f"درجات_{issue['subject']}_{issue['exam_date']}.xlsx"
  return send_file(
    output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    as_attachment=True, download_name=filename
  )

# ==========================================
# الواجهة الأمامية الشاملة (HTML/CSS/JS)
# ==========================================

HTML_TEMPLATE = """
<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<meta charset="UTF-8">
<title>برنامج امتحانات طلاب الكلية العسكرية</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html,body{width:100%;height:100%;margin:0;}
body{font-family:"Segoe UI",Tahoma,Arial,sans-serif;direction:rtl;background:radial-gradient(circle at top, #223d4e 0%, #121d2b 28%, #0c1521 55%, #070d15 100%);color:#edf3f7;min-height:100vh;height:100vh;height:100dvh;overflow:hidden;}
:root{
  --navy:#172236;--navy-dark:#0c1521;--navy-light:#243a52;
  --olive:#1e332b;--olive-deep:#12231c;--steel:#d2d9df;--steel-soft:#a7b4be;
  --gold:#d7b15a;--gold-deep:#b68a2d;--gold-soft:#f4d889;--surface:#1a2b40;
  --surface2:#223a57;--border:#3b516d;--text:#edf3f7;--text-muted:#a7b7c9;
  --good:#66d99a;--warn:#f3c969;--bad:#ef6b6b;--radius:10px;--shadow:0 12px 30px rgba(0,0,0,.35);
}
.app-wrapper{position:relative;z-index:1;display:flex;flex-direction:column;min-height:100vh;height:100vh;height:100dvh;overflow:hidden;}
.app-header{background:linear-gradient(135deg,#0d1620 0%,#162b35 35%,#1d3446 100%);border-bottom:2px solid var(--gold);padding:.8rem 1.5rem;display:grid;grid-template-columns:minmax(0,1fr) auto minmax(0,1fr);align-items:center;gap:1rem;flex-shrink:0;box-shadow:0 10px 24px rgba(0,0,0,.25);}
.app-header-title{grid-column:2;grid-row:1;text-align:center;min-width:0;position:relative;padding-inline:2rem;}
.app-header-title::before,.app-header-title::after{content:"";position:absolute;top:50%;width:1.25rem;height:1px;background:linear-gradient(90deg,transparent,var(--gold));}
.app-header-title::before{inset-inline-start:0;}
.app-header-title::after{inset-inline-end:0;transform:scaleX(-1);}
.app-header-title h1{font-family:'Traditional Arabic','Times New Roman',serif;font-size:1.8rem;font-weight:900;color:var(--gold-soft);letter-spacing:0;line-height:1.4;text-shadow:0 2px 3px rgba(0,0,0,.55),0 0 18px rgba(215,177,90,.2);}
.app-header-title p{font-size:.75rem;color:var(--text-muted);margin-top:.1rem;}
.app-header-logo{grid-column:3;grid-row:1;justify-self:end;width:76px;height:76px;display:grid;place-items:center;padding:4px;background:#fff;border:1px solid rgba(244,216,137,.9);border-radius:10px;box-shadow:0 6px 18px rgba(0,0,0,.3),0 0 16px rgba(215,177,90,.16);}
.app-header-logo img{display:block;width:100%;height:100%;object-fit:contain;border-radius:6px;}
.app-header-stats{grid-column:1;grid-row:1;justify-self:start;display:flex;gap:1rem;}
.hstat{text-align:center;min-width:80px;background:linear-gradient(180deg, rgba(255,255,255,.03), rgba(18,35,28,.4));padding:.45rem .6rem;border:1px solid rgba(215,177,90,.35);border-radius:var(--radius);box-shadow:inset 0 1px 0 rgba(255,255,255,.06);}
.hstat-val{font-size:1.1rem;font-weight:700;color:var(--gold);}
.hstat-lbl{font-size:.65rem;color:var(--text-muted);}
.nav-tabs{display:flex;background:linear-gradient(180deg,#152639,#0f1a2a);border-bottom:1px solid rgba(215,177,90,.35);overflow-x:auto;flex-shrink:0;scrollbar-width:thin;}
.nav-tab{padding:.8rem 1.2rem;font-size:.875rem;font-weight:700;color:var(--steel-soft);cursor:pointer;border-bottom:3px solid transparent;white-space:nowrap;transition:all .2s ease;letter-spacing:.02em;}
.nav-tab:hover{color:var(--text);background:rgba(255,255,255,.02);}
.nav-tab.active{color:var(--gold-soft);border-bottom-color:var(--gold);background:linear-gradient(180deg, rgba(215,177,90,.14), rgba(255,255,255,.02));box-shadow:inset 0 -2px 0 rgba(215,177,90,.5);}
.tab-content{position:relative;isolation:isolate;display:none;padding:1rem 1.25rem 1rem;flex:1;overflow-x:hidden;overflow-y:auto;min-height:0;}
.tab-content.active{display:flex;flex-direction:column;}
.tab-content.active::before{content:"";position:fixed;inset:110px 0 0;z-index:0;background-image:url('/logo-watermark');background-repeat:no-repeat;background-position:center calc(70% + 24px);background-size:min(130vw,140vh,1200px);opacity:.40;mix-blend-mode:screen;filter:drop-shadow(0 0 12px rgba(226,64,255,.84)) drop-shadow(0 0 32px rgba(85,95,255,.70));pointer-events:none;}
.tab-content > *{position:relative;z-index:1;}
.card{background:linear-gradient(180deg,rgba(21,34,48,.88),rgba(15,24,35,.9));border:1px solid rgba(215,177,90,.42);border-radius:var(--radius);padding:1rem;margin-bottom:.8rem;box-shadow:var(--shadow), inset 0 1px 0 rgba(255,255,255,.04);}
.card-title{font-size:.9rem;font-weight:800;color:var(--gold-soft);margin-bottom:1rem;padding-bottom:.5rem;border-bottom:1px solid rgba(215,177,90,.28);letter-spacing:.02em;}
.grid-2{display:grid;grid-template-columns:repeat(2,1fr);gap:1rem;}
.grid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:1rem;}
label.lbl{display:block;font-size:.75rem;color:var(--text-muted);margin-bottom:.35rem;}
select,input[type=number],input[type=text],textarea{width:100%;padding:.62rem .75rem;background:rgba(9,17,24,.85);border:1px solid rgba(184,208,226,.25);border-radius:var(--radius);color:var(--text);font:inherit;direction:rtl;transition:border-color .2s ease, box-shadow .2s ease, transform .15s ease;}
select:focus,input:focus,textarea:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px rgba(215,177,90,.18);}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:.4rem;padding:.68rem 1rem;border-radius:var(--radius);font:inherit;font-weight:700;cursor:pointer;border:1px solid rgba(215,177,90,.45);background:linear-gradient(180deg,#2c3f53,#1d2a3a);color:var(--text);transition:transform .15s ease, filter .15s ease, box-shadow .15s ease;box-shadow:inset 0 1px 0 rgba(255,255,255,.08), 0 6px 16px rgba(0,0,0,.18);letter-spacing:.01em;}
.btn:hover{transform:translateY(-1px);filter:brightness(1.08);}
.btn-primary{background:linear-gradient(135deg,#ebd08d,#c69b39 35%,#a97922);color:#111a1a;border-color:rgba(245,232,182,.8);box-shadow:0 8px 18px rgba(182,138,45,.28), inset 0 1px 0 rgba(255,255,255,.25);}
.btn-outline{background:rgba(15,24,35,.28);border-color:rgba(215,177,90,.7);color:var(--gold-soft);}
.btn-danger{background:rgba(110,23,31,.12);border-color:rgba(239,107,107,.7);color:var(--bad);}
.btn-edit{background:rgba(90,71,18,.14);border-color:rgba(243,201,105,.72);color:var(--warn);}
.kpi{text-align:center;padding:1rem;background:linear-gradient(180deg,#1a2c3f,#172437);border:1px solid rgba(215,177,90,.38);border-radius:var(--radius);box-shadow:var(--shadow), inset 0 1px 0 rgba(255,255,255,.05);}
.kpi-val{font-size:2rem;font-weight:800;color:var(--gold-soft);}
.kpi-lbl{font-size:.7rem;color:var(--text-muted);}
#toast{position:fixed;bottom:1.5rem;left:50%;transform:translateX(-50%);background:#1a2744;border:1px solid var(--gold);color:var(--text);padding:.65rem 1.25rem;border-radius:var(--radius);opacity:0;transition:all .25s;z-index:9999;pointer-events:none;box-shadow:var(--shadow);}
#toast.show{opacity:1;}
.q-item{padding:.8rem .9rem;background:linear-gradient(180deg, rgba(18,28,40,.9), rgba(9,17,24,.95));border:1px solid rgba(215,177,90,.2);margin-bottom:.5rem;border-radius:var(--radius);display:flex;justify-content:space-between;align-items:center;gap:1rem;box-shadow:inset 0 1px 0 rgba(255,255,255,.03);}
.q-ans{color:var(--good);font-size:.8rem;margin-top:.2rem;}
.type-stat-tag{font-size:.7rem;padding:.18rem .5rem;border-radius:999px;background:rgba(36,52,96,.85);color:var(--gold);margin-right:.4rem;font-weight:bold;}
.manual-topic-list{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:.45rem;margin-top:.6rem;}
.manual-topic-option,.manual-question-row{display:flex;align-items:flex-start;gap:.6rem;padding:.7rem .8rem;background:rgba(9,17,24,.55);border:1px solid rgba(215,177,90,.22);border-radius:var(--radius);}
.manual-topic-option{align-items:center;cursor:pointer;color:var(--text);}
.manual-question-groups{display:grid;gap:.8rem;}
.manual-type-group{padding:.75rem;border:1px solid rgba(215,177,90,.2);border-radius:var(--radius);background:rgba(9,17,24,.25);}
.manual-type-title{color:var(--gold-soft);font-size:.85rem;margin-bottom:.55rem;}
.manual-question-list{display:grid;gap:.4rem;}
.manual-question-row{cursor:pointer;}
.manual-question-row input,.manual-topic-option input{flex:0 0 auto;margin-top:.18rem;accent-color:var(--gold);}
.manual-question-text{display:grid;gap:.2rem;}
.manual-question-thumbnail{max-width:120px;max-height:90px;object-fit:contain;align-self:start;}
.manual-question-meta{color:var(--text-muted);font-size:.72rem;}
.manual-actions{display:flex;flex-wrap:wrap;gap:.5rem;margin-top:.8rem;}
.manual-preview-actions{display:flex;gap:.5rem;flex-wrap:wrap;margin-bottom:.8rem;}
.answer-key{margin-top:1rem;padding:1rem;background:#fff;color:#111;border:1px solid #b58a32;border-radius:var(--radius);direction:rtl;}
.answer-key section{margin-top:.8rem;}
.answer-key section:first-of-type{margin-top:0;}
.answer-key h3{margin:.2rem 0 .5rem;color:#6b4b12;}
.answer-key-group{margin:.5rem 0;padding:.5rem;border-top:1px solid #d8c18a;}
.answer-key-group div{margin-top:.25rem;}
.word-import-controls{display:flex;align-items:center;gap:.6rem;flex-wrap:wrap;}
.word-import-status{min-height:1.4rem;margin-top:.5rem;color:var(--text-muted);font-size:.78rem;}
.word-import-summary{margin:.5rem 0;color:var(--gold-soft);font-size:.82rem;}
.word-import-row{padding:.8rem 0;border-top:1px solid rgba(215,177,90,.25);}
.word-import-row-head{display:flex;align-items:center;gap:.5rem;margin-bottom:.5rem;font-weight:700;color:var(--gold-soft);}
.word-import-fields{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.5rem;}
.word-import-field{display:grid;gap:.25rem;min-width:0;}
.word-import-field span{font-size:.7rem;color:var(--text-muted);}
.word-import-field.wide{grid-column:1/-1;}
.word-import-field textarea{min-height:3.2rem;resize:vertical;}
.word-import-warning{margin-top:.35rem;color:var(--warn);font-size:.72rem;}
.exam-forms-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:1rem;align-items:start;}
.exam-form-wrap{min-width:0;overflow-wrap:anywhere;background:#fff;color:#000;padding:1.5rem;font-family:'Traditional Arabic','Times New Roman',serif;direction:rtl;border-radius:var(--radius);border:1px solid #ccc;}
.exam-question-image{display:block;max-width:100%;max-height:360px;object-fit:contain;margin:.5rem auto;}
.image-size-control{display:grid;gap:.25rem;margin-top:.5rem;max-width:360px;}
.image-size-control input[type=range]{width:100%;accent-color:var(--gold);}
.exam-form-header{text-align:center;border-bottom:2px solid #000;padding-bottom:.5rem;margin-bottom:.8rem;font-family:'Traditional Arabic','Times New Roman',serif;}
.exam-header-center h2,.exam-header-center h3,.exam-header-center p{font-family:inherit;}
.exam-header-meta{display:flex;justify-content:flex-start;gap:1rem;direction:ltr;text-align:left;font-size:.9rem;margin-top:.45rem;}
.exam-header-meta span{direction:rtl;}
.exam-booklet{width:100%;border-collapse:collapse;background:#fff;color:#000;font-family:'Traditional Arabic','Times New Roman',serif;direction:rtl;}
.exam-booklet thead{display:table-header-group;}
.exam-booklet th,.exam-booklet td{position:relative;padding:0;vertical-align:top;text-align:right;font-weight:normal;}
.exam-booklet thead th{padding-top:0;}
.exam-booklet tbody tr{break-inside:avoid;page-break-inside:avoid;}
.exam-booklet-qr{display:none;}
.exam-issue-screen-preview .exam-booklet-qr{display:block;position:absolute;top:8px;left:8px;width:96px;height:96px;object-fit:contain;}
.exam-preview-modal{position:fixed;inset:0;z-index:10000;display:none;align-items:center;justify-content:center;padding:1rem;background:rgba(0,0,0,.78);}
.exam-preview-modal.open{display:flex;}
.exam-preview-dialog{display:flex;flex-direction:column;width:min(1200px,100%);height:min(94vh,1100px);background:#162437;border:1px solid var(--gold);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;}
.exam-preview-heading,.exam-preview-controls{display:flex;align-items:center;justify-content:space-between;gap:.7rem;padding:.75rem 1rem;border-bottom:1px solid rgba(215,177,90,.3);}
.exam-preview-heading h2{font-size:1rem;color:var(--gold-soft);}
.exam-preview-controls{flex-wrap:wrap;}
.exam-preview-controls select{flex:1 1 280px;max-width:620px;}
.exam-preview-sheet{flex:1;min-height:0;overflow:auto;margin:.75rem;background:#fff;color:#000;padding:14mm 12mm;font-family:'Traditional Arabic','Times New Roman',serif;direction:rtl;}
.exam-preview-sheet .exam-booklet thead th{padding-top:30mm;}
.exam-preview-sheet .exam-booklet-qr{display:block;position:absolute;top:0;left:0;width:25mm;height:25mm;object-fit:contain;}
.exam-preview-sheet .exam-form-header{min-height:25mm;}
.exam-preview-sheet .exam-booklet tbody tr{break-inside:avoid;page-break-inside:avoid;}
.exam-preview-sheet .omr-bubble{print-color-adjust:exact;-webkit-print-color-adjust:exact;}
.exam-booklet-question{margin:.65rem 0;break-inside:avoid;page-break-inside:avoid;}
.omr-options{margin:.25rem 1.5rem .2rem 0;display:grid;gap:.08rem;}
.omr-option{min-height:24px;display:flex;align-items:center;gap:.35rem;direction:rtl;}
.omr-bubble{display:inline-block;width:15px;height:15px;flex:0 0 15px;border:1.4px solid #000;border-radius:50%;background:#fff;}
.omr-option-text{font-size:.92rem;}
.exam-issue-status,.exam-correction-status{min-height:1.5rem;margin-top:.45rem;color:var(--gold-soft);font-size:.82rem;}
.exam-scoring-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:.8rem;}
.exam-scoring-block{padding:.8rem;border:1px solid rgba(215,177,90,.3);border-radius:var(--radius);margin-top:.5rem;}
.exam-scoring-question{display:flex;align-items:flex-start;gap:.5rem;margin:.4rem 0;font-size:.8rem;}
.exam-review-card{padding:.8rem;margin:.6rem 0;border:1px solid rgba(215,177,90,.35);border-radius:var(--radius);background:rgba(9,17,24,.55);}
.exam-review-card img{display:block;max-width:min(100%,720px);max-height:520px;object-fit:contain;background:#fff;margin:.5rem 0;}
.exam-review-choice{display:flex;gap:.5rem;align-items:center;margin:.4rem 0;}
#print-exam-preview{display:none;}
#booklet-print{display:none;}
@media (max-width:760px){.word-import-fields{grid-template-columns:repeat(2,minmax(0,1fr));}}
@media print {
  html, body { height: auto !important; overflow: visible !important; }
  body * { visibility: hidden; }
  body > * { display: none !important; }
  #print-exam-preview, #print-exam-preview * { visibility: visible !important; }
  #print-exam-preview { display: block !important; width: 100%; overflow: visible !important; }
  #booklet-print, #booklet-print * { visibility: visible !important; }
  #booklet-print { display: block !important; width: 100%; overflow: visible !important; }
  #booklet-print .exam-booklet { display:table !important; width:calc(100% - 20mm) !important; margin:0 10mm; break-after:page; page-break-after:always; }
  #booklet-print .exam-booklet thead th{padding-top:36mm;}
  #booklet-print .exam-booklet:last-child { break-after:auto; page-break-after:auto; }
  #booklet-print .exam-booklet-qr { display:block !important; position:absolute; top:4mm; left:0; width:30mm; height:30mm; object-fit:contain; }
  #booklet-print .exam-booklet-question { break-inside:avoid; page-break-inside:avoid; }
  #booklet-print .omr-bubble { width:15px; height:15px; border:1.4px solid #000; }
  #print-exam-preview .print-model-pair { display: block; break-after: page; page-break-after: always; }
  #print-exam-preview .print-model-pair:last-child { break-after: auto; page-break-after: auto; }
  #print-exam-preview .print-model-pair .exam-form-wrap { display: block; break-after: page; page-break-after: always; border: none; padding: 0; margin-bottom: 2rem; }
  #print-exam-preview .print-model-pair .answer-key { display: block; break-before: auto; page-break-before: auto; border: none; border-radius: 0; }
}
@media (max-width: 720px) {
  .tab-content.active::before{inset:100px 0 0;background-size:min(82vw,48vh,440px);background-position:center 68%;opacity:.1;}
  .app-header{grid-template-columns:minmax(0,1fr);gap:.55rem;padding:.65rem .8rem;}
  .app-header-title{grid-column:1;grid-row:1;padding-inline:.25rem;}
  .app-header-title h1{font-size:1.2rem;}
  .app-header-title::before,.app-header-title::after{display:none;}
  .app-header-logo{grid-column:1;grid-row:2;justify-self:center;width:62px;height:62px;}
  .app-header-stats{grid-column:1;grid-row:3;justify-self:center;gap:.45rem;}
  .hstat{min-width:68px;padding:.35rem .45rem;}
}
</style>
</head>
<body>
<div class="app-wrapper">
  <header class="app-header no-print">
    <div class="app-header-title">
      <h1>برنامج امتحانات طلاب الكلية العسكرية</h1>
      <p>منظومة إدارة قواعد البيانات وإصدار الامتحانات العسكرية</p>
    </div>

    <div class="tab-content" id="tab-data-dashboard">
      <div class="card no-print">
        <div class="card-title">قاعدة بيانات JWT</div>
        <div class="manual-actions">
          <input type="search" id="jwt-search" placeholder="ابحث بالرقم العسكري أو اسم الطالب أو JWT">
          <button class="btn btn-primary" onclick="loadJwtDatabase()">استعراض</button>
        </div>
        <div id="jwt-summary" class="exam-issue-status"></div>
        <div id="jwt-table" style="overflow:auto;margin-top:.75rem"></div>
      </div>
      <div class="card no-print">
        <div class="card-title">ذاكرة الأسئلة حسب الفوج</div>
        <div id="cohort-memory-summary" class="grid-3"></div>
        <div class="manual-actions" style="margin-top:.75rem">
          <select id="memory-reset-cohort"></select>
          <button class="btn btn-outline" onclick="resetCohortMemory()">تصفير ذاكرة الفوج</button>
          <button class="btn btn-primary" onclick="backupQuestions()">حفظ نسخة احتياطية</button>
        </div>
        <div id="cohort-history" style="overflow:auto;margin-top:.75rem"></div>
      </div>
    </div>
    <div class="app-header-logo">
      <img src="/logo" alt="شعار الكلية العسكرية الأولى" fetchpriority="high">
    </div>
    <div class="app-header-stats">
      <div class="hstat"><div class="hstat-val" id="hdr-total">0</div><div class="hstat-lbl">إجمالي</div></div>
      <div class="hstat"><div class="hstat-val" style="color:var(--good)" id="hdr-unused">0</div><div class="hstat-lbl">متاحة</div></div>
      <div class="hstat"><div class="hstat-val" style="color:var(--warn)" id="hdr-used">0</div><div class="hstat-lbl">مستخدمة</div></div>
    </div>
  </header>

  <nav class="nav-tabs no-print">
    <div class="nav-tab active" onclick="showTab('home')">الرئيسية</div>
    <div class="nav-tab" onclick="showTab('entry')">إدخال وتعديل الأسئلة</div>
    <div class="nav-tab" onclick="showTab('generate')">توليد الامتحان</div>
    <div class="nav-tab" onclick="showTab('manual')">اختيار يدوي</div>
    <div class="nav-tab" onclick="showTab('issue-paper')">إصدار ورقة الامتحان</div>
    <div class="nav-tab" onclick="showTab('grade-exams')">تصحيح الدفاتر الامتحانية</div>
    <div class="nav-tab" onclick="showTab('stats')">الإحصائيات</div>
    <div class="nav-tab" onclick="showTab('data-dashboard')">قاعدة JWT وذاكرة الأفواج</div>
  </nav>

  <!-- TAB: HOME -->
  <div class="tab-content active" id="tab-home">
    <div class="grid-4" style="margin-bottom:1.5rem">
      <div class="kpi"><div class="kpi-val" id="kpi-total">0</div><div class="kpi-lbl">إجمالي الأسئلة</div></div>
      <div class="kpi"><div class="kpi-val" style="color:var(--good)" id="kpi-unused">0</div><div class="kpi-lbl">غير مستخدمة</div></div>
      <div class="kpi"><div class="kpi-val" style="color:var(--warn)" id="kpi-used">0</div><div class="kpi-lbl">مستخدمة</div></div>
      <div class="kpi"><div class="kpi-val" style="color:#60a5fa">12</div><div class="kpi-lbl">المواد الدراسية</div></div>
    </div>

    <div class="card">
      <div class="card-title">إدارة العدادات والبحث الشامل</div>
      <div style="display:flex;gap:1rem;margin-bottom:1rem">
        <button class="btn btn-outline" onclick="resetAllUsed()">إعادة تعيين العدادات (تصفير المستخدم والمتاح)</button>
        <button class="btn btn-primary" onclick="backupQuestions()">نسخ احتياطي</button>
        <button class="btn btn-outline" onclick="restoreQuestionsBackup()">استيراد الأسئلة من النسخة الاحتياطية</button>
      </div>
      <div style="margin-top:1rem">
        <label class="lbl">البحث الشامل في نص الأسئلة أو الإجابات الصحيحة</label>
        <div style="display:flex;gap:.5rem">
          <input type="text" id="search-input" placeholder="اكتب كلمة أو جملة للبحث عنها في بنك الأسئلة...">
          <button class="btn btn-primary" onclick="performSearch()">بحث</button>
        </div>
        <div id="search-results" style="margin-top:1rem"></div>
      </div>
    </div>
  </div>

  <!-- TAB: ENTRY -->
  <div class="tab-content" id="tab-entry">
    <div class="card no-print" id="word-import-card">
      <div class="card-title">استيراد أسئلة من مستند Word</div>
      <div class="word-import-controls">
        <input type="file" id="word-import-file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onchange="resetWordImportPreview()">
        <button type="button" class="btn btn-outline" onclick="previewWordImport()">قراءة المستند ومعاينة الأسئلة</button>
      </div>
      <div class="word-import-status" id="word-import-status" role="status"></div>
      <div id="word-import-preview" style="display:none">
        <div class="word-import-summary" id="word-import-summary"></div>
        <div class="manual-actions">
          <button type="button" class="btn btn-outline" id="word-import-details-toggle" onclick="toggleWordImportDetails()">مراجعة وتعديل التفاصيل</button>
        </div>
        <div id="word-import-rows" style="display:none"></div>
        <div class="manual-actions">
          <button type="button" class="btn btn-primary" id="word-import-submit" onclick="importSelectedWordQuestions()">ترحيل الأسئلة الجاهزة</button>
          <button type="button" class="btn btn-outline" onclick="clearWordImportPreview()">إلغاء المعاينة</button>
        </div>
      </div>
    </div>
    <div class="grid-2">
      <div class="card">
        <div class="card-title" id="entry-card-title">إدخال سؤال جديد</div>
        <input type="hidden" id="ent-edit-id" value="">
        <div class="grid-2">
          <div><label class="lbl">المادة الدراسية</label><select id="ent-subject" onchange="loadAvailableQuestions()"></select></div>
          <div><label class="lbl">المفردة (الموضوع)</label><input type="text" id="ent-topic" placeholder="مثال: التعبية"></div>
        </div>
        <div class="grid-2" style="margin-top:.5rem">
          <div><label class="lbl">السنة الدراسية</label><select id="ent-year"><option value="1">الأولى</option><option value="2">الثانية</option><option value="3">الثالثة</option></select></div>
          <div><label class="lbl">النصف الدراسي</label><select id="ent-sem"><option value="1">الأول</option><option value="2">الثاني</option></select></div>
        </div>
        <div style="margin-top:.5rem">
          <label class="lbl">نوع السؤال الرئيسي</label>
          <select id="ent-type" onchange="toggleEntryFields()">
            <option value="mc">1. اختيار من متعدد</option>
            <option value="tf">2. صح وخطأ</option>
            <option value="fill">3. ملء الفراغات</option>
            <option value="essay">4. سؤال شرح</option>
            <option value="multi_select">5. اختيار الإجابات الصحيحة فقط (متعدد الخيارات)</option>
          </select>
        </div>
        <div style="margin-top:.5rem">
          <label class="lbl">نص السؤال الفرعي</label>
          <textarea id="ent-text" rows="3" placeholder="أكتب نص السؤال هنا..."></textarea>
        </div>
        <div style="margin-top:.75rem">
          <label class="lbl" for="ent-image">صورة مرفقة بالسؤال (PNG أو JPEG، حتى 5 ميغابايت)</label>
          <input type="file" id="ent-image" accept="image/png,image/jpeg" onchange="loadQuestionImage(this)">
          <div id="ent-image-preview-wrap" style="display:none;margin-top:.5rem">
            <img id="ent-image-preview" alt="معاينة صورة السؤال" style="display:block;max-width:100%;max-height:220px;object-fit:contain;margin-bottom:.5rem">
            <button type="button" class="btn btn-outline" onclick="removeQuestionImage()">إزالة الصورة</button>
          </div>
          <div id="ent-image-size-control" class="image-size-control" style="display:none">
            <label class="lbl" for="ent-image-size">حجم الصورة في السؤال: <span id="ent-image-size-value">100%</span></label>
            <input type="range" id="ent-image-size" min="20" max="100" step="5" value="100" oninput="updateQuestionImageSize(this.value)">
          </div>
        </div>

        <div id="field-mc" class="type-fields" style="margin-top:.5rem">
          <label class="lbl">الخيارات الفرعية (مشرّرة أبجدياً)</label>
          <div style="display:flex;gap:.3rem;align-items:center;margin-bottom:.2rem"><span>أ-</span><input type="text" id="mc-a" placeholder="الخيار الأول"></div>
          <div style="display:flex;gap:.3rem;align-items:center;margin-bottom:.2rem"><span>ب-</span><input type="text" id="mc-b" placeholder="الخيار الثاني"></div>
          <div style="display:flex;gap:.3rem;align-items:center;margin-bottom:.2rem"><span>جـ-</span><input type="text" id="mc-c" placeholder="الخيار الثالث"></div>
          <div style="display:flex;gap:.3rem;align-items:center;"><span>د-</span><input type="text" id="mc-d" placeholder="الخيار الرابع"></div>
        </div>

        <div id="field-tf" class="type-fields" style="display:none;margin-top:.5rem">
          <label class="lbl">تحديد الخيار الصحيح</label>
          <select id="tf-ans"><option value="صح">صح ✓</option><option value="خطأ">خطأ ✗</option></select>
        </div>

        <div id="field-fill" class="type-fields" style="display:none;margin-top:.5rem">
          <label class="lbl">اختيارات سؤال الفراغ (تُطبع فقاعات تحت السؤال)</label>
          <div style="display:flex;gap:.3rem;align-items:center;margin-bottom:.2rem"><span>أ-</span><input type="text" id="fill-a" placeholder="الخيار الأول"></div>
          <div style="display:flex;gap:.3rem;align-items:center;margin-bottom:.2rem"><span>ب-</span><input type="text" id="fill-b" placeholder="الخيار الثاني"></div>
          <div style="display:flex;gap:.3rem;align-items:center;margin-bottom:.2rem"><span>جـ-</span><input type="text" id="fill-c" placeholder="الخيار الثالث"></div>
          <div style="display:flex;gap:.3rem;align-items:center"><span>د-</span><input type="text" id="fill-d" placeholder="الخيار الرابع"></div>
        </div>

        <div id="field-multi" class="type-fields" style="display:none;margin-top:.5rem">
          <label class="lbl">الخيارات الأبجدية المتاحة (حدّد المربع أمام الخيارات الصحيحة)</label>
          <div id="multi-options-container"></div>
          <button type="button" class="btn btn-outline" style="margin-top:.4rem" onclick="addMultiOptionField()">+ إضافة خيار أبجدي جديد</button>
        </div>

        <div style="margin-top:.75rem">
          <label class="lbl" style="color:var(--gold);font-weight:bold">الإجابة الصحيحة / نموذج الإجابة</label>
          <input type="text" id="ent-answer" placeholder="أدخل رمز الإجابة الصحيحة أو نصها الكامل لتسهيل التصحيح...">
        </div>

        <div style="display:flex;gap:.5rem;margin-top:1rem">
          <button class="btn btn-primary" onclick="saveQuestion()">حفظ السؤال</button>
          <button class="btn btn-outline" onclick="cancelEdit()" id="btn-cancel-edit" style="display:none">إلغاء التعديل</button>
          <button class="btn btn-outline" onclick="loadAvailableQuestions()">عرض الأسئلة المتاحة</button>
        </div>
      </div>

      <div class="card">
        <div class="card-title">قائمة الأسئلة المتاحة للمادة المحددة</div>
        <div id="available-questions-list"></div>
      </div>
    </div>
  </div>

  <!-- TAB: GENERATE -->
  <div class="tab-content" id="tab-generate">
    <div class="card no-print">
      <div class="card-title">إعدادات إصدار الامتحانات العسكرية</div>
      <div class="grid-2">
        <div><label class="lbl">المادة الدراسية</label><select id="gen-subject" onchange="updateAvailableCounts()"></select></div>
        <div><label class="lbl">الفوج المعني</label><select id="gen-cohort" onchange="changeQuestionCohort(this.value)"></select></div>
        <div>
          <label class="lbl">السنة الدراسية</label>
          <select id="gen-year" onchange="updateAvailableCounts()">
            <option value="1">سنة أولى</option>
            <option value="2">سنة ثانية</option>
            <option value="3">سنة ثالثة</option>
            <option value="all">شامل (كل السنين والأحضان)</option>
            <option value="all1">شامل1 (السنة الأولى كلاً)</option>
            <option value="all2">شامل2 (السنة الثانية كلاً)</option>
            <option value="all3">شامل3 (السنة الثالثة كلاً)</option>
          </select>
        </div>
      </div>
      <div class="grid-2" style="margin-top:.5rem">
        <div><label class="lbl">النصف الدراسي</label><select id="gen-sem" onchange="updateAvailableCounts()"><option value="1">النصف الأول</option><option value="2">النصف الثاني</option></select></div>
        <div><label class="lbl">حالة الأسئلة</label><select id="gen-pool" onchange="updateAvailableCounts()"><option value="unused" selected>غير مستخدم فقط</option><option value="used">مستخدم فقط</option><option value="mix">مكس (مستخدم وغير مستخدم)</option></select></div>
      </div>
      <div style="margin-top:.5rem">
        <label class="lbl">ترتيب الأسئلة والخيارات داخل النموذج ب</label>
        <select id="gen-order">
          <option value="random">ترتيب عشوائي (عشوائية كاملة للأسئلة والخيارات داخل النموذج ب بدقة جذرياً)</option>
          <option value="seq">ترتيب حسب التسلسل الثابت</option>
        </select>
      </div>
      <div style="margin-top:.5rem;max-width:260px">
        <label class="lbl" for="gen-duration">مدة الامتحان</label>
        <select id="gen-duration" onchange="syncExamDuration('gen-duration')">
          <option value="1">ساعة واحدة</option><option value="1.5">ساعة ونصف</option>
          <option value="2" selected>ساعتان</option><option value="3">ثلاث ساعات</option><option value="4">أربع ساعات</option>
        </select>
      </div>

      <div style="margin-top:1rem">
        <label class="lbl">تحديد أعداد الأسئلة المطلوبة (العداد يظهر المتاح/الكلي بالمقابلة):</label>
        <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:.8rem;margin-top:.6rem">
          <div>
            <label class="lbl">
              <input type="checkbox" id="chk-mc" checked> س1: اختيار من متعدد
              <span class="type-stat-tag" id="stat-mc">[ متاح: 0 | كلي: 0 ]</span>
            </label>
            <input type="number" id="num-mc" value="5" min="0">
          </div>
          <div>
            <label class="lbl">
              <input type="checkbox" id="chk-tf" checked> س2: صح وخطأ
              <span class="type-stat-tag" id="stat-tf">[ متاح: 0 | كلي: 0 ]</span>
            </label>
            <input type="number" id="num-tf" value="5" min="0">
          </div>
          <div>
            <label class="lbl">
              <input type="checkbox" id="chk-fill" checked> س3: ملء الفراغات
              <span class="type-stat-tag" id="stat-fill">[ متاح: 0 | كلي: 0 ]</span>
            </label>
            <input type="number" id="num-fill" value="5" min="0">
          </div>
          <div>
            <label class="lbl">
              <input type="checkbox" id="chk-essay" checked> س4: سؤال شرح ومقالي
              <span class="type-stat-tag" id="stat-essay">[ متاح: 0 | كلي: 0 ]</span>
            </label>
            <input type="number" id="num-essay" value="2" min="0">
          </div>
          <div>
            <label class="lbl">
              <input type="checkbox" id="chk-multi" checked> س5: اختيار الإجابات الصحيحة
              <span class="type-stat-tag" id="stat-multi_select">[ متاح: 0 | كلي: 0 ]</span>
            </label>
            <input type="number" id="num-multi" value="2" min="0">
          </div>
        </div>
      </div>

      <div style="display:flex;gap:.5rem;margin-top:1.2rem">
        <button class="btn btn-primary" onclick="generateExam()">اصدار الأسئلة</button>
        <button class="btn btn-outline" onclick="restoreLastIssueForGenerated()">إعادة تعديل آخر إصدار</button>
        <button class="btn btn-outline" onclick="showAndPrintExamModels()">طباعة الأسئلة</button>
        <button class="btn btn-outline" onclick="exportToWord()">تصدير الأسئلة إلى مستند WORD</button>
        <button class="btn btn-outline" onclick="toggleAnswerKey('generated-answer-key')">عرض الإجابات النموذجية</button>
      </div>
    </div>

    <div class="card no-print" id="generated-issued-editor" style="display:none">
      <div class="card-title">تعديل أسئلة آخر إصدار</div>
      <div class="manual-question-groups" id="generated-issued-list"></div>
      <div class="manual-actions">
        <button class="btn btn-outline" onclick="selectGeneratedIssuedQuestions(true)">تحديد كل الأسئلة</button>
        <button class="btn btn-outline" onclick="selectGeneratedIssuedQuestions(false)">إلغاء التحديد</button>
        <button class="btn btn-primary" onclick="generateExam()">إصدار الأسئلة بعد التعديل</button>
      </div>
    </div>

    <div id="exam-preview" style="display:none">
      <div class="card-title no-print">معاينة النموذجين (نموذج أ ونموذج ب)</div>
      <div class="exam-forms-grid" id="exam-container"></div>
      <div id="generated-answer-key" class="answer-key no-print" style="display:none"></div>
    </div>
  </div>

  <!-- TAB: MANUAL EXAM -->
  <div class="tab-content" id="tab-manual">
    <div class="card no-print">
      <div class="card-title">اختيار أسئلة الامتحان يدويًا</div>
      <div class="grid-2">
        <div><label class="lbl">المادة الدراسية</label><select id="manual-subject" onchange="manualEditingIssue=null; manualEditingSelectedIds=null; loadManualTopics()"></select></div>
        <div><label class="lbl">الفوج المعني</label><select id="manual-cohort" onchange="changeQuestionCohort(this.value)"></select></div>
      </div>
      <div style="margin-top:.5rem;max-width:260px">
        <label class="lbl" for="manual-duration">مدة الامتحان</label>
        <select id="manual-duration" onchange="syncExamDuration('manual-duration')">
          <option value="1">ساعة واحدة</option><option value="1.5">ساعة ونصف</option>
          <option value="2" selected>ساعتان</option><option value="3">ثلاث ساعات</option><option value="4">أربع ساعات</option>
        </select>
      </div>
      <label class="lbl" style="margin-top:.8rem">حدد المفردات التي تريد الاختيار منها</label>
      <div id="manual-topic-list" class="manual-topic-list"></div>
      <div class="manual-actions">
        <button class="btn btn-outline" onclick="selectManualTopics(true)">تحديد كل المفردات</button>
        <button class="btn btn-outline" onclick="selectManualTopics(false)">إلغاء تحديد المفردات</button>
        <button class="btn btn-primary" onclick="showManualQuestions()">عرض الأسئلة</button>
        <button class="btn btn-outline" onclick="restoreLastIssueForManual()">إعادة تعديل آخر إصدار</button>
      </div>
    </div>

    <div class="card no-print" id="manual-question-panel" style="display:none">
      <div class="card-title">اختر الأسئلة المطلوبة حسب النوع</div>
      <div id="manual-question-list" class="manual-question-groups"></div>
      <div class="manual-actions">
        <button class="btn btn-outline" onclick="selectManualQuestions(true)">تحديد كل الأسئلة</button>
        <button class="btn btn-outline" onclick="selectManualQuestions(false)">إلغاء التحديد</button>
        <button class="btn btn-primary" onclick="previewManualExam()">معاينة النموذجين</button>
      </div>
    </div>

    <div id="manual-exam-preview" style="display:none">
      <div class="card-title no-print">معاينة النموذجين (نموذج أ ونموذج ب)</div>
      <div class="manual-preview-actions no-print">
        <button class="btn btn-primary" onclick="showAndPrintExamModels()">طباعة النموذجين</button>
        <button class="btn btn-outline" onclick="exportToWord()">تصدير إلى مستند Word</button>
        <button class="btn btn-outline" onclick="toggleAnswerKey('manual-answer-key')">عرض الإجابات النموذجية</button>
      </div>
      <div class="exam-forms-grid" id="manual-exam-container"></div>
      <div id="manual-answer-key" class="answer-key no-print" style="display:none"></div>
    </div>
  </div>

  <!-- TAB: STATS -->
  <div class="tab-content" id="tab-stats">
    <div class="card">
      <div class="card-title">إحصائيات الأسئلة والمواد دراسياً</div>
      <div id="stats-table"></div>
    </div>
  </div>

  <div class="tab-content" id="tab-issue-paper">
    <div class="card no-print">
      <div class="card-title">إصدار ورقة الامتحان</div>
      <div class="grid-2">
        <div><label class="lbl" for="issue-course">اسم الدورة (إدخال يدوي)</label><input id="issue-course" type="text" placeholder="اسم الدورة"></div>
        <div><label class="lbl" for="issue-cohort">الفوج</label><select id="issue-cohort"><option value="A">A</option><option value="K">K</option><option value="O">O</option></select></div>
        <div><label class="lbl" for="issue-year">السنة الدراسية</label><select id="issue-year"><option value="1">الأولى</option><option value="2">الثانية</option><option value="3">الثالثة</option></select></div>
        <div><label class="lbl" for="issue-semester">النصف الدراسي</label><select id="issue-semester"><option value="1">الأول</option><option value="2">الثاني</option></select></div>
        <div><label class="lbl" for="issue-subject">المادة الامتحانية</label><select id="issue-subject"></select></div>
        <div><label class="lbl" for="issue-platoon">الفصيل (من Excel)</label><select id="issue-platoon"><option value="">كل الفصائل</option></select></div>
        <div><label class="lbl" for="issue-date">تاريخ الامتحان</label><input id="issue-date" type="date"></div>
        <div><label class="lbl" for="issue-time">وقت الامتحان</label><input id="issue-time" type="time"></div>
      </div>
      <div class="card-title" style="margin-top:1rem">قائمة الطلاب</div>
      <div class="word-import-controls">
        <input type="file" id="exam-roster-file" multiple accept=".xlsx,.xlsm,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel">
        <button type="button" class="btn btn-outline" onclick="importExamRoster()">استيراد ملفات Excel</button>
      </div>
      <div class="manual-question-meta">يمكن اختيار ملف واحد أو عدة ملفات. الصيغ المدعومة: XLSX وXLSM وXLS. يتم استبدال قائمة الطلاب الحالية بعد التحقق من الملفات.</div>
      <div id="exam-roster-summary" class="exam-issue-status">لم تُستورد قائمة طلاب بعد.</div>
      <div class="grid-4" style="margin-top:.8rem">
        <div><label class="lbl" for="issue-count-1">عدد دفاتر السرية 1</label><input id="issue-count-1" type="number" min="0" value="0"></div>
        <div><label class="lbl" for="issue-count-2">عدد دفاتر السرية 2</label><input id="issue-count-2" type="number" min="0" value="0"></div>
        <div><label class="lbl" for="issue-count-3">عدد دفاتر السرية 3</label><input id="issue-count-3" type="number" min="0" value="0"></div>
        <div><label class="lbl" for="issue-count-4">عدد دفاتر السرية 4</label><input id="issue-count-4" type="number" min="0" value="0"></div>
      </div>
      <div class="manual-actions">
        <button type="button" class="btn btn-primary" onclick="createExamBooklets()">توليد باركود امتحاني وإصدار الدفاتر</button>
      </div>
      <div id="exam-issue-status" class="exam-issue-status" role="status"></div>
    </div>
    <div id="exam-issue-preview" class="card no-print" style="display:none">
      <div class="card-title">الدفاتر الصادرة</div>
      <div id="exam-issue-summary"></div>
      <div class="manual-actions">
        <button class="btn btn-outline" onclick="openExamBookletPreview()">معاينة الدفاتر قبل الطباعة</button>
        <button class="btn btn-primary" onclick="printExamBooklets()">طباعة الدفاتر</button>
      </div>
    </div>
  </div>

  <div class="tab-content" id="tab-grade-exams">
    <div class="card no-print">
      <div class="card-title">تصحيح الدفاتر الامتحانية</div>
      <label class="lbl" for="correction-issue-select">إصدار الامتحان/المادة</label>
      <select id="correction-issue-select" onchange="loadCorrectionIssue(this.value)"><option value="">اختر إصداراً</option></select>
      <div id="correction-issue-status" class="exam-correction-status" role="status"></div>
      <div id="correction-key-transfer" style="display:none;margin-top:.5rem">
        <button class="btn btn-outline" onclick="approveCorrectionKey()">اعتماد مفتاح الإجابة لهذا الإصدار</button>
        <span id="correction-key-state"></span>
      </div>
      <div id="correction-key-import" style="display:none;margin-top:.8rem">
        <label class="lbl" for="correction-key-file">استيراد مفتاح JSON: modelA وmodelB، ولكل مجموعة قائمة إجابات بترتيب الأسئلة</label>
        <div class="manual-question-meta">مثال للشكل: {"modelA":[["أ","ب"],["صح"]],"modelB":[["ج","أ"],["خطأ"]]}. عدد المجموعات والإجابات يجب أن يطابق النماذج الصادرة.</div>
        <input type="file" id="correction-key-file" accept=".json,application/json">
        <button class="btn btn-outline" onclick="importCorrectionKey()">استيراد مفتاح الإجابة</button>
      </div>
    </div>
    <div class="card no-print" id="correction-scoring-card" style="display:none">
      <div class="card-title">درجات الأسئلة الرئيسية والتصحيح اليدوي</div>
      <div class="exam-scoring-grid">
        <section><h3>النموذج A</h3><div id="grading-model-A"></div></section>
        <section><h3>النموذج B</h3><div id="grading-model-B"></div></section>
      </div>
      <button class="btn btn-primary" style="margin-top:.8rem" onclick="saveCorrectionGrading()">حفظ الدرجات وإعداد التصحيح</button>
    </div>
    <div class="card no-print" id="correction-scan-card" style="display:none">
      <div class="card-title">المسح الضوئي لورقة الامتحان</div>
      <label class="lbl" for="exam-scan-file">استورد الدفاتر الممسوحة: ملف PDF أو صور بأي صيغة (PNG, JPG, JPEG, BMP, TIFF, WEBP) — يمكن اختيار عدة ملفات دفعة واحدة (أبقِ صفحات كل دفتر بترتيبها)</label>
      <input type="file" id="exam-scan-file" accept=".pdf,.png,.jpg,.jpeg,.bmp,.tif,.tiff,.webp,.gif,application/pdf,image/*" multiple>
      <button class="btn btn-primary" style="margin-top:.5rem" onclick="scanExamFiles()">مسح وتصحيح الإجابات المظللة</button>
      <button class="btn btn-outline" style="margin-top:.5rem" onclick="exportExamGrades()">إصدار الدرجات الامتحانية إلى Excel</button>
      <div id="exam-scan-status" class="exam-correction-status" role="status"></div>
      <div id="exam-scan-problems"></div>
      <div id="exam-review-list"></div>
    </div>
  </div>
</div>

<div id="exam-preview-modal" class="exam-preview-modal no-print" role="dialog" aria-modal="true" aria-labelledby="exam-preview-title">
  <div class="exam-preview-dialog">
    <div class="exam-preview-heading">
      <h2 id="exam-preview-title">معاينة الدفاتر الصادرة قبل الطباعة</h2>
      <button type="button" class="btn btn-outline" onclick="closeExamBookletPreview()" aria-label="إغلاق المعاينة">إغلاق</button>
    </div>
    <div class="exam-preview-controls">
      <button type="button" class="btn btn-outline" id="exam-preview-previous" onclick="moveExamBookletPreview(-1)">الدفتر السابق</button>
      <label class="lbl" for="exam-preview-booklet-select">الدفتر</label>
      <select id="exam-preview-booklet-select" onchange="renderExamBookletPreview(this.value)"></select>
      <button type="button" class="btn btn-outline" id="exam-preview-next" onclick="moveExamBookletPreview(1)">الدفتر التالي</button>
      <span id="exam-preview-position" class="manual-question-meta"></span>
      <button type="button" class="btn btn-primary" onclick="printExamBooklets()">طباعة جميع الدفاتر</button>
    </div>
    <div id="exam-preview-sheet" class="exam-preview-sheet"></div>
  </div>
</div>

<div id="toast"></div>
<div id="print-exam-preview" aria-hidden="true"></div>
<div id="booklet-print" aria-hidden="true"></div>

<script>
var DB = [];
var wordImportQuestions = [];
var selectedQuestionImage = null;
var selectedQuestionImageSize = 100;
var SUBJECTS = ['التعبية والتنظيم','التحصين','التخطيط وقراءة الخريطة','الإسعافات الأولية','الحاسبات الالكترونية','الاتصالات العسكرية','القوانين العسكرية','القيادة وعلم النفس العسكري','الأسلحة الكتلوية','التاريخ العسكري','الجغرافية العسكرية','اللغة الانكليزية'];
var ARABIC_ABJAD = ['أ', 'ب', 'جـ', 'د', 'هـ', 'و', 'ز', 'ح', 'ط', 'ي', 'ك', 'ل', 'م', 'ن', 'س', 'ع', 'ف', 'ص', 'ق', 'ر', 'ش', 'ت', 'ث', 'خ', 'ذ', 'ض', 'ظ', 'غ'];
var currentGeneratedExam = null;
var currentTypeCounts = {};
var generatedIssuedEditMode = false;
var generatedIssuedEditIssue = null;
var manualEditingIssue = null;
var manualEditingSelectedIds = null;
var printDialogPending = false;
var persistedLastIssuedExam = null;
var importedExamRoster = null;
var currentExamIssueId = null;
var currentExamBooklets = [];
var currentExamIssueSnapshot = null;
var currentCorrectionIssue = null;
var examScanResults = [];
var LAST_ISSUE_STORAGE_KEY = 'militaryExamLastIssued';
var currentQuestionCohort = 'A';
var COHORTS = {};

// توحيد مطابقة أنواع الأسئلة بدقة
function isSameType(type1, type2) {
  if (type1 === type2) return true;
  if ((type1 === 'multi_select' || type1 === 'multi') && (type2 === 'multi_select' || type2 === 'multi')) return true;
  return false;
}

function showToast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  setTimeout(function(){ t.classList.remove('show'); }, 3500);
}

function syncExamDuration(sourceId) {
  var value = document.getElementById(sourceId).value;
  document.getElementById('gen-duration').value = value;
  document.getElementById('manual-duration').value = value;
}

function getExamDurationLabel(value) {
  var labels = {'1': 'ساعة واحدة', '1.5': 'ساعة ونصف', '2': 'ساعتان', '3': 'ثلاث ساعات', '4': 'أربع ساعات'};
  return labels[String(value)] || 'ساعتان';
}

function getExamDateLabel() {
  return new Date().toLocaleDateString('ar-IQ');
}

function rememberLastIssuedExam(sourceMode) {
  if (!currentGeneratedExam || !Array.isArray(currentGeneratedExam.ids)) return;
  currentGeneratedExam.sourceMode = sourceMode;
  var issue = {
    subject: currentGeneratedExam.subject,
    year: currentGeneratedExam.year,
    semester: currentGeneratedExam.semester,
    pool: currentGeneratedExam.pool,
    order: currentGeneratedExam.order,
    duration: currentGeneratedExam.duration,
    durationValue: currentGeneratedExam.durationValue,
    examDate: currentGeneratedExam.examDate,
    sourceMode: sourceMode,
    cohort: currentQuestionCohort,
    ids: currentGeneratedExam.ids
  };
  persistedLastIssuedExam = issue;
  try {
    localStorage.setItem(LAST_ISSUE_STORAGE_KEY, JSON.stringify(issue));
  } catch (error) {}
  fetch('/api/issued/last', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(issue)
  }).catch(() => showToast('تعذر حفظ آخر إصدار في قاعدة البرنامج'));
}

function getLastIssuedExam() {
  if (currentGeneratedExam && Array.isArray(currentGeneratedExam.ids) && currentGeneratedExam.ids.length) {
    return currentGeneratedExam;
  }
  if (persistedLastIssuedExam && Array.isArray(persistedLastIssuedExam.ids)) return persistedLastIssuedExam;
  try {
    var saved = JSON.parse(localStorage.getItem(LAST_ISSUE_STORAGE_KEY) || 'null');
    return saved && Array.isArray(saved.ids) && saved.ids.length ? saved : null;
  } catch (error) {
    return null;
  }
}

function resolveLastIssuedExam() {
  var issue = getLastIssuedExam();
  if (issue) return Promise.resolve(issue);
  return fetch('/api/issued/last')
    .then(response => response.json())
    .then(data => {
      persistedLastIssuedExam = data.issue;
      return data.issue;
    });
}

function loadLastIssuedExam() {
  fetch('/api/issued/last')
    .then(response => response.json())
    .then(data => {
      persistedLastIssuedExam = data.issue;
      if (data.issue) localStorage.setItem(LAST_ISSUE_STORAGE_KEY, JSON.stringify(data.issue));
    })
    .catch(() => {});
}

function restoreIssuedQuestionStatuses(issue) {
  var ids = Array.from(new Set((issue.ids || []).filter(id => Number.isInteger(id) && id > 0)));
  if (!ids.length) return Promise.reject(new Error('لا توجد أسئلة محفوظة لآخر إصدار'));

  return fetch('/api/questions/mark-available', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ids: ids, cohort: issue.cohort || currentQuestionCohort})
  }).then(response => response.json().then(data => ({response: response, data: data})))
    .then(result => {
      if (!result.response.ok || result.data.status !== 'success') {
        throw new Error(result.data.message || 'تعذر استرجاع الأسئلة');
      }
      return loadDB(issue.cohort || currentQuestionCohort).then(() => result.data);
    });
}

function showTab(name) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  if(name === 'generate') updateAvailableCounts();
  if(name === 'manual') loadManualTopics();
  if(name === 'entry') loadAvailableQuestions();
  if(name === 'data-dashboard') {
    loadJwtDatabase();
    loadQuestionMemory();
  }
}

function loadDB(cohort) {
  var selectedCohort = cohort || currentQuestionCohort;
  return fetch('/api/questions?cohort=' + encodeURIComponent(selectedCohort))
    .then(res => res.json())
    .then(data => {
      DB = data;
      updateKPIs();
      renderStats();
      updateAvailableCounts();
      loadAvailableQuestions();
      loadManualTopics();
    });
}

function changeQuestionCohort(cohort) {
  currentQuestionCohort = cohort;
  ['gen-cohort', 'manual-cohort'].forEach(function(id) {
    var select = document.getElementById(id);
    if (select) select.value = cohort;
  });
  loadDB(cohort);
}

function initializeCohorts() {
  return fetch('/api/cohorts').then(response => response.json()).then(items => {
    COHORTS = {};
    ['gen-cohort', 'manual-cohort', 'memory-reset-cohort'].forEach(function(id) {
      var select = document.getElementById(id);
      if (!select) return;
      select.innerHTML = '';
      items.forEach(function(item) {
        COHORTS[item.code] = item.name;
        var option = document.createElement('option');
        option.value = item.code;
        option.textContent = item.name;
        select.appendChild(option);
      });
      select.value = currentQuestionCohort;
    });
  });
}

function loadJwtDatabase() {
  var input = document.getElementById('jwt-search');
  var query = encodeURIComponent(input ? input.value : '');
  fetch('/api/jwt-database?q=' + query + '&page_size=100')
    .then(response => response.json())
    .then(data => {
      if (data.status === 'error') throw new Error(data.message);
      document.getElementById('jwt-summary').textContent = 'عدد سجلات JWT: ' + data.total;
      document.getElementById('jwt-table').innerHTML = (data.rows || []).length
        ? '<table class="data-table"><thead><tr><th>تسلسل الطالب</th><th>اسم الطالب</th><th>السرية</th><th>الفصيل</th><th>JWT</th></tr></thead><tbody>' +
          data.rows.map(row => '<tr><td>' + escapeHtml(String(row.student_sequence)) + '</td><td>' +
          escapeHtml(row.student_name) + '</td><td>' + escapeHtml(String(row.company)) +
          '</td><td>' + escapeHtml(row.platoon) + '</td><td style="direction:ltr;text-align:left;min-width:420px">' +
          escapeHtml(row.jwt) + '</td></tr>').join('') + '</tbody></table>'
        : '<div class="manual-question-meta">لا توجد سجلات مطابقة.</div>';
    })
    .catch(error => {
      document.getElementById('jwt-summary').textContent = 'تعذر تحميل قاعدة JWT: ' + error.message;
    });
}

function loadQuestionMemory() {
  fetch('/api/question-memory').then(response => response.json()).then(data => {
    if (data.status === 'error') throw new Error(data.message);
    document.getElementById('cohort-memory-summary').innerHTML =
      Object.keys(data.cohorts).map(code => '<div class="kpi"><div class="kpi-val">' +
      (data.used_counts[code] || 0) + '</div><div class="kpi-lbl">' +
      escapeHtml(data.cohorts[code]) + '</div></div>').join('');
    document.getElementById('cohort-history').innerHTML =
      '<table class="data-table"><thead><tr><th>التاريخ</th><th>الفوج</th><th>المادة</th><th>المصدر</th><th>عدد الأسئلة</th></tr></thead><tbody>' +
      (data.history || []).map(item => '<tr><td>' + escapeHtml(item.created_at) +
      '</td><td>' + escapeHtml(item.cohort_name) + '</td><td>' + escapeHtml(item.subject) +
      '</td><td>' + escapeHtml(item.source_mode) + '</td><td>' + item.question_count +
      '</td></tr>').join('') + '</tbody></table>';
  }).catch(error => showToast('تعذر تحميل ذاكرة الأفواج: ' + error.message));
}

function resetCohortMemory() {
  var cohort = document.getElementById('memory-reset-cohort').value;
  if (!confirm('سيتم تصفير ذاكرة أسئلة الفوج المحدد وحذف سجل إصداراته. هل تريد المتابعة؟')) return;
  fetch('/api/questions/reset-used', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({cohort: cohort})
  }).then(response => response.json().then(data => ({response: response, data: data})))
    .then(result => {
      if (!result.response.ok || result.data.status === 'error') throw new Error(result.data.message);
      showToast('تم تصفير ذاكرة الفوج وحفظ نسخة احتياطية ✓');
      return loadDB(currentQuestionCohort);
    }).then(loadQuestionMemory)
    .catch(error => showToast('تعذر تصفير ذاكرة الفوج: ' + error.message));
}

function updateKPIs() {
  var total = DB.length;
  var used = DB.filter(q => q.used).length;
  var unused = total - used;
  document.getElementById('kpi-total').textContent = total;
  document.getElementById('hdr-total').textContent = total;
  document.getElementById('kpi-used').textContent = used;
  document.getElementById('hdr-used').textContent = used;
  document.getElementById('kpi-unused').textContent = unused;
  document.getElementById('hdr-unused').textContent = unused;
}

function populateDropdowns() {
  ['ent-subject', 'gen-subject', 'manual-subject', 'issue-subject'].forEach(id => {
    var el = document.getElementById(id);
    el.innerHTML = '';
    SUBJECTS.forEach(s => {
      var opt = document.createElement('option');
      opt.value = s; opt.textContent = s;
      el.appendChild(opt);
    });
  });
  ['gen-cohort', 'manual-cohort', 'memory-reset-cohort'].forEach(id => {
    var el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = '';
    Object.keys(COHORTS).forEach(code => {
      var opt = document.createElement('option');
      opt.value = code; opt.textContent = COHORTS[code];
      el.appendChild(opt);
    });
    el.value = currentQuestionCohort;
  });
}

function updateAvailableCounts() {
  var subj = document.getElementById('gen-subject').value;
  var yr = document.getElementById('gen-year').value;
  var sem = document.getElementById('gen-sem').value;
  var pool = document.getElementById('gen-pool').value;

  var filtered = DB.filter(q => {
    if (q.subject !== subj) return false;
    if (pool === 'unused' && q.used) return false;
    if (pool === 'used' && !q.used) return false;
    if (yr === 'all1') return q.year == 1;
    if (yr === 'all2') return q.year == 2;
    if (yr === 'all3') return q.year == 3;
    if (yr !== 'all') {
      if (q.year != yr) return false;
      if (q.semester != sem) return false;
    }
    return true;
  });

  var types = ['mc', 'tf', 'fill', 'essay', 'multi_select'];
  currentTypeCounts = {};

  types.forEach(t => {
    var totalForType = DB.filter(q => {
      if (q.subject !== subj || !isSameType(q.type, t)) return false;
      if (yr === 'all1') return q.year == 1;
      if (yr === 'all2') return q.year == 2;
      if (yr === 'all3') return q.year == 3;
      if (yr !== 'all') {
        if (q.year != yr) return false;
        if (q.semester != sem) return false;
      }
      return true;
    }).length;

    var available = filtered.filter(q => isSameType(q.type, t)).length;
    currentTypeCounts[t] = available;

    var tag = document.getElementById('stat-' + t);
    if (tag) {
      tag.textContent = `[ متاح: ${available} | كلي: ${totalForType} ]`;
    }
  });
}

function issueMatchesGeneratedScope(question, issue) {
  var year = String(issue.year || 'all');
  if (year === 'all1' && question.year != 1) return false;
  if (year === 'all2' && question.year != 2) return false;
  if (year === 'all3' && question.year != 3) return false;
  if (['1', '2', '3'].includes(year) && question.year != year) return false;
  if (['1', '2', '3'].includes(year) && issue.semester && question.semester != issue.semester) return false;
  return true;
}

function renderGeneratedIssuedEditor(issue) {
  var originalIds = new Set(issue.ids.map(Number));
  var candidates = DB.filter(question => question.subject === issue.subject &&
    issueMatchesGeneratedScope(question, issue) && (!question.used || originalIds.has(question.id)));
  var groups = [
    {id: 'mc', label: 'اختيار من متعدد'},
    {id: 'tf', label: 'صح وخطأ'},
    {id: 'fill', label: 'ملء الفراغات'},
    {id: 'essay', label: 'سؤال شرح ومقالي'},
    {id: 'multi_select', label: 'اختيار الإجابات الصحيحة'}
  ];
  var list = document.getElementById('generated-issued-list');
  list.innerHTML = '';

  groups.forEach(group => {
    var groupQuestions = candidates.filter(question => isSameType(question.type, group.id));
    if (!groupQuestions.length) return;
    var section = document.createElement('section');
    section.className = 'manual-type-group';
    var heading = document.createElement('h3');
    heading.className = 'manual-type-title';
    heading.textContent = group.label;
    var rows = document.createElement('div');
    rows.className = 'manual-question-list';

    groupQuestions.forEach(question => {
      var label = document.createElement('label');
      label.className = 'manual-question-row';
      var checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'generated-issued-check';
      checkbox.value = question.id;
      checkbox.checked = originalIds.has(question.id);
      checkbox.onchange = updateGeneratedIssueCounts;
      var content = document.createElement('span');
      content.className = 'manual-question-text';
      var text = document.createElement('strong');
      text.textContent = question.text;
      var meta = document.createElement('span');
      meta.className = 'manual-question-meta';
      meta.textContent = 'المفردة: ' + question.topic + ' | السنة: ' + question.year + ' | الفصل: ' + question.semester;
      content.appendChild(text);
      content.appendChild(meta);
      label.appendChild(checkbox);
      label.appendChild(content);
      rows.appendChild(label);
    });

    section.appendChild(heading);
    section.appendChild(rows);
    list.appendChild(section);
  });
  updateGeneratedIssueCounts();
  document.getElementById('generated-issued-editor').style.display = 'block';
}

function updateGeneratedIssueCounts() {
  var selectedIds = new Set(Array.from(document.querySelectorAll('.generated-issued-check:checked')).map(input => Number(input.value)));
  var fields = [['mc', 'num-mc'], ['tf', 'num-tf'], ['fill', 'num-fill'], ['essay', 'num-essay'], ['multi_select', 'num-multi']];
  fields.forEach(([type, inputId]) => {
    var count = DB.filter(question => selectedIds.has(question.id) && isSameType(question.type, type)).length;
    document.getElementById(inputId).value = count;
  });
}

function selectGeneratedIssuedQuestions(checked) {
  document.querySelectorAll('.generated-issued-check').forEach(input => input.checked = checked);
  updateGeneratedIssueCounts();
}

function restoreLastIssueForGenerated() {
  resolveLastIssuedExam()
    .then(issue => {
      if (!issue) {
        showToast('لا يوجد إصدار سابق لاسترجاعه');
        return;
      }
      generatedIssuedEditIssue = issue;
      generatedIssuedEditMode = true;
      currentQuestionCohort = issue.cohort || currentQuestionCohort;
      document.getElementById('gen-cohort').value = currentQuestionCohort;
      document.getElementById('manual-cohort').value = currentQuestionCohort;
      var savedYear = ['1', '2', '3', 'all', 'all1', 'all2', 'all3'].includes(String(issue.year)) ? String(issue.year) : 'all';
      document.getElementById('gen-subject').value = issue.subject;
      document.getElementById('gen-year').value = savedYear;
      document.getElementById('gen-sem').value = ['1', '2'].includes(String(issue.semester)) ? String(issue.semester) : '1';
      document.getElementById('gen-pool').value = 'unused';
      document.getElementById('gen-order').value = issue.order || 'random';
      document.getElementById('gen-duration').value = String(issue.durationValue || '2');
      document.getElementById('manual-duration').value = String(issue.durationValue || '2');
      ['chk-mc', 'chk-tf', 'chk-fill', 'chk-essay', 'chk-multi'].forEach(id => {
        document.getElementById(id).checked = true;
      });
      return restoreIssuedQuestionStatuses(issue).then(() => {
        showTab('generate');
        renderGeneratedIssuedEditor(issue);
        showToast('تم استرجاع أسئلة النموذج أ؛ أضف أو ألغِ ما تريد ثم أصدر الأسئلة');
      });
    })
    .catch(error => {
      generatedIssuedEditMode = false;
      generatedIssuedEditIssue = null;
      showToast('تعذر استرجاع الإصدار: ' + error.message);
    });
}

function toggleEntryFields() {
  var t = document.getElementById('ent-type').value;
  document.querySelectorAll('.type-fields').forEach(el => el.style.display = 'none');
  if (t === 'mc') document.getElementById('field-mc').style.display = 'block';
  if (t === 'tf') document.getElementById('field-tf').style.display = 'block';
  if (t === 'fill') document.getElementById('field-fill').style.display = 'block';
  if (t === 'multi_select') {
    document.getElementById('field-multi').style.display = 'block';
    if(document.getElementById('multi-options-container').children.length === 0) {
      addMultiOptionField(); addMultiOptionField();
    }
  }
}

function addMultiOptionField(value = '', isChecked = false) {
  var container = document.getElementById('multi-options-container');
  var idx = container.children.length;
  var symbol = ARABIC_ABJAD[idx % ARABIC_ABJAD.length];
  var div = document.createElement('div');
  div.style.display = 'flex'; div.style.gap = '.4rem'; div.style.alignItems = 'center'; div.style.marginBottom = '.3rem';
  div.innerHTML = `<input type="checkbox" class="multi-check" ${isChecked?'checked':''}> <span>${symbol}-</span> <input type="text" class="multi-text" value="${escapeHtml(value)}" placeholder="الخيار الفرعي"> <button type="button" class="btn btn-danger" onclick="this.parentElement.remove(); reIndexMulti();">✕</button>`;
  container.appendChild(div);
}

function reIndexMulti() {
  document.querySelectorAll('#multi-options-container > div').forEach((row, i) => {
    row.querySelector('span').textContent = ARABIC_ABJAD[i % ARABIC_ABJAD.length] + '-';
  });
}

function renderEntryImagePreview() {
  var wrap = document.getElementById('ent-image-preview-wrap');
  var sizeControl = document.getElementById('ent-image-size-control');
  var preview = document.getElementById('ent-image-preview');
  preview.src = selectedQuestionImage || '';
  preview.style.width = selectedQuestionImageSize + '%';
  wrap.style.display = selectedQuestionImage ? 'block' : 'none';
  sizeControl.style.display = selectedQuestionImage ? 'grid' : 'none';
  document.getElementById('ent-image-size').value = selectedQuestionImageSize;
  document.getElementById('ent-image-size-value').textContent = selectedQuestionImageSize + '%';
}

function getQuestionImageSize(value) {
  var size = Number(value);
  return Number.isFinite(size) ? Math.min(100, Math.max(20, Math.round(size))) : 100;
}

function updateQuestionImageSize(value) {
  selectedQuestionImageSize = getQuestionImageSize(value);
  document.getElementById('ent-image-size').value = selectedQuestionImageSize;
  document.getElementById('ent-image-size-value').textContent = selectedQuestionImageSize + '%';
  document.getElementById('ent-image-preview').style.width = selectedQuestionImageSize + '%';
}

function loadQuestionImage(input) {
  var file = input.files[0];
  if (!file) return;
  if (!['image/png', 'image/jpeg'].includes(file.type)) {
    showToast('اختر صورة بصيغة PNG أو JPEG');
    input.value = '';
    return;
  }
  if (file.size > 5 * 1024 * 1024) {
    showToast('الحد الأقصى لحجم الصورة 5 ميغابايت');
    input.value = '';
    return;
  }
  var reader = new FileReader();
  reader.onload = function() {
    selectedQuestionImage = reader.result;
    renderEntryImagePreview();
  };
  reader.onerror = function() { showToast('تعذرت قراءة الصورة'); };
  reader.readAsDataURL(file);
}

function removeQuestionImage() {
  selectedQuestionImage = null;
  selectedQuestionImageSize = 100;
  document.getElementById('ent-image').value = '';
  renderEntryImagePreview();
}

function saveQuestion() {
  var type = document.getElementById('ent-type').value;
  var editId = document.getElementById('ent-edit-id').value;
  var qData = {
    subject: document.getElementById('ent-subject').value,
    topic: document.getElementById('ent-topic').value,
    year: document.getElementById('ent-year').value,
    semester: document.getElementById('ent-sem').value,
    type: type,
    text: document.getElementById('ent-text').value,
    answer: document.getElementById('ent-answer').value,
    image: selectedQuestionImage,
    image_size: selectedQuestionImageSize
  };

  if (type === 'mc') {
    qData.options = [
      document.getElementById('mc-a').value,
      document.getElementById('mc-b').value,
      document.getElementById('mc-c').value,
      document.getElementById('mc-d').value
    ];
  } else if (type === 'multi_select') {
    var opts = [], correctSymbols = [];
    document.querySelectorAll('#multi-options-container > div').forEach((row, i) => {
      var txt = row.querySelector('.multi-text').value;
      var chk = row.querySelector('.multi-check').checked;
      if (txt) { 
        opts.push(txt); 
        if(chk) correctSymbols.push(ARABIC_ABJAD[i % ARABIC_ABJAD.length]); 
      }
    });
    qData.options = opts;
    if(!qData.answer) qData.answer = correctSymbols.join('، ');
  } else if (type === 'tf') {
    if(!qData.answer) qData.answer = document.getElementById('tf-ans').value;
  } else if (type === 'fill') {
    qData.options = ['fill-a', 'fill-b', 'fill-c', 'fill-d']
      .map(id => document.getElementById(id).value.trim())
      .filter(Boolean);
    if (qData.options.length === 1) {
      showToast('أدخل خيارين على الأقل لسؤال الفراغ أو اترك الخيارات فارغة للتصحيح اليدوي.');
      return;
    }
    if (qData.options.length && !answerIndexes({answer:qData.answer}, qData.options).length) {
      showToast('حدد إجابة سؤال الفراغ الصحيحة برمز الخيار أو نصه.');
      return;
    }
  }

  var url = editId ? '/api/questions/' + editId : '/api/questions';
  var method = editId ? 'PUT' : 'POST';

  fetch(url, {
    method: method,
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(qData)
  }).then(res => res.json().then(data => ({response: res, data: data}))).then(result => {
    if (!result.response.ok) throw new Error(result.data.message || 'تعذر حفظ السؤال');
    showToast(editId ? 'تم تعديل السؤال بنجاح ✓' : 'تم حفظ السؤال وتفريغ الحقول بنجاح ✓');
    clearEntryForm();
    loadDB();
    if (editId) cancelEdit();
  }).catch(error => showToast('فشل حفظ السؤال: ' + error.message));
}

function createWordImportField(label, field, value, choices, isWide) {
  var wrapper = document.createElement('label');
  wrapper.className = 'word-import-field' + (isWide ? ' wide' : '');
  var caption = document.createElement('span');
  caption.textContent = label;
  wrapper.appendChild(caption);

  var control;
  if (choices) {
    control = document.createElement('select');
    choices.forEach(choice => {
      var option = document.createElement('option');
      option.value = choice.value;
      option.textContent = choice.label;
      control.appendChild(option);
    });
    control.value = value == null ? '' : String(value);
  } else if (isWide) {
    control = document.createElement('textarea');
    control.value = Array.isArray(value) ? value.join('\\n') : (value || '');
  } else {
    control = document.createElement('input');
    control.type = 'text';
    control.value = value == null ? '' : String(value);
  }
  control.dataset.importField = field;
  wrapper.appendChild(control);
  return wrapper;
}

function renderWordImportPreview(questions) {
  wordImportQuestions = questions;
  var container = document.getElementById('word-import-rows');
  container.innerHTML = '';
  container.style.display = 'none';
  document.getElementById('word-import-details-toggle').textContent = 'مراجعة وتعديل التفاصيل';
  var types = [
    {value:'', label:'اختر نوع السؤال'},
    {value:'mc', label:'اختيار من متعدد'},
    {value:'tf', label:'صح وخطأ'},
    {value:'fill', label:'ملء الفراغات'},
    {value:'essay', label:'شرح'},
    {value:'multi_select', label:'اختيار الإجابات الصحيحة'}
  ];
  var years = [{value:'', label:'اختر السنة'}, {value:'1', label:'الأولى'}, {value:'2', label:'الثانية'}, {value:'3', label:'الثالثة'}];
  var semesters = [{value:'', label:'اختر النصف'}, {value:'1', label:'الأول'}, {value:'2', label:'الثاني'}];

  questions.forEach((question, index) => {
    var row = document.createElement('section');
    row.className = 'word-import-row';
    row.dataset.questionIndex = index;
    row.dataset.duplicate = question.duplicate ? 'true' : 'false';
    var heading = document.createElement('div');
    heading.className = 'word-import-row-head';
    var checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = !question.duplicate && isWordImportQuestionReady(question);
    checkbox.addEventListener('change', updateWordImportSummary);
    checkbox.dataset.importSelected = 'true';
    var title = document.createElement('span');
    title.textContent = 'السؤال ' + (index + 1);
    heading.append(checkbox, title);
    if (question.duplicate) {
      var duplicate = document.createElement('span');
      duplicate.className = 'type-stat-tag';
      duplicate.textContent = 'مكرر؛ لن يُضاف مرة أخرى';
      heading.appendChild(duplicate);
    }
    row.appendChild(heading);

    var fields = document.createElement('div');
    fields.className = 'word-import-fields';
    fields.appendChild(createWordImportField('المادة', 'subject', question.subject));
    fields.appendChild(createWordImportField('المفردة', 'topic', question.topic));
    fields.appendChild(createWordImportField('السنة الدراسية', 'year', question.year, years));
    fields.appendChild(createWordImportField('النصف الدراسي', 'semester', question.semester, semesters));
    fields.appendChild(createWordImportField('نوع السؤال', 'type', question.type, types));
    fields.appendChild(createWordImportField('نص السؤال', 'text', question.text, null, true));
    fields.appendChild(createWordImportField('الخيارات (كل خيار في سطر)', 'options', question.options, null, true));
    fields.appendChild(createWordImportField('الإجابة', 'answer', question.answer, null, true));
    row.appendChild(fields);

    var warning = document.createElement('div');
    warning.className = 'word-import-warning';
    warning.textContent = question.duplicate
      ? 'يوجد سؤال مطابق في قاعدة البيانات؛ لن يُضاف مرة أخرى.'
      : (isWordImportQuestionReady(question) ? 'جاهز للترحيل.' : 'تحتاج بعض الحقول إلى استكمال.');
    row.appendChild(warning);
    fields.addEventListener('input', function() {
      if (row.dataset.duplicate === 'true') {
        row.dataset.duplicate = 'false';
        warning.textContent = 'تم تعديل بيانات السؤال؛ راجعها قبل الترحيل.';
      }
      updateWordImportSummary();
    });
    fields.addEventListener('change', updateWordImportSummary);
    container.appendChild(row);
  });
  document.getElementById('word-import-preview').style.display = 'block';
  updateWordImportSummary();
}

function isWordImportQuestionReady(question) {
  return !!(question.subject && question.topic && question.text && question.type &&
    ['1', '2', '3'].includes(String(question.year)) && ['1', '2'].includes(String(question.semester)));
}

function updateWordImportSummary() {
  var rows = Array.from(document.querySelectorAll('#word-import-rows .word-import-row'));
  var duplicateCount = 0;
  var incompleteCount = 0;
  var readyCount = 0;
  var selectedReadyCount = 0;
  rows.forEach(row => {
    var question = getWordImportQuestion(row);
    var duplicate = row.dataset.duplicate === 'true';
    var ready = isWordImportQuestionReady(question);
    if (duplicate) duplicateCount++;
    else if (ready) readyCount++;
    else incompleteCount++;
    if (ready && !duplicate && row.querySelector('[data-import-selected]').checked) selectedReadyCount++;
  });
  document.getElementById('word-import-summary').textContent =
    'الأسئلة: ' + rows.length + ' | جاهز: ' + readyCount + ' | مكرر: ' + duplicateCount + ' | يحتاج استكمال: ' + incompleteCount;
  var submit = document.getElementById('word-import-submit');
  submit.disabled = selectedReadyCount === 0;
  submit.textContent = selectedReadyCount ? 'ترحيل ' + selectedReadyCount + ' سؤال جاهز' : 'لا توجد أسئلة جاهزة محددة';
}

function toggleWordImportDetails() {
  var rows = document.getElementById('word-import-rows');
  var showDetails = rows.style.display === 'none';
  rows.style.display = showDetails ? 'block' : 'none';
  document.getElementById('word-import-details-toggle').textContent = showDetails ? 'إخفاء التفاصيل' : 'مراجعة وتعديل التفاصيل';
}

function previewWordImport() {
  var input = document.getElementById('word-import-file');
  var file = input.files && input.files[0];
  var status = document.getElementById('word-import-status');
  if (!file) {
    status.textContent = 'اختر مستند DOCX أولاً.';
    return;
  }
  if (!file.name.toLowerCase().endsWith('.docx')) {
    status.textContent = 'اختر مستند Word بصيغة DOCX.';
    return;
  }
  if (file.size > 8 * 1024 * 1024) {
    status.textContent = 'الحد الأقصى لحجم المستند 8 ميغابايت.';
    return;
  }

  var formData = new FormData();
  formData.append('file', file);
  status.textContent = 'جارٍ قراءة المستند...';
  document.getElementById('word-import-preview').style.display = 'none';
  fetch('/api/import-word/preview', {method:'POST', body:formData})
    .then(response => response.json().then(data => ({response:response, data:data})))
    .then(result => {
      if (!result.response.ok || result.data.status !== 'success') {
        throw new Error(result.data.message || 'تعذرت قراءة الأسئلة');
      }
      status.textContent = 'اكتملت قراءة المستند. راجع بيانات كل سؤال أدناه.';
      renderWordImportPreview(result.data.questions);
    })
    .catch(error => { status.textContent = 'فشلت قراءة المستند: ' + error.message; });
}

function getWordImportQuestion(row) {
  var value = field => row.querySelector('[data-import-field="' + field + '"]').value.trim();
  return {
    subject:value('subject'),
    topic:value('topic'),
    year:value('year'),
    semester:value('semester'),
    type:value('type'),
    text:value('text'),
    options:value('options').split(String.fromCharCode(10)).map(option => option.trim()).filter(Boolean),
    answer:value('answer')
  };
}

function importSelectedWordQuestions() {
  var rows = Array.from(document.querySelectorAll('#word-import-rows .word-import-row'));
  var selectedRows = rows.filter(row => row.querySelector('[data-import-selected]').checked);
  if (!selectedRows.length) {
    showToast('حدد سؤالًا واحدًا على الأقل للترحيل');
    return;
  }
  var questions = selectedRows.map(getWordImportQuestion);
  var missingIndex = questions.findIndex(question => !question.subject || !question.topic || !question.text || !question.type || !question.year || !question.semester);
  if (missingIndex >= 0) {
    showToast('أكمل المادة والمفردة والسنة والنصف والنوع للسؤال المحدد رقم ' + (missingIndex + 1));
    return;
  }
  fetch('/api/import-word', {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({questions:questions})
  }).then(response => response.json().then(data => ({response:response, data:data})))
    .then(result => {
      if (!result.response.ok || result.data.status !== 'success') {
        throw new Error(result.data.message || 'تعذر ترحيل الأسئلة');
      }
      showToast('تم ترحيل ' + result.data.imported + ' سؤالًا، وتجاوز ' + result.data.skippedDuplicates + ' سؤال مكرر');
      document.getElementById('word-import-status').textContent = 'اكتمل الترحيل إلى قاعدة بيانات البرنامج. جارٍ تحديث حالة التكرار...';
      return loadDB().then(() => previewWordImport());
    })
    .catch(error => showToast('فشل ترحيل الأسئلة: ' + error.message));
}

function resetWordImportPreview() {
  wordImportQuestions = [];
  document.getElementById('word-import-status').textContent = '';
  document.getElementById('word-import-preview').style.display = 'none';
  document.getElementById('word-import-summary').textContent = '';
  document.getElementById('word-import-rows').innerHTML = '';
}

function clearWordImportPreview() {
  document.getElementById('word-import-file').value = '';
  resetWordImportPreview();
}

function cancelEdit() {
  document.getElementById('ent-edit-id').value = '';
  document.getElementById('entry-card-title').textContent = 'إدخال سؤال جديد';
  document.getElementById('btn-cancel-edit').style.display = 'none';
  clearEntryForm();
}

function clearEntryForm() {
  document.getElementById('ent-text').value = '';
  document.getElementById('ent-answer').value = '';
  document.getElementById('mc-a').value = '';
  document.getElementById('mc-b').value = '';
  document.getElementById('mc-c').value = '';
  document.getElementById('mc-d').value = '';
  ['fill-a', 'fill-b', 'fill-c', 'fill-d'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('multi-options-container').innerHTML = '';
  selectedQuestionImage = null;
  selectedQuestionImageSize = 100;
  document.getElementById('ent-image').value = '';
  renderEntryImagePreview();
}

function editQuestion(id) {
  var q = DB.find(item => item.id === id);
  if (!q) return;

  document.getElementById('ent-edit-id').value = q.id;
  document.getElementById('ent-subject').value = q.subject;
  document.getElementById('ent-topic').value = q.topic;
  document.getElementById('ent-year').value = q.year;
  document.getElementById('ent-sem').value = q.semester;
  
  var normalizedType = (q.type === 'multi') ? 'multi_select' : q.type;
  document.getElementById('ent-type').value = normalizedType;
  document.getElementById('ent-text').value = q.text;
  document.getElementById('ent-answer').value = q.answer || '';
  selectedQuestionImage = q.image || null;
  selectedQuestionImageSize = getQuestionImageSize(q.image_size);
  renderEntryImagePreview();

  toggleEntryFields();

  if (normalizedType === 'mc' && Array.isArray(q.options)) {
    document.getElementById('mc-a').value = q.options[0] || '';
    document.getElementById('mc-b').value = q.options[1] || '';
    document.getElementById('mc-c').value = q.options[2] || '';
    document.getElementById('mc-d').value = q.options[3] || '';
  } else if (normalizedType === 'fill' && Array.isArray(q.options)) {
    ['fill-a', 'fill-b', 'fill-c', 'fill-d'].forEach((field, index) => {
      document.getElementById(field).value = q.options[index] || '';
    });
  } else if (normalizedType === 'multi_select' && Array.isArray(q.options)) {
    document.getElementById('multi-options-container').innerHTML = '';
    q.options.forEach(opt => addMultiOptionField(opt, false));
  }

  document.getElementById('entry-card-title').textContent = 'تعديل السؤال الحالي';
  document.getElementById('btn-cancel-edit').style.display = 'inline-flex';
  showTab('entry');
}

function loadAvailableQuestions() {
  var subj = document.getElementById('ent-subject').value;
  var list = document.getElementById('available-questions-list');
  if(!list) return;
  list.innerHTML = '';
  var filtered = DB.filter(q => q.subject === subj && !q.used);
  if(filtered.length === 0) {
    list.innerHTML = '<div style="color:var(--text-muted);font-size:.85rem">لا توجد أسئلة متاحة حالياً لهذه المادة.</div>';
    return;
  }
  filtered.forEach(q => {
    var displayType = isSameType(q.type, 'multi_select') ? 'اختيار الإجابات' : q.type;
    list.innerHTML += `<div class="q-item">
      <div>
        <div><strong>[${escapeHtml(displayType)}]</strong> ${escapeHtml(q.text)}</div>
        <div class="q-ans">الإجابة الصحيحة: ${escapeHtml(q.answer || 'غير محددة')}</div>
      </div>
      <div style="display:flex;gap:.3rem">
        <button class="btn btn-edit" onclick="editQuestion(${q.id})">تعديل</button>
        <button class="btn btn-danger" onclick="deleteQ(${q.id})">حذف</button>
      </div>
    </div>`;
  });
}

function deleteQ(id) {
  if(confirm('هل أنت تأكد من حذف هذا السؤال؟')) {
    fetch('/api/questions/' + id, {method: 'DELETE'})
      .then(() => { showToast('تم الحذف بنجاح ✓'); loadDB(); });
  }
}

function resetAllUsed() {
  var cohort = currentQuestionCohort || document.getElementById('gen-cohort').value || 'A';
  var label = COHORTS[cohort] || cohort;
  if (confirm('هل تريد تصفير ذاكرة الأسئلة للفوج: ' + label + '؟')) {
    fetch('/api/questions/reset-used', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({cohort: cohort})
    })
      .then(response => response.json().then(data => ({response: response, data: data})))
      .then(result => {
        if (!result.response.ok) throw new Error(result.data.message || 'تعذر تصفير الذاكرة');
        showToast('تم تصفير ذاكرة الفوج وإعادة الأسئلة المتاحة ✓');
        loadDB();
      })
      .catch(error => showToast('فشل التصفير: ' + error.message));
  }
}

function backupQuestions() {
  fetch('/api/backup', {method: 'POST'})
    .then(response => response.json().then(data => ({response: response, data: data})))
    .then(result => {
      if (!result.response.ok || result.data.status !== 'success') {
        throw new Error(result.data.message || 'تعذر إنشاء النسخة الاحتياطية');
      }
      showToast('تم إنشاء النسخة الاحتياطية بنجاح');
    })
    .catch(error => showToast('فشل النسخ الاحتياطي: ' + error.message));
}

function restoreQuestionsBackup() {
  if (!confirm('سيتم استبدال الأسئلة الحالية ببيانات أحدث نسخة احتياطية. سيحفظ البرنامج نسخة أمان من الحالة الحالية. هل تريد المتابعة؟')) return;

  fetch('/api/backup/restore', {method: 'POST'})
    .then(response => response.json().then(data => ({response: response, data: data})))
    .then(result => {
      if (!result.response.ok || result.data.status !== 'success') {
        throw new Error(result.data.message || 'تعذر استيراد النسخة الاحتياطية');
      }
      loadDB();
      showToast('تم استيراد الأسئلة بنجاح: ' + result.data.questions);
    })
    .catch(error => showToast('فشل الاستيراد: ' + error.message));
}

function performSearch() {
  var q = document.getElementById('search-input').value;
  fetch('/api/search?q=' + encodeURIComponent(q))
    .then(res => res.json())
    .then(data => {
      var wrap = document.getElementById('search-results');
      wrap.innerHTML = '';
      data.forEach(item => {
        wrap.innerHTML += `<div class="q-item">
          <div>
            <div><b>${escapeHtml(item.subject)} (${escapeHtml(item.year)}):</b> ${escapeHtml(item.text)}</div>
            <div class="q-ans">الإجابة الصحيحة: ${escapeHtml(item.answer || 'غير محددة')}</div>
          </div>
          <div style="display:flex;gap:.3rem">
            <button class="btn btn-edit" onclick="editQuestion(${item.id})">تعديل</button>
            <button class="btn btn-danger" onclick="deleteQ(${item.id})">حذف</button>
          </div>
        </div>`;
      });
    });
}

function shuffleArray(array) {
  var arr = array.slice();
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = arr[i]; arr[i] = arr[j]; arr[j] = temp;
  }
  if (arr.length > 1 && arr.every((item, index) => item === array[index])) {
    arr.push(arr.shift());
  }
  return arr;
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function answerIndexes(item, options) {
  var answer = String(item.answer || '').trim();
  if (!answer || !Array.isArray(options)) return [];
  var tokens = answer.split(/[،,\x2f\\n]+/).map(token => token.trim()).filter(Boolean);
  return options.reduce((indexes, option, index) => {
    var symbol = ARABIC_ABJAD[index % ARABIC_ABJAD.length];
    if (tokens.some(token => token === symbol || token.replace(/ـ/g, '') === symbol.replace(/ـ/g, '') ||
        token === String(option).trim())) {
      indexes.push(index);
    }
    return indexes;
  }, []);
}

function shuffleQuestionOptions(question) {
  if (!Array.isArray(question.options)) return question;
  var originalOptions = question.options.slice();
  var correctIndexes = answerIndexes(question, originalOptions);
  question.options = shuffleArray(originalOptions);
  if (correctIndexes.length) {
    var correctOptions = correctIndexes.map(index => originalOptions[index]);
    var newIndexes = answerIndexes({answer: correctOptions.join('، ')}, question.options);
    question.answer = newIndexes.map(index => ARABIC_ABJAD[index % ARABIC_ABJAD.length]).join('، ');
  }
  return question;
}

function answerText(item) {
  if (!item.answer) return 'غير محددة';
  if (Array.isArray(item.options) && (item.type === 'mc' || isSameType(item.type, 'multi_select'))) {
    var indexes = answerIndexes(item, item.options);
    if (indexes.length) {
      return indexes.map(index => `(${ARABIC_ABJAD[index % ARABIC_ABJAD.length]}) ${item.options[index]}`).join('، ');
    }
  }
  return item.answer;
}

function buildAnswerKeyModelHtml(model) {
  var html = `<div class="card-title">الإجابات النموذجية - ${model.name}</div>`;
  model.blocks.forEach(block => {
    html += `<div class="answer-key-group"><strong>${escapeHtml(block.title)}</strong>`;
    block.items.forEach((item, index) => {
      var symbol = ARABIC_ABJAD[index % ARABIC_ABJAD.length];
      html += `<div>${symbol}- ${escapeHtml(answerText(item))}</div>`;
    });
    html += '</div>';
  });
  return html;
}

function toggleAnswerKey(elementId) {
  if (!currentGeneratedExam || !Array.isArray(currentGeneratedExam.modelA) ||
      !Array.isArray(currentGeneratedExam.modelB)) {
    showToast('يرجى إصدار الأسئلة أولاً');
    return;
  }
  var isManualExam = elementId === 'manual-answer-key';
  var preview = document.getElementById(isManualExam ? 'manual-exam-preview' : 'exam-preview');
  var modelAnswerKeys = preview.querySelectorAll('.model-answer-key');
  if (modelAnswerKeys.length !== 2) {
    showToast('يرجى عرض النموذجين أولاً');
    return;
  }
  var shouldShow = modelAnswerKeys[0].style.display === 'none';
  modelAnswerKeys.forEach((answerKey, index) => {
    if (shouldShow) {
      var model = index === 0
        ? {name: 'نموذج أ', blocks: currentGeneratedExam.modelA}
        : {name: 'نموذج ب', blocks: currentGeneratedExam.modelB};
      answerKey.innerHTML = buildAnswerKeyModelHtml(model);
      answerKey.style.display = 'block';
    } else {
      answerKey.style.display = 'none';
    }
  });
  preview.dataset.answerKeyVisible = shouldShow ? 'true' : 'false';
  if (!shouldShow) {
    return;
  }
}

function loadManualTopics() {
  var subject = document.getElementById('manual-subject').value;
  var container = document.getElementById('manual-topic-list');
  var topics = [...new Set(DB.filter(q => q.subject === subject).map(q => q.topic).filter(Boolean))].sort();
  container.innerHTML = '';
  document.getElementById('manual-question-panel').style.display = 'none';
  document.getElementById('manual-exam-preview').style.display = 'none';

  if (topics.length === 0) {
    container.textContent = 'لا توجد مفردات مسجلة لهذه المادة.';
    return;
  }

  topics.forEach(topic => {
    var label = document.createElement('label');
    label.className = 'manual-topic-option';
    var checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.className = 'manual-topic-check';
    checkbox.value = topic;
    var text = document.createElement('span');
    text.textContent = topic;
    label.appendChild(checkbox);
    label.appendChild(text);
    container.appendChild(label);
  });
}

function selectManualTopics(checked) {
  document.querySelectorAll('.manual-topic-check').forEach(input => input.checked = checked);
}

function restoreLastIssueForManual() {
  resolveLastIssuedExam()
    .then(issue => {
      if (!issue) {
        showToast('لا يوجد إصدار سابق لاسترجاعه');
        return;
      }
      manualEditingIssue = issue;
      manualEditingSelectedIds = new Set(issue.ids.map(Number));
      currentQuestionCohort = issue.cohort || currentQuestionCohort;
      document.getElementById('manual-cohort').value = currentQuestionCohort;
      document.getElementById('gen-cohort').value = currentQuestionCohort;
      document.getElementById('manual-subject').value = issue.subject;
      document.getElementById('gen-duration').value = String(issue.durationValue || '2');
      document.getElementById('manual-duration').value = String(issue.durationValue || '2');
      return restoreIssuedQuestionStatuses(issue).then(() => {
        showTab('manual');
        var originalTopics = new Set(DB.filter(question => manualEditingSelectedIds.has(question.id)).map(question => question.topic));
        document.querySelectorAll('.manual-topic-check').forEach(input => {
          input.checked = originalTopics.has(input.value);
        });
        document.getElementById('manual-question-list').innerHTML = '';
        showManualQuestions();
        showToast('تم استرجاع أسئلة الإصدار؛ عدّل الاختيارات ثم اضغط عرض الأسئلة');
      });
    })
    .catch(error => {
      manualEditingIssue = null;
      manualEditingSelectedIds = null;
      showToast('تعذر استرجاع الإصدار: ' + error.message);
    });
}

function showManualQuestions() {
  var existingChecks = Array.from(document.querySelectorAll('.manual-question-check'));
  if (manualEditingIssue && existingChecks.length) {
    manualEditingSelectedIds = new Set(existingChecks.filter(input => input.checked).map(input => Number(input.value)));
  }
  var subject = document.getElementById('manual-subject').value;
  var selectedTopics = Array.from(document.querySelectorAll('.manual-topic-check:checked')).map(input => input.value);
  if (selectedTopics.length === 0) {
    showToast('حدد مفردة واحدة على الأقل أولاً');
    return;
  }

  var questions = DB.filter(q => q.subject === subject && selectedTopics.includes(q.topic) &&
    (!manualEditingIssue || !q.used || manualEditingSelectedIds.has(q.id)));
  var typeGroups = [
    {id: 'mc', label: 'اختيار من متعدد'},
    {id: 'tf', label: 'صح وخطأ'},
    {id: 'fill', label: 'ملء الفراغات'},
    {id: 'essay', label: 'سؤال شرح ومقالي'},
    {id: 'multi_select', label: 'اختيار الإجابات الصحيحة'}
  ];
  var list = document.getElementById('manual-question-list');
  list.innerHTML = '';

  typeGroups.forEach(group => {
    var groupQuestions = questions.filter(q => isSameType(q.type, group.id));
    if (groupQuestions.length === 0) return;

    var section = document.createElement('section');
    section.className = 'manual-type-group';
    var heading = document.createElement('h3');
    heading.className = 'manual-type-title';
    heading.textContent = group.label;
    var rows = document.createElement('div');
    rows.className = 'manual-question-list';

    groupQuestions.forEach(question => {
      var label = document.createElement('label');
      label.className = 'manual-question-row';
      var checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'manual-question-check';
      checkbox.value = question.id;
      checkbox.checked = manualEditingSelectedIds ? manualEditingSelectedIds.has(question.id) : false;
      var content = document.createElement('span');
      content.className = 'manual-question-text';
      var text = document.createElement('strong');
      text.textContent = question.text;
      if (question.image) {
        var thumbnail = document.createElement('img');
        thumbnail.className = 'manual-question-thumbnail';
        thumbnail.src = question.image;
        thumbnail.alt = 'صورة مرفقة بالسؤال';
        content.appendChild(thumbnail);
      }
      var meta = document.createElement('span');
      meta.className = 'manual-question-meta';
      meta.textContent = 'المفردة: ' + question.topic + ' | السنة: ' + question.year + ' | الفصل: ' + question.semester;
      content.appendChild(text);
      content.appendChild(meta);
      label.appendChild(checkbox);
      label.appendChild(content);
      rows.appendChild(label);
    });

    section.appendChild(heading);
    section.appendChild(rows);
    list.appendChild(section);
  });

  document.getElementById('manual-question-panel').style.display = 'block';
  document.getElementById('manual-exam-preview').style.display = 'none';
}

function selectManualQuestions(checked) {
  document.querySelectorAll('.manual-question-check').forEach(input => input.checked = checked);
}

function previewManualExam() {
  var selectedIds = Array.from(document.querySelectorAll('.manual-question-check:checked')).map(input => Number(input.value));
  if (selectedIds.length === 0) {
    showToast('حدد سؤالاً واحداً على الأقل للمعاينة');
    return;
  }

  var selectedQuestions = selectedIds.map(id => DB.find(q => q.id === id)).filter(Boolean);
  var typeGroups = [
    {id: 'mc', title: 'السؤال الأول: اختر الإجابة الصحيحة من بين الخيارات'},
    {id: 'tf', title: 'السؤال الثاني: ضع علامة (صح) أو (خطأ) أمام العبارات'},
    {id: 'fill', title: 'السؤال الثالث: أملأ الفراغات الآتية بما يناسبها'},
    {id: 'essay', title: 'السؤال الرابع: أجب عن الأسئلة الشارحة الآتية'},
    {id: 'multi_select', title: 'السؤال الخامس: حدد الإجابات الصحيحة للأسئلة التالية'}
  ];
  var modelA = typeGroups.map(group => ({
    title: group.title,
    items: selectedQuestions.filter(q => isSameType(q.type, group.id))
  })).filter(group => group.items.length > 0);
  var modelB = modelA.map(group => ({
    title: group.title,
    items: shuffleArray(JSON.parse(JSON.stringify(group.items))).map(question => {
      return shuffleQuestionOptions(question);
    })
  }));

  currentGeneratedExam = {
    subject: document.getElementById('manual-subject').value,
    year: 'حسب الأسئلة المختارة',
    semester: 'حسب الأسئلة المختارة',
    pool: 'mix',
    order: 'random',
    durationValue: document.getElementById('manual-duration').value,
    duration: getExamDurationLabel(document.getElementById('manual-duration').value),
    examDate: getExamDateLabel(),
    modelA: modelA,
    modelB: modelB,
    ids: selectedIds,
    cohort: currentQuestionCohort,
    sourceMode: 'manual'
  };
  rememberLastIssuedExam('manual');
  manualEditingIssue = null;
  manualEditingSelectedIds = null;
  renderExamPreview(modelA, modelB, 'manual-exam-preview', 'manual-exam-container');
}

function generateExam() {
  var subj = document.getElementById('gen-subject').value;
  var yr = document.getElementById('gen-year').value;
  var sem = document.getElementById('gen-sem').value;
  var pool = document.getElementById('gen-pool').value;
  var isRandom = document.getElementById('gen-order').value === 'random';

  var typesConfig = [
    { id: 'mc', chk: 'chk-mc', num: 'num-mc', label: 'اختيار من متعدد', title: 'السؤال الأول: اختر الإجابة الصحيحة من بين الخيارات' },
    { id: 'tf', chk: 'chk-tf', num: 'num-tf', label: 'صح وخطأ', title: 'السؤال الثاني: ضع علامة (صح) أو (خطأ) أمام العبارات' },
    { id: 'fill', chk: 'chk-fill', num: 'num-fill', label: 'ملء الفراغات', title: 'السؤال الثالث: أملأ الفراغات الآتية بما يناسبها' },
    { id: 'essay', chk: 'chk-essay', num: 'num-essay', label: 'سؤال شرح ومقالي', title: 'السؤال الرابع: أجب عن الأسئلة الشارحة الآتية' },
    { id: 'multi_select', chk: 'chk-multi', num: 'num-multi', label: 'اختيار الإجابات الصحيحة', title: 'السؤال الخامس: حدد الإجابات الصحيحة للأسئلة التالية' }
  ];

  if (!generatedIssuedEditMode) {
    for (var i = 0; i < typesConfig.length; i++) {
      var conf = typesConfig[i];
      if (document.getElementById(conf.chk).checked) {
        var reqCount = parseInt(document.getElementById(conf.num).value) || 0;
        var availCount = currentTypeCounts[conf.id] || 0;
        if (reqCount > availCount) {
          alert(`تنبيه: عدد الأسئلة المطلوبة في [${conf.label}] هو (${reqCount}) سؤالاً، بينما المتاح حالياً هو (${availCount}) فقط!\n\nيرجى تقليل العدد أو إضافة أسئلة جديدة ليتسنى لك التوليد.`);
          return;
        }
      }
    }
  }

  var filtered = DB.filter(q => {
    if (q.subject !== subj) return false;
    if (pool === 'unused' && q.used) return false;
    if (pool === 'used' && !q.used) return false;
    if (yr === 'all1') return q.year == 1;
    if (yr === 'all2') return q.year == 2;
    if (yr === 'all3') return q.year == 3;
    if (yr !== 'all') {
      if (q.year != yr) return false;
      if (q.semester != sem) return false;
    }
    return true;
  });

  var blocksA = [];
  var blocksB = [];
  var usedIds = [];
  var selectedIssuedIds = new Set(Array.from(document.querySelectorAll('.generated-issued-check:checked')).map(input => Number(input.value)));
  if (generatedIssuedEditMode && selectedIssuedIds.size === 0) {
    showToast('حدد سؤالاً واحداً على الأقل قبل الإصدار');
    return;
  }

  typesConfig.forEach(conf => {
    if (document.getElementById(conf.chk).checked) {
      var count = parseInt(document.getElementById(conf.num).value) || 0;
      if (!generatedIssuedEditMode && count <= 0) return;

      var qList = generatedIssuedEditMode
        ? DB.filter(q => selectedIssuedIds.has(q.id) && q.subject === subj &&
            issueMatchesGeneratedScope(q, {year: yr, semester: sem}) && isSameType(q.type, conf.id))
        : filtered.filter(q => isSameType(q.type, conf.id)).slice(0, count);
      if (generatedIssuedEditMode) document.getElementById(conf.num).value = qList.length;
      
      if (qList.length > 0) {
        qList.forEach(item => usedIds.push(item.id));

        blocksA.push({ title: conf.title, items: JSON.parse(JSON.stringify(qList)) });

        var qListB = JSON.parse(JSON.stringify(qList));
        if (isRandom) {
          qListB = shuffleArray(qListB);
          qListB.forEach(item => {
            shuffleQuestionOptions(item);
          });
        }
        blocksB.push({ title: conf.title, items: qListB });
      }
    }
  });

  currentGeneratedExam = {
    subject: subj,
    year: yr,
    semester: sem,
    pool: pool,
    order: document.getElementById('gen-order').value,
    durationValue: document.getElementById('gen-duration').value,
    duration: getExamDurationLabel(document.getElementById('gen-duration').value),
    examDate: getExamDateLabel(),
    modelA: blocksA,
    modelB: blocksB,
    ids: usedIds,
    sourceMode: 'generated',
    cohort: currentQuestionCohort
  };

  rememberLastIssuedExam('generated');
  generatedIssuedEditMode = false;
  generatedIssuedEditIssue = null;
  document.getElementById('generated-issued-editor').style.display = 'none';
  renderExamPreview(blocksA, blocksB);
}

function renderExamPreview(blocksA, blocksB, previewId = 'exam-preview', containerId = 'exam-container', shouldScroll = true, resetAnswerKey = true) {
  var container = document.getElementById(containerId);
  container.innerHTML = '';
  var preview = document.getElementById(previewId);
  preview.style.display = 'block';
  var answerKey = document.getElementById(previewId === 'manual-exam-preview' ? 'manual-answer-key' : 'generated-answer-key');
  if (answerKey && resetAnswerKey) {
    answerKey.style.display = 'none';
    answerKey.innerHTML = '';
    preview.dataset.answerKeyVisible = 'false';
  }
  var showAnswerKeys = preview.dataset.answerKeyVisible === 'true';

  [ {name: 'نموذج أ', blocks: blocksA}, {name: 'نموذج ب', blocks: blocksB} ].forEach(m => {
    var html = `<div class="exam-form-wrap">
      <div class="exam-form-header">
        <div class="exam-header-center">
          <h2>جمهورية العراق ـــــ وزارة الدفاع</h2>
          <h3>/ جامعة الدفاع للدراسات العسكرية / الكلية العسكرية جناح التدريس /</h3>
          <p>المادة: ${escapeHtml(currentGeneratedExam.subject)}</p>
          <p><strong>${m.name}</strong></p>
        </div>
        <div class="exam-header-meta">
          <span>التاريخ: ${currentGeneratedExam.examDate || getExamDateLabel()}</span>
          <span>الوقت: ${currentGeneratedExam.duration || 'ساعتان'}</span>
        </div>
      </div>`;
    
    m.blocks.forEach(b => {
      html += `<div style="font-weight:bold;margin-top:.8rem;border-bottom:1px solid #000;padding-bottom:.2rem">${escapeHtml(b.title)}</div>`;
      b.items.forEach((item, idx) => {
        var abjad = ARABIC_ABJAD[idx % ARABIC_ABJAD.length];
        html += `<div style="margin-right:.5rem;margin-top:.4rem"><strong>${abjad}-</strong> ${escapeHtml(item.text)}</div>`;
        if (item.image) {
          var imageSize = getQuestionImageSize(item.image_size);
          html += `<img class="exam-question-image" src="${item.image}" alt="صورة السؤال" style="width:${imageSize}%">`;
        }
        
        if (item.type === 'mc' && Array.isArray(item.options)) {
          var optsHtml = item.options.map((opt, i) => `(${ARABIC_ABJAD[i % ARABIC_ABJAD.length]}) ${escapeHtml(opt)}`).join('&nbsp;&nbsp;&nbsp;&nbsp;');
          html += `<div style="margin-right:1.5rem;font-size:.9rem;margin-top:.2rem">${optsHtml}</div>`;
        } else if (isSameType(item.type, 'multi_select') && Array.isArray(item.options)) {
          html += `<div style="margin-right:1.5rem;font-size:.85rem;margin-top:.2rem">الخيارات الأبجدية المتاحة:</div>`;
          item.options.forEach((opt, i) => {
            html += `<div style="margin-right:2.5rem;font-size:.85rem">[${ARABIC_ABJAD[i % ARABIC_ABJAD.length]}] ${escapeHtml(opt)}</div>`;
          });
          html += `<div style="margin-right:1.5rem;font-size:.85rem;margin-top:.2rem">إجابة الطالب (أكتب رموز الخيارات الصحيحة فقط): [ ________________________ ]</div>`;
        } else if (item.type === 'essay') {
          html += `<div style="margin-right:1.5rem;margin-top:.3rem;color:#555">الإجابة: __________________________________________________________________________________________________</div>`;
        }
      });
    });
    html += `<div class="answer-key model-answer-key no-print" style="display:${showAnswerKeys ? 'block' : 'none'}">${showAnswerKeys ? buildAnswerKeyModelHtml(m) : ''}</div>`;
    html += '</div>';
    container.innerHTML += html;
  });
  if (shouldScroll) preview.scrollIntoView({behavior: 'smooth', block: 'start'});
}

function prepareExamPreviewForPrint() {
  if (!currentGeneratedExam ||
      !Array.isArray(currentGeneratedExam.modelA) ||
      !Array.isArray(currentGeneratedExam.modelB)) {
    showToast('بيانات نموذجي الامتحان غير مكتملة');
    return false;
  }

  var isManualExam = currentGeneratedExam.sourceMode === 'manual';
  var previewId = isManualExam ? 'manual-exam-preview' : 'exam-preview';
  var containerId = isManualExam ? 'manual-exam-container' : 'exam-container';
  var preview = document.getElementById(previewId);
  var container = document.getElementById(containerId);
  document.getElementById('exam-preview').style.display = isManualExam ? 'none' : 'block';
  document.getElementById('manual-exam-preview').style.display = isManualExam ? 'block' : 'none';
  renderExamPreview(
    currentGeneratedExam.modelA,
    currentGeneratedExam.modelB,
    previewId,
    containerId,
    false,
    false
  );
  if (document.querySelectorAll('#' + containerId + ' .exam-form-wrap').length !== 2) {
    showToast('تعذر تجهيز النموذجين للطباعة');
    return false;
  }
  var printPreview = document.getElementById('print-exam-preview');
  printPreview.innerHTML = '';
  container.querySelectorAll('.exam-form-wrap').forEach(form => {
    var pair = document.createElement('div');
    pair.className = 'print-model-pair';
    var formCopy = form.cloneNode(true);
    formCopy.querySelectorAll('.model-answer-key').forEach(answerKey => answerKey.remove());
    pair.appendChild(formCopy);
    if (preview.dataset.answerKeyVisible === 'true') {
      var answerKey = form.querySelector('.model-answer-key').cloneNode(true);
      answerKey.classList.remove('no-print');
      answerKey.style.display = 'block';
      pair.appendChild(answerKey);
    }
    printPreview.appendChild(pair);
  });
  if (printPreview.querySelectorAll('.print-model-pair').length !== 2) {
    printPreview.innerHTML = '';
    showToast('تعذر نسخ النموذجين إلى نافذة الطباعة');
    return false;
  }
  return true;
}

function showAndPrintExamModels() {
  if (!currentGeneratedExam) { showToast('يرجى إصدار الأسئلة أولاً'); return; }
  if (!prepareExamPreviewForPrint()) return;
  if (printDialogPending) return;

  var issuedIds = currentGeneratedExam.ids.slice();
  printDialogPending = true;
  var handleAfterPrint = function() {
    printDialogPending = false;
    if (!confirm('هل تمت طباعة النموذجين بنجاح؟')) {
      showToast('تم إلغاء طباعة النموذجين');
      return;
    }
    if (!confirm('هل تود خصم الأسئلة الصادرة من رصيد المتاح؟')) {
      showToast('تمت الطباعة دون خصم الأسئلة من الرصيد المتاح');
      return;
    }

    fetch('/api/questions/mark-used', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        ids: issuedIds,
        cohort: currentGeneratedExam.cohort || currentQuestionCohort,
        source_mode: currentGeneratedExam.sourceMode || 'exam'
      })
    }).then(response => {
      if (!response.ok) throw new Error('تعذر خصم الأسئلة');
      showToast('تم خصم الأسئلة من الرصيد المتاح');
      return loadDB();
    }).catch(error => showToast(error.message));
  };

  window.addEventListener('afterprint', handleAfterPrint, {once: true});
  var requestFrame = window.requestAnimationFrame || function(callback) { return setTimeout(callback, 0); };
  requestFrame(function() {
    requestFrame(function() {
      try {
        window.print();
      } catch (error) {
        window.removeEventListener('afterprint', handleAfterPrint);
        printDialogPending = false;
        showToast('تعذرت طباعة النموذجين');
      }
    });
  });
}

function downloadWordDocument(examData) {
  return fetch('/api/export-word', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(examData)
  }).then(response => response.json().then(data => ({response: response, data: data})))
    .then(result => {
      if (!result.response.ok || result.data.status !== 'success') {
        throw new Error(result.data.message || 'تعذر إنشاء مستند Word');
      }
      showToast('تم حفظ مستند Word في: ' + result.data.path);
      return result.data;
  });
}

function exportToWord() {
  if (!currentGeneratedExam) { showToast('يرجى إصدار الأسئلة أولاً'); return; }
  if (!Array.isArray(currentGeneratedExam.modelA) || !Array.isArray(currentGeneratedExam.modelB)) {
    showToast('بيانات نموذجي الامتحان غير مكتملة');
    return;
  }
  var examData = currentGeneratedExam;
  if (window.pywebview && window.pywebview.api) {
    Promise.resolve()
      .then(() => window.pywebview.api.export_word_save_as(JSON.stringify(examData)))
      .then(result => {
        if (result && result.status === 'success') {
          showToast('تم حفظ مستند الـ Word بنجاح ✓');
        } else if (result && result.status === 'cancelled') {
          showToast('تم إلغاء حفظ مستند Word');
        } else {
          return downloadWordDocument(examData);
        }
      })
      .catch(() => downloadWordDocument(examData)
        .catch(error => showToast('فشل تصدير Word: ' + error.message)));
  } else {
    downloadWordDocument(examData)
      .catch(error => showToast('فشل تصدير Word: ' + error.message));
  }
}

function renderStats() {
  var wrap = document.getElementById('stats-table');
  wrap.innerHTML = '';
  SUBJECTS.forEach(s => {
    var total = DB.filter(q => q.subject === s).length;
    var used = DB.filter(q => q.subject === s && q.used).length;
    var unused = total - used;
    wrap.innerHTML += `<div class="q-item"><div><b>${escapeHtml(s)}</b></div><div>الكلي: ${total} | المستخدم: ${used} | المتاح: ${unused}</div></div>`;
  });
}

function examApiJson(url, options) {
  return fetch(url, options).then(response => response.json().then(data => {
    if (!response.ok || data.status === 'error') {
      throw new Error(data.message || 'تعذر إكمال العملية');
    }
    return data;
  }));
}

function loadExamRoster() {
  return examApiJson('/api/exams/roster')
    .then(data => {
      importedExamRoster = data;
      var summary = document.getElementById('exam-roster-summary');
      summary.textContent = data.total
        ? `عدد الطلاب المستوردين: ${data.total}. أعداد السرايا: ${[1,2,3,4].map(company => company + ': ' + (data.counts[String(company)] || 0)).join(' | ')}`
        : 'لم تُستورد قائمة طلاب بعد.';
      var platoon = document.getElementById('issue-platoon');
      var selected = platoon.value;
      platoon.innerHTML = '<option value="">كل الفصائل</option>';
      (data.platoons || []).forEach(name => {
        var option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        platoon.appendChild(option);
      });
      platoon.value = (data.platoons || []).includes(selected) ? selected : '';
    })
    .catch(error => {
      document.getElementById('exam-roster-summary').textContent = 'تعذر تحميل قائمة الطلاب: ' + error.message;
    });
}

function importExamRoster() {
  var input = document.getElementById('exam-roster-file');
  var files = Array.from(input.files || []);
  var status = document.getElementById('exam-issue-status');
  if (!files.length) {
    status.textContent = 'اختر ملف Excel واحداً على الأقل بصيغة XLSX أو XLSM أو XLS.';
    return;
  }
  var unsupported = files.find(file => {
    var extension = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
    return !['.xlsx', '.xlsm', '.xls'].includes(extension);
  });
  if (unsupported) {
    status.textContent = 'صيغة الملف غير مدعومة: ' + unsupported.name;
    return;
  }
  if (importedExamRoster && importedExamRoster.total &&
      !confirm('سيستبدل هذا الاستيراد قائمة الطلاب الحالية. هل تريد المتابعة؟')) return;
  var formData = new FormData();
  files.forEach(file => formData.append('files', file));
  status.textContent = 'جارٍ التحقق من ' + files.length + ' ملف Excel...';
  fetch('/api/exams/roster', {method:'POST', body:formData})
    .then(response => response.json().then(data => ({response:response, data:data})))
    .then(result => {
      if (!result.response.ok || result.data.status !== 'success') {
        throw new Error(result.data.message || 'تعذر استيراد قائمة الطلاب');
      }
      status.textContent = 'تم استيراد ' + result.data.total + ' طالباً من ' +
        (result.data.files || files.map(file => file.name)).length + ' ملف.';
      return loadExamRoster();
    })
    .catch(error => { status.textContent = 'فشل استيراد Excel: ' + error.message; });
}

function openExamBookletPreview(index) {
  if (!currentExamBooklets.length || !currentExamIssueSnapshot || !currentGeneratedExam) {
    showToast('أصدر الدفاتر أولاً حتى تتمكن من معاينتها.');
    return;
  }
  var select = document.getElementById('exam-preview-booklet-select');
  if (!select.options.length) {
    currentExamBooklets.forEach((booklet, bookletIndex) => {
      var option = document.createElement('option');
      option.value = String(bookletIndex);
      option.textContent = `الدفتر ${booklet.serial_number} | السرية ${booklet.company} | الفصيل ${booklet.platoon} | الرقم العسكري ${booklet.military_number} | نموذج ${booklet.model}`;
      select.appendChild(option);
    });
  }
  document.getElementById('exam-preview-modal').classList.add('open');
  document.getElementById('exam-preview-modal').setAttribute('aria-hidden', 'false');
  renderExamBookletPreview(index == null ? select.value || '0' : index);
}

function closeExamBookletPreview() {
  var modal = document.getElementById('exam-preview-modal');
  modal.classList.remove('open');
  modal.setAttribute('aria-hidden', 'true');
}

function renderExamBookletPreview(value) {
  if (!currentExamBooklets.length || !currentExamIssueSnapshot || !currentGeneratedExam) return;
  var index = Number(value);
  if (!Number.isInteger(index) || index < 0 || index >= currentExamBooklets.length) return;
  var select = document.getElementById('exam-preview-booklet-select');
  select.value = String(index);
  document.getElementById('exam-preview-sheet').innerHTML = renderExamBooklet(
    currentExamBooklets[index], currentExamIssueSnapshot, currentGeneratedExam
  );
  document.getElementById('exam-preview-position').textContent =
    `الدفتر ${index + 1} من ${currentExamBooklets.length}`;
  document.getElementById('exam-preview-previous').disabled = index === 0;
  document.getElementById('exam-preview-next').disabled = index === currentExamBooklets.length - 1;
}

function moveExamBookletPreview(offset) {
  var select = document.getElementById('exam-preview-booklet-select');
  renderExamBookletPreview(Number(select.value) + offset);
}

function printQuestionChoices(item) {
  if (item.type === 'essay') return [];
  if (item.type === 'tf' && (!Array.isArray(item.options) || !item.options.length)) return ['صح', 'خطأ'];
  return Array.isArray(item.options) ? item.options : [];
}

function renderExamBooklet(booklet, issue, exam) {
  var modelName = booklet.model === 'A' ? 'نموذج A' : 'نموذج B';
  var blocks = booklet.model === 'A' ? exam.modelA : exam.modelB;
  var html = `<table class="exam-booklet">
    <thead><tr><th>
    <img class="exam-booklet-qr" src="${booklet.qr_data}" alt="QR دفتر الامتحان">
    <header class="exam-form-header">
      <h2>جمهورية العراق ـــــ وزارة الدفاع</h2>
      <h3>الكلية العسكرية</h3>
      <p>${escapeHtml(issue.course)} | الفوج ${escapeHtml(issue.cohort)} | السرية ${booklet.company} | الفصيل ${escapeHtml(booklet.platoon)}</p>
      <p>المادة: ${escapeHtml(issue.subject)} | السنة ${issue.year} | النصف ${issue.semester}</p>
      <p><strong>${modelName}</strong></p>
      <div class="exam-header-meta">
        <span>تسلسل الطالب: ${escapeHtml(String(booklet.student_sequence || booklet.serial_number))}</span>
        <span>التاريخ: ${escapeHtml(issue.exam_date)} | الوقت: ${escapeHtml(issue.exam_time)}</span>
      </div>
    </header>
    </th></tr></thead><tbody>`;
  var questionNumber = 0;
  blocks.forEach(block => {
    html += `<tr><td><h3 style="margin-top:.8rem;border-bottom:1px solid #000">${escapeHtml(block.title || '')}</h3></td></tr>`;
    (block.items || []).forEach(item => {
      questionNumber++;
      html += `<tr class="exam-booklet-question"><td><strong>${questionNumber}-</strong> ${escapeHtml(item.text || '')}`;
      if (item.image) {
        html += `<img class="exam-question-image" src="${item.image}" alt="صورة السؤال" style="width:${getQuestionImageSize(item.image_size)}%">`;
      }
      var choices = printQuestionChoices(item);
      if (choices.length) {
        html += '<div class="omr-options">';
        choices.forEach((choice, index) => {
          html += `<div class="omr-option"><span class="omr-bubble" aria-hidden="true"></span><span class="omr-option-text">${escapeHtml(ARABIC_ABJAD[index % ARABIC_ABJAD.length])}- ${escapeHtml(choice)}</span></div>`;
        });
        html += '</div>';
        if (item.type === 'multi_select' || item.type === 'multi') {
          html += '<div style="font-size:.8rem">ظلل جميع الاختيارات الصحيحة.</div>';
        }
      } else if (item.type === 'essay' || item.type === 'fill') {
        html += '<div style="margin:.4rem 1rem">الإجابة: __________________________________________________________</div>';
      }
      html += '</td></tr>';
    });
  });
  return html + '</tbody></table>';
}

function createExamBooklets() {
  var status = document.getElementById('exam-issue-status');
  if (!currentGeneratedExam || !Array.isArray(currentGeneratedExam.modelA) ||
      !Array.isArray(currentGeneratedExam.modelB) || !currentGeneratedExam.modelA.length ||
      !currentGeneratedExam.modelB.length) {
    status.textContent = 'أصدر نموذجي الأسئلة A وB من نافذة التوليد أولاً.';
    return;
  }
  var subject = document.getElementById('issue-subject').value;
  if (currentGeneratedExam.subject !== subject) {
    status.textContent = 'المادة المحددة لا تطابق مادة نموذجي الأسئلة الصادرين.';
    return;
  }
  var counts = {};
  for (var company = 1; company <= 4; company++) {
    counts[String(company)] = Number(document.getElementById('issue-count-' + company).value || 0);
    if (!Number.isInteger(counts[String(company)]) || counts[String(company)] < 0) {
      status.textContent = 'أدخل أعداداً صحيحة غير سالبة للسرايا الأربع.';
      return;
    }
  }
  var payload = {
    course: document.getElementById('issue-course').value.trim(),
    cohort: document.getElementById('issue-cohort').value,
    subject: subject,
    platoon: document.getElementById('issue-platoon').value,
    year: document.getElementById('issue-year').value,
    semester: document.getElementById('issue-semester').value,
    exam_date: document.getElementById('issue-date').value,
    exam_time: document.getElementById('issue-time').value,
    counts: counts,
    modelA: currentGeneratedExam.modelA,
    modelB: currentGeneratedExam.modelB
  };
  status.textContent = 'جارٍ إعداد الدفاتر والرموز الفريدة...';
  examApiJson('/api/exams/issues', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)
  }).then(data => {
    currentExamIssueId = data.issue_id;
    var issue = Object.assign({}, payload, {id:data.issue_id});
    currentExamBooklets = data.booklets;
    currentExamIssueSnapshot = issue;
    document.getElementById('exam-preview-booklet-select').innerHTML = '';
    var preview = document.getElementById('exam-issue-summary');
    preview.textContent = `تم إعداد ${data.booklets.length} دفترًا؛ A وB موزعان بالتناوب داخل كل سرية. الاسم غير مطبوع، ويظهر تسلسل الطالب فقط.`;
    var printArea = document.getElementById('booklet-print');
    printArea.innerHTML = data.booklets.map(booklet => renderExamBooklet(booklet, issue, currentGeneratedExam)).join('');
    document.getElementById('exam-issue-preview').style.display = 'block';
    status.textContent = 'اكتمل إصدار الدفاتر. افتح نافذة المعاينة وراجع النماذج والبيانات قبل الطباعة.';
    openExamBookletPreview(0);
  }).catch(error => {
    status.textContent = 'تعذر إصدار الدفاتر: ' + error.message;
  });
}

function printExamBooklets() {
  if (!currentExamIssueId || !document.getElementById('booklet-print').children.length) {
    showToast('ولّد الدفاتر قبل الطباعة.');
    return;
  }
  if (printDialogPending) return;
  printDialogPending = true;
  var printArea = document.getElementById('booklet-print');
  printArea.style.display = 'block';
  var afterPrint = function() {
    printDialogPending = false;
    printArea.style.display = 'none';
    var printed = confirm('هل تمت طباعة الدفاتر بنجاح؟');
    var transfer = printed && confirm('هل تريد ترحيل مفتاحي النموذجين A وB إلى نافذة التصحيح؟');
    examApiJson('/api/exams/issues/' + encodeURIComponent(currentExamIssueId) + '/transfer-key', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({printed:printed, transfer:transfer})
    }).then(() => {
      showToast(transfer ? 'تم حفظ مفتاح الإجابة للتصحيح.' : 'لم يتم ترحيل مفتاح الإجابة.');
      loadCorrectionIssues();
    }).catch(error => showToast('تعذر تحديث حالة مفتاح الإجابة: ' + error.message));
  };
  window.addEventListener('afterprint', afterPrint, {once:true});
  try {
    window.print();
  } catch (error) {
    window.removeEventListener('afterprint', afterPrint);
    printDialogPending = false;
    printArea.style.display = 'none';
    showToast('تعذرت طباعة الدفاتر: ' + error.message);
  }
}

function loadCorrectionIssues() {
  return examApiJson('/api/exams/issues')
    .then(data => {
      var select = document.getElementById('correction-issue-select');
      var previous = select.value;
      select.innerHTML = '<option value="">اختر إصداراً</option>';
      data.issues.forEach(issue => {
        var option = document.createElement('option');
        option.value = issue.id;
        option.textContent = `${issue.subject} | ${issue.course} | ${issue.exam_date} | الفوج ${issue.cohort}`;
        select.appendChild(option);
      });
      if (data.issues.some(issue => issue.id === previous)) {
        select.value = previous;
        loadCorrectionIssue(previous);
      }
    })
    .catch(error => {
      document.getElementById('correction-issue-status').textContent = 'تعذر تحميل الإصدارات: ' + error.message;
    });
}

function loadCorrectionIssue(issueId) {
  currentCorrectionIssue = null;
  document.getElementById('correction-scoring-card').style.display = 'none';
  document.getElementById('correction-scan-card').style.display = 'none';
  if (!issueId) return Promise.resolve();
  return examApiJson('/api/exams/issues/' + encodeURIComponent(issueId))
    .then(data => {
      currentCorrectionIssue = data.issue;
      var issue = data.issue;
      document.getElementById('correction-key-transfer').style.display = 'block';
      document.getElementById('correction-key-import').style.display = 'block';
      document.getElementById('correction-key-state').textContent = issue.key_transferred
        ? ' المفتاح معتمد.'
        : ' المفتاح غير مرحّل بعد؛ اختر الترحيل بعد تأكيد الطباعة أو استورد مفتاحاً.';
      document.querySelector('#correction-key-transfer button').style.display = issue.key_transferred ? 'none' : 'inline-flex';
      document.getElementById('correction-issue-status').textContent =
        `المادة ${issue.subject}، عدد الدفاتر ${issue.booklets.length}، المكتمل تصحيحه ${issue.booklets.filter(booklet => booklet.status === 'corrected').length}.`;
      renderCorrectionScoring(issue);
      document.getElementById('correction-scoring-card').style.display = 'block';
      document.getElementById('correction-scan-card').style.display = issue.key_transferred ? 'block' : 'none';
      examScanResults = issue.booklets.filter(booklet => booklet.status === 'review');
      renderExamScanReview(examScanResults);
    })
    .catch(error => {
      document.getElementById('correction-issue-status').textContent = 'تعذر تحميل الإصدار: ' + error.message;
    });
}

function renderCorrectionScoring(issue) {
  ['A', 'B'].forEach(model => {
    var blocks = model === 'A' ? issue.modelA : issue.modelB;
    var saved = (issue.grading || {})[model] || [];
    var container = document.getElementById('grading-model-' + model);
    container.innerHTML = '';
    blocks.forEach((block, blockIndex) => {
      var savedBlock = saved[blockIndex] || {};
      var manual = new Set(savedBlock.manual || []);
      var section = document.createElement('section');
      section.className = 'exam-scoring-block';
      var title = document.createElement('strong');
      title.textContent = block.title || ('المجموعة ' + (blockIndex + 1));
      section.appendChild(title);
      var points = document.createElement('input');
      points.type = 'number';
      points.min = '0';
      points.step = '0.25';
      points.dataset.blockPoints = model + ':' + blockIndex;
      points.value = savedBlock.points == null ? '' : savedBlock.points;
      points.setAttribute('aria-label', 'الدرجة الكلية للمجموعة');
      var pointsLabel = document.createElement('label');
      pointsLabel.className = 'lbl';
      pointsLabel.textContent = 'الدرجة الكلية للسؤال الرئيسي / المجموعة';
      pointsLabel.appendChild(points);
      section.appendChild(pointsLabel);
      (block.items || []).forEach((item, itemIndex) => {
        var label = document.createElement('label');
        label.className = 'exam-scoring-question';
        var checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.dataset.manualQuestion = `${model}:${blockIndex}:${itemIndex}`;
        checkbox.checked = manual.has(itemIndex) || item.type === 'essay';
        var text = document.createElement('span');
        text.textContent = item.text || ('السؤال ' + (itemIndex + 1));
        label.append(checkbox, text);
        section.appendChild(label);
      });
      container.appendChild(section);
    });
  });
}

function approveCorrectionKey() {
  if (!currentCorrectionIssue) return;
  if (!confirm('هل تم طباعة هذا الإصدار بنجاح، وتريد اعتماد مفتاح إجابته للتصحيح؟')) return;
  examApiJson('/api/exams/issues/' + encodeURIComponent(currentCorrectionIssue.id) + '/transfer-key', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body:JSON.stringify({printed:true, transfer:true})
  }).then(() => loadCorrectionIssue(currentCorrectionIssue.id))
    .catch(error => showToast('تعذر اعتماد المفتاح: ' + error.message));
}

function importCorrectionKey() {
  if (!currentCorrectionIssue) return;
  var file = document.getElementById('correction-key-file').files[0];
  if (!file) {
    showToast('اختر ملف JSON لمفتاح الإجابة.');
    return;
  }
  file.text().then(text => JSON.parse(text))
    .then(key => examApiJson('/api/exams/issues/' + encodeURIComponent(currentCorrectionIssue.id) + '/answer-key', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(key)
    }))
    .then(() => loadCorrectionIssue(currentCorrectionIssue.id))
    .then(() => showToast('تم استيراد مفتاح الإجابة واعتماده لهذا الإصدار.'))
    .catch(error => showToast('تعذر استيراد المفتاح: ' + error.message));
}

function saveCorrectionGrading() {
  if (!currentCorrectionIssue) return;
  var grading = {A:[], B:[]};
  ['A', 'B'].forEach(model => {
    var blocks = model === 'A' ? currentCorrectionIssue.modelA : currentCorrectionIssue.modelB;
    blocks.forEach((block, blockIndex) => {
      var pointInput = document.querySelector(`[data-block-points="${model}:${blockIndex}"]`);
      var manual = [];
      (block.items || []).forEach((item, itemIndex) => {
        var checkbox = document.querySelector(`[data-manual-question="${model}:${blockIndex}:${itemIndex}"]`);
        if (checkbox && checkbox.checked) manual.push(itemIndex);
      });
      grading[model].push({points:pointInput ? pointInput.value : '', manual:manual});
    });
  });
  examApiJson('/api/exams/issues/' + encodeURIComponent(currentCorrectionIssue.id) + '/grading', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({grading:grading})
  }).then(() => {
    showToast('تم حفظ توزيع الدرجات.');
    return loadCorrectionIssue(currentCorrectionIssue.id);
  }).catch(error => showToast('تعذر حفظ الدرجات: ' + error.message));
}

function scanExamFiles() {
  if (!currentCorrectionIssue || !currentCorrectionIssue.key_transferred) {
    showToast('اعتمد مفتاح الإجابة لهذا الإصدار أولاً.');
    return;
  }
  if (!currentCorrectionIssue.grading ||
      ['A', 'B'].some(model => !Array.isArray(currentCorrectionIssue.grading[model]) ||
        currentCorrectionIssue.grading[model].length !== currentCorrectionIssue['model' + model].length)) {
    showToast('أدخل درجات المجموعات للنموذجين واحفظها قبل المسح.');
    return;
  }
  var files = document.getElementById('exam-scan-file').files;
  if (!files || !files.length) {
    showToast('اختر ملفات المسح أولاً (PDF أو صور).');
    return;
  }
  var formData = new FormData();
  for (var fileIndex = 0; fileIndex < files.length; fileIndex++) {
    formData.append('files', files[fileIndex]);
  }
  document.getElementById('exam-scan-status').textContent = 'جارٍ قراءة QR والفقاعات من ' + files.length + ' ملف...';
  examScanResults = [];
  fetch('/api/exams/scan', {method:'POST', body:formData})
    .then(response => response.json().then(data => ({response:response, data:data})))
    .then(result => {
      if (!result.response.ok || result.data.status !== 'success') {
        throw new Error(result.data.message || 'تعذر تحليل ملف المسح');
      }
      examScanResults = result.data.booklets;
      document.getElementById('exam-scan-status').textContent =
        `تم تحليل ${examScanResults.length} دفتر. راجع الإجابات الملتبسة وأدخل درجات الأسئلة اليدوية.`;
      var problemsHtml = (result.data.page_problems || [])
        .map(problem => `<div class="word-import-warning">${problem.page > 0 ? 'صفحة ' + problem.page + ': ' : ''}${escapeHtml(problem.message)}</div>`).join('');
      if (result.data.engines && result.data.engines.ocr === false) {
        problemsHtml += '<div class="word-import-warning">تنبيه: محرك OCR الاحتياطي (Tesseract) غير مثبت؛ التعرف على الدفاتر يعتمد على رمز QR فقط.</div>';
      }
      document.getElementById('exam-scan-problems').innerHTML = problemsHtml;
      renderExamScanReview(examScanResults);
    })
    .catch(error => {
      document.getElementById('exam-scan-status').textContent = 'فشل المسح: ' + error.message;
    });
}

function renderExamScanReview(results) {
  var container = document.getElementById('exam-review-list');
  container.innerHTML = '';
  results.forEach((result, index) => {
    var card = document.createElement('section');
    card.className = 'exam-review-card';
    card.dataset.reviewToken = result.qr_token;
    var title = document.createElement('h3');
    title.textContent = `الدفتر ${result.serial_number} | السرية ${result.company} | الرقم العسكري ${result.military_number} | النموذج ${result.model}` +
      (result.identified_by === 'ocr' ? ' | تعُرف عليه عبر OCR بدل QR — راجعه بتمعن' : '');
    card.appendChild(title);
    (result.page_images || []).forEach(page => {
      var image = document.createElement('img');
      image.src = page.data;
      image.alt = 'صورة صفحة المسح رقم ' + page.page;
      image.title = image.alt;
      card.appendChild(image);
    });
    if (!result.review.length) {
      var clear = document.createElement('p');
      clear.textContent = 'لم تُرصد أسئلة بحاجة إلى مراجعة يدوية.';
      card.appendChild(clear);
    }
    result.review.forEach(entry => {
      var item = document.createElement('div');
      item.className = 'exam-review-card';
      var prompt = document.createElement('p');
      prompt.textContent = entry.prompt + ' — ' + entry.reason;
      item.appendChild(prompt);
      var key = entry.question_key;
      if (entry.manual_score) {
        var label = document.createElement('label');
        label.className = 'lbl';
        label.textContent = 'الدرجة اليدوية (الحد الأعلى ' + entry.points + ')';
        var input = document.createElement('input');
        input.type = 'number';
        input.min = '0';
        input.max = String(entry.points);
        input.step = '0.25';
        input.dataset.manualScore = key;
        label.appendChild(input);
        item.appendChild(label);
      } else {
        var controls = document.createElement('div');
        controls.className = 'exam-review-choice';
        controls.innerHTML = `<label><input type="radio" name="decision-${index}-${key}" value="true" data-review-decision="${key}"> احتسابها صحيحة</label>
          <label><input type="radio" name="decision-${index}-${key}" value="false" data-review-decision="${key}"> اعتبارها خاطئة</label>`;
        item.appendChild(controls);
      }
      card.appendChild(item);
    });
    var button = document.createElement('button');
    button.className = 'btn btn-primary';
    button.textContent = 'اعتماد التصحيح وجمع الدرجة';
    button.onclick = function() { finalizeExamBookletReview(card); };
    card.appendChild(button);
    container.appendChild(card);
  });
}

function finalizeExamBookletReview(card) {
  var decisions = {};
  card.querySelectorAll('[data-review-decision]').forEach(input => {
    if (input.checked) decisions[input.dataset.reviewDecision] = input.value === 'true';
  });
  var manualScores = {};
  card.querySelectorAll('[data-manual-score]').forEach(input => {
    if (input.value !== '') manualScores[input.dataset.manualScore] = input.value;
  });
  var token = card.dataset.reviewToken;
  examApiJson('/api/exams/booklets/' + encodeURIComponent(token) + '/finalize', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body:JSON.stringify({decisions:decisions, manual_scores:manualScores})
  }).then(data => {
    showToast('اعتمدت الدرجة: ' + data.score);
    card.innerHTML = `<strong>اكتمل التصحيح. الدرجة النهائية: ${escapeHtml(data.score)}</strong>`;
    return loadCorrectionIssue(currentCorrectionIssue.id);
  }).catch(error => showToast('لم يعتمد التصحيح: ' + error.message));
}

function exportExamGrades() {
  if (!currentCorrectionIssue) return;
  fetch('/api/exams/issues/' + encodeURIComponent(currentCorrectionIssue.id) + '/grades.xlsx')
    .then(response => {
      if (!response.ok) return response.json().then(data => { throw new Error(data.message || 'تعذر إصدار ملف الدرجات'); });
      return response.blob();
    }).then(blob => {
      var link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'درجات_' + currentCorrectionIssue.subject + '.xlsx';
      link.click();
      URL.revokeObjectURL(link.href);
    }).catch(error => showToast('تعذر إصدار الدرجات: ' + error.message));
}

document.addEventListener('DOMContentLoaded', function() {
  populateDropdowns();
  initializeCohorts().then(() => loadDB()).catch(error => showToast('تعذر تحميل قائمة الأفواج: ' + error.message));
  loadLastIssuedExam();
  loadExamRoster();
  loadCorrectionIssues();
  var examPreviewModal = document.getElementById('exam-preview-modal');
  examPreviewModal.addEventListener('click', function(event) {
    if (event.target === examPreviewModal) closeExamBookletPreview();
  });
  window.addEventListener('keydown', function(event) {
    if (event.key === 'Escape' && examPreviewModal.classList.contains('open')) {
      closeExamBookletPreview();
    }
  });
  var localToday = new Date();
  localToday.setMinutes(localToday.getMinutes() - localToday.getTimezoneOffset());
  document.getElementById('issue-date').value = localToday.toISOString().slice(0, 10);
});
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/logo')
def college_logo():
  return send_file(LOGO_PATH, mimetype='image/png', conditional=True)

@app.route('/logo-watermark')
def college_logo_watermark():
  return send_file(WATERMARK_LOGO_PATH, mimetype='image/gif', conditional=True)

def start_flask():
    try:
        with open(RUNTIME_LOG_PATH, 'a', encoding='utf-8') as log_file:
            log_file.write(f'{datetime.now().isoformat()} Starting local Flask server.\n')
        app.run(host='127.0.0.1', port=54321, debug=False, use_reloader=False)
    except BaseException:
        with open(RUNTIME_LOG_PATH, 'a', encoding='utf-8') as log_file:
            traceback.print_exc(file=log_file)
        raise


def get_desktop_window_size():
    try:
        if os.name == 'nt':
            width = ctypes.windll.user32.GetSystemMetrics(0)
            height = ctypes.windll.user32.GetSystemMetrics(1)
            return width, height
    except Exception:
        pass
    return 1400, 900


if __name__ == '__main__':
    init_db()

    t = threading.Thread(target=start_flask)
    t.daemon = True
    t.start()

    api_bridge = ApiBridge()
    window_width, window_height = get_desktop_window_size()

    webview.create_window(
        title='برنامج امتحانات طلاب الكلية العسكرية',
        url='http://127.0.0.1:54321',
        js_api=api_bridge,
        width=window_width,
        height=window_height,
        resizable=True,
        maximized=True,
        fullscreen=False,
        min_size=(min(900, window_width), min(600, window_height)),
        background_color='#0f1623'
    )
    webview.start()
