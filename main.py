import os
import json
import random
import sqlite3
import time
import shutil
import threading
from datetime import datetime
from flask import Flask, render_template, request, jsonify
import webview

# Initialize Flask application
app = Flask(__name__, template_folder='.', static_folder='.')

# Set up storage directory and database paths
APP_DATA_DIR = os.path.join(os.path.expanduser('~'), '.military_exam_app')
if not os.path.exists(APP_DATA_DIR):
    os.makedirs(APP_DATA_DIR)

DB_PATH = os.path.join(APP_DATA_DIR, 'exam_bank.db')
BACKUP_DIR = os.path.join(APP_DATA_DIR, 'backups')
if not os.path.exists(BACKUP_DIR):
    os.makedirs(BACKUP_DIR)


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def create_backup():
    """Creates a timestamped backup copy of the SQLite database."""
    try:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_filename = f"exam_bank_backup_{timestamp}.db"
        backup_path = os.path.join(BACKUP_DIR, backup_filename)
        if os.path.exists(DB_PATH):
            shutil.copy2(DB_PATH, backup_path)
            # Retain only the last 10 backup files
            backups = sorted([os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR)])
            if len(backups) > 10:
                for old_b in backups[:-10]:
                    os.remove(old_b)
    except Exception as e:
        print(f"Error creating auto-backup: {e}")


def init_db():
    """Initializes schema for questions and archived exams."""
    conn = get_db()
    cursor = conn.cursor()
    
    # Core questions table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY,
            subject TEXT NOT NULL,
            topic TEXT,
            year INTEGER,
            semester INTEGER,
            type TEXT,
            text TEXT NOT NULL,
            options TEXT,
            answer TEXT,
            used INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Exam history/archive table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS exam_archive (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            subject TEXT NOT NULL,
            generated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            model_a_json TEXT NOT NULL,
            model_b_json TEXT NOT NULL
        )
    ''')
    
    conn.commit()
    conn.close()


# ==========================================
# Web Routes & API (Replaces Electron IPC)
# ==========================================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/db-all', methods=['GET'])
def db_all():
    """Fetch all questions with optional search and filter parameters."""
    subject = request.args.get('subject', '')
    q_type = request.args.get('type', '')
    search = request.args.get('search', '')
    
    conn = get_db()
    cursor = conn.cursor()
    
    query = "SELECT * FROM questions WHERE 1=1"
    params = []
    
    if subject:
        query += " AND subject = ?"
        params.append(subject)
    if q_type:
        query += " AND type = ?"
        params.append(q_type)
    if search:
        query += " AND text LIKE ?"
        params.append(f"%{search}%")
        
    query += " ORDER BY id DESC"
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    
    questions = []
    for r in rows:
        q = dict(r)
        q['used'] = bool(q['used'])
        questions.append(q)
    return jsonify(questions)


@app.route('/api/db-add', methods=['POST'])
def db_add():
    """Insert a new question."""
    q = request.json
    q_id = int(time.time() * 1000) + random.randint(0, 999)
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO questions (id, subject, topic, year, semester, type, text, options, answer, used)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        q_id,
        q.get('subject', ''),
        q.get('topic', ''),
        q.get('year', 0),
        q.get('semester', 0),
        q.get('type', ''),
        q.get('text', ''),
        json.dumps(q.get('options', [])) if isinstance(q.get('options'), list) else str(q.get('options', '')),
        str(q.get('answer', '')),
        0
    ))
    conn.commit()
    conn.close()
    
    create_backup()
    return jsonify({"id": q_id})


@app.route('/api/db-update', methods=['POST'])
def db_update():
    """Update an existing question."""
    q = request.json
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE questions 
        SET subject=?, topic=?, year=?, semester=?, type=?, text=?, options=?, answer=?, used=?
        WHERE id=?
    ''', (
        q.get('subject'),
        q.get('topic'),
        q.get('year'),
        q.get('semester'),
        q.get('type'),
        q.get('text'),
        json.dumps(q.get('options', [])) if isinstance(q.get('options'), list) else str(q.get('options', '')),
        str(q.get('answer', '')),
        1 if q.get('used') else 0,
        q.get('id')
    ))
    conn.commit()
    conn.close()
    return jsonify(True)


@app.route('/api/db-delete', methods=['POST'])
def db_delete():
    """Delete a specific question."""
    data = request.json
    q_id = data.get('id')
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM questions WHERE id = ?", (q_id,))
    conn.commit()
    conn.close()
    return jsonify(True)


@app.route('/api/db-clear', methods=['POST'])
def db_clear():
    """Purge all questions."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM questions")
    conn.commit()
    conn.close()
    create_backup()
    return jsonify(True)


@app.route('/api/db-import', methods=['POST'])
def db_import():
    """Bulk import questions from JSON or external sources."""
    questions = request.json
    conn = get_db()
    cursor = conn.cursor()
    
    for i, q in enumerate(questions):
        q_id = int(time.time() * 1000) + i + random.randint(0, 999)
        cursor.execute('''
            INSERT INTO questions (id, subject, topic, year, semester, type, text, options, answer, used)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            q_id,
            q.get('subject', ''),
            q.get('topic', ''),
            q.get('year', 0),
            q.get('semester', 0),
            q.get('type', ''),
            q.get('text', ''),
            json.dumps(q.get('options', [])) if isinstance(q.get('options'), list) else str(q.get('options', '')),
            str(q.get('answer', '')),
            0
        ))
    
    conn.commit()
    conn.close()
    create_backup()
    return jsonify(True)


@app.route('/api/archive-exam', methods=['POST'])
def archive_exam():
    """Save generated Model A and Model B exam layouts."""
    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO exam_archive (title, subject, model_a_json, model_b_json)
        VALUES (?, ?, ?, ?)
    ''', (
        data.get('title', 'امتحان عسكري'),
        data.get('subject', 'عام'),
        json.dumps(data.get('model_a', [])),
        json.dumps(data.get('model_b', []))
    ))
    conn.commit()
    conn.close()
    return jsonify(True)


@app.route('/api/db-path', methods=['GET'])
def db_path():    return jsonify({"dbPath": DB_PATH, "backupPath": BACKUP_DIR})


# ==========================================
# Application Launcher
# ==========================================

def start_flask():
    app.run(host='127.0.0.1', port=54321, debug=False, use_reloader=False)


if __name__ == '__main__':
    init_db()
    create_backup()
    
    # Launch Flask server thread
    t = threading.Thread(target=start_flask)
    t.daemon = True
    t.start()
    
    # Create Native Desktop App Window
    webview.create_window(
        title='تطبيق امتحانات الكلية العسكرية العراقية',
        url='http://127.0.0.1:54321',
        width=1280,
        height=800,
        min_size=(900, 600),
        background_color='#0d1b2a'
    )
    webview.start()