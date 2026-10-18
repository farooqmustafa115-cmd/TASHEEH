# -*- coding: utf-8 -*-
"""فحص شامل لمحرك التصحيح على النموذج الفعلي Scan_20261017 - Copy (5).pdf
يعمل على نسخة مؤقتة من قاعدة البيانات ولا يلمس بياناتك الحقيقية.
شغّله بأمر:  .venv\\Scripts\\python.exe _verify_scan.py
"""
import sys, os, io, json, shutil, tempfile, unittest

WS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, WS)
os.chdir(WS)

print('=' * 70)
print('1) استيراد التطبيق والمحركات')
print('=' * 70)
import app
print('IMPORT OK')
print('HAS_SCAN_LIBS =', app.HAS_SCAN_LIBS, '| HAS_QRCODE =', app.HAS_QRCODE,
      '| HAS_TESSERACT =', app.HAS_TESSERACT, '| HAS_PYZBAR =', app.HAS_PYZBAR,
      '| HAS_DOCX =', app.HAS_DOCX, '| HAS_OPENPYXL =', app.HAS_OPENPYXL)
print('tesseract:', app.tesseract_cmd())

print()
print('=' * 70)
print('2) اختبارات الوحدات (test_exam_workflow)')
print('=' * 70)
suite = unittest.TestLoader().loadTestsFromName('test_exam_workflow')
result = unittest.TextTestRunner(verbosity=1).run(suite)
print('UNIT TESTS:', 'OK - كلها ناجحة' if result.wasSuccessful() else 'FAILED')

print()
print('=' * 70)
print('3) عزل نسخة مؤقتة من قاعدة البيانات وفحص جاهزية الإصدارات')
print('=' * 70)
tmp = tempfile.mkdtemp(prefix='exam_verify_')
app.DB_PATH = os.path.join(tmp, 'exam_bank.db')
shutil.copyfile(os.path.join(WS, 'exam_bank.db'), app.DB_PATH)
app.BACKUP_DIR = os.path.join(tmp, 'backups')
app.BACKUP_PATH = os.path.join(app.BACKUP_DIR, 'backup_exam_bank.db')
print('نسخة مؤقتة:', app.DB_PATH)

conn = app.get_db()
issues = conn.execute('SELECT id, subject, cohort, key_transferred, grading_json, model_a_json FROM exam_issues').fetchall()
for r in issues:
    blocks = json.loads(r['model_a_json'])
    grading = json.loads(r['grading_json'] or '{}')
    ready = r['key_transferred'] and all(len(grading.get(m, [])) == len(blocks) for m in ('A', 'B'))
    print('إصدار:', r['subject'], '| فوج', r['cohort'], '| مفتاح معتمد =', bool(r['key_transferred']),
          '| مجموعات =', len(blocks), '| جاهز للمسح =', bool(ready))
booklets = conn.execute('SELECT COUNT(*) c, SUM(status = "corrected") done FROM exam_booklets').fetchone()
print('عدد الدفاتر:', booklets['c'], '| مصححة مسبقاً:', booklets['done'] or 0)
conn.close()

print()
print('=' * 70)
print('4) تحليل صفحات النموذج الفعلي صفحةً صفحة')
print('=' * 70)
PDF = os.path.join(WS, 'Scan_20261017 - Copy (5).pdf')
if not os.path.exists(PDF):
    candidates = [f for f in os.listdir(WS) if f.lower().endswith('.pdf')]
    PDF = os.path.join(WS, candidates[0]) if candidates else None
print('الملف:', PDF)
data = open(PDF, 'rb').read()
print('الحجم:', round(len(data) / 1024 / 1024, 2), 'MB')

detector = app.cv2.QRCodeDetector()
conn = app.get_db()
secret = app.qr_signing_secret(conn)
conn.commit()
conn.close()
page_count = 0
for page_no, image in app.iter_scan_pages('scan.pdf', data):
    page_count += 1
    code = app.read_qr_from_image(image, detector)
    payload = app.decode_exam_qr(code, secret) if code else None
    bubbles = app.detect_page_bubbles(image)
    marked = sum(1 for b in bubbles if b['marked'])
    uncertain = sum(1 for b in bubbles if b['uncertain'])
    print(f'صفحة {page_no}: {image.shape[1]}x{image.shape[0]} | QR={"مقروء" if payload else "غير مقروء"}'
          f' | فقاعات={len(bubbles)} | مظللة={marked} | ملتبسة={uncertain}')
    if payload:
        print('   QR payload:', payload)
print('مجموع الصفحات:', page_count)

print()
print('=' * 70)
print('5) المسح الكامل عبر /api/exams/scan (على النسخة المؤقتة)')
print('=' * 70)
client = app.app.test_client()
resp = client.post('/api/exams/scan',
                   data={'files': [(io.BytesIO(data), os.path.basename(PDF))]},
                   content_type='multipart/form-data')
print('HTTP', resp.status_code)
out = resp.get_json()
print('المحركات:', out.get('engines'))
print('دفاتر محللة:', len(out.get('booklets', [])))
conn = app.get_db()
for b in out.get('booklets', []):
    row = conn.execute('SELECT responses_json FROM exam_booklets WHERE qr_token = ?', (b['qr_token'],)).fetchone()
    responses = json.loads(row[0] or '{}') if row else {}
    correct = sum(1 for v in responses.values() if v.get('correct'))
    print(f"  دفتر تسلسل {b['serial_number']} | سرية {b['company']} | نموذج {b['model']}"
          f" | تعرف عبر {b['identified_by']} | إجابات صحيحة آلياً: {correct}/{len(responses)}"
          f" | مواضع مراجعة: {len(b['review'])}")
    for entry in b['review'][:8]:
        print('     مراجعة:', entry['question_key'], '-', entry['reason'])
conn.close()
for p in out.get('page_problems', []):
    print('  مشكلة صفحة', p['page'], ':', p['message'])

print()
print('انتهى الفحص. قاعدة بياناتك الأصلية لم تُمس.')
