import base64
import io
import os
import tempfile
import unittest

import app


class TestExamWorkflow(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_paths = {
            'DB_PATH': app.DB_PATH,
            'BACKUP_DIR': app.BACKUP_DIR,
            'BACKUP_PATH': app.BACKUP_PATH,
        }
        app.DB_PATH = os.path.join(self.temp_dir.name, 'exam_bank.db')
        app.BACKUP_DIR = os.path.join(self.temp_dir.name, 'backups')
        app.BACKUP_PATH = os.path.join(app.BACKUP_DIR, 'backup_exam_bank.db')
        app.init_db()
        self.client = app.app.test_client()

    def scan_booklet(self, booklet, selected_option):
        qr_bytes = base64.b64decode(booklet['qr_data'].partition(',')[2])
        document = app.fitz.open()
        page = document.new_page(width=595, height=842)
        page.insert_image(app.fitz.Rect(15, 15, 100, 100), stream=qr_bytes)
        for option in range(2):
            page.draw_circle(
                app.fitz.Point(520, 280 + option * 20),
                radius=5,
                color=(0, 0, 0),
                fill=(0, 0, 0) if option == selected_option else None,
                width=1
            )
        pdf_bytes = document.tobytes()
        document.close()
        rendered_document = app.fitz.open(stream=pdf_bytes, filetype='pdf')
        pixmap = rendered_document.load_page(0).get_pixmap(matrix=app.fitz.Matrix(3, 3), alpha=False)
        rendered = app.np.frombuffer(pixmap.samples, dtype=app.np.uint8).reshape(
            pixmap.height, pixmap.width, pixmap.n
        )
        decoded, _, _ = app.cv2.QRCodeDetector().detectAndDecode(rendered)
        rendered_document.close()
        self.assertTrue(decoded, f'QR should remain readable after rendering to PDF: {pixmap.width}x{pixmap.height}')
        return self.client.post(
            '/api/exams/scan',
            data={'file': (io.BytesIO(pdf_bytes), 'scan.pdf')},
            content_type='multipart/form-data'
        )

    def tearDown(self):
        for name, value in self.original_paths.items():
            setattr(app, name, value)
        self.temp_dir.cleanup()

    def test_company_number_accepts_arabic_and_english_forms(self):
        self.assertEqual(app.normalize_company_number('السرية ٢'), 2)
        self.assertEqual(app.normalize_company_number('Company 4'), 4)
        self.assertIsNone(app.normalize_company_number('سرية خامسة'))

    def test_answer_indexes_follow_current_model_option_order(self):
        question = {'type': 'mc', 'options': ['الخيار أ', 'الخيار ب', 'الخيار ج'], 'answer': 'ب'}
        self.assertEqual(app.exam_answer_indexes(question), [1])
        shuffled = {'type': 'mc', 'options': ['الخيار ج', 'الخيار أ', 'الخيار ب'], 'answer': 'الخيار ب'}
        self.assertEqual(app.exam_answer_indexes(shuffled), [2])

    def test_true_false_answers_have_scan_choices(self):
        self.assertEqual(app.exam_answer_indexes({'type': 'tf', 'answer': 'صح'}), [0])
        self.assertEqual(app.exam_answer_indexes({'type': 'tf', 'answer': 'خطأ'}), [1])

    def test_multi_select_answer_returns_all_matching_indexes(self):
        question = {
            'type': 'multi_select',
            'options': ['الأول', 'الثاني', 'الثالث'],
            'answer': 'أ، ج'
        }
        self.assertEqual(app.exam_answer_indexes(question), [0, 2])

    def test_exam_question_keys_preserve_block_and_question_order(self):
        questions = app.exam_questions([
            {'title': 'السؤال الأول', 'items': [{'text': '1'}, {'text': '2'}]},
            {'title': 'السؤال الثاني', 'items': [{'text': '3'}]}
        ])
        self.assertEqual([question['key'] for question in questions], ['0:0', '0:1', '1:0'])

    def test_application_home_and_questions_api(self):
        home = self.client.get('/')
        self.assertEqual(home.status_code, 200)
        self.assertIn(b'exam-preview-modal', home.data)
        self.assertIn('معاينة الدفاتر قبل الطباعة'.encode('utf-8'), home.data)
        questions = self.client.get('/api/questions')
        self.assertEqual(questions.status_code, 200)
        self.assertEqual(questions.get_json(), [])

    def test_roster_import_finds_headers_after_title_on_non_active_sheet(self):
        workbook = app.openpyxl.Workbook()
        workbook.active.title = 'ملاحظات'
        workbook.active['A1'] = 'قائمة الطلبة للدورة'
        roster = workbook.create_sheet('بيانات الطلاب')
        roster['A1'] = 'قائمة الأسماء'
        roster['A3'] = 'التسلسل'
        roster['B3'] = 'الاسم الكامل'
        roster['C3'] = 'رقم السرية'
        roster['D3'] = 'رقم الفصيل'
        roster['A4'] = 1234
        roster['B4'] = 'طالب تجريبي'
        roster['C4'] = 'السرية الأولى'
        roster['D4'] = 2
        file_data = io.BytesIO()
        workbook.save(file_data)
        file_data.seek(0)

        response = self.client.post(
            '/api/exams/roster',
            data={'file': (file_data, 'قائمة الطلاب.xlsx')},
            content_type='multipart/form-data'
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        self.assertEqual(response.get_json()['total'], 1)
        self.assertEqual(response.get_json()['counts'], {'1': 1})

        imported = app.get_db()
        student = imported.execute('SELECT * FROM exam_roster').fetchone()
        imported.close()
        self.assertEqual(student['military_number'], 'C1-S1234')
        self.assertEqual(student['student_sequence'], 1234)
        self.assertEqual(student['student_name'], 'طالب تجريبي')
        self.assertEqual(student['company'], 1)
        self.assertEqual(student['platoon'], '2')

        invalid_workbook = app.openpyxl.Workbook()
        invalid_workbook.active.append(['عمود غير مدعوم'])
        invalid_data = io.BytesIO()
        invalid_workbook.save(invalid_data)
        invalid_data.seek(0)
        rejected = self.client.post(
            '/api/exams/roster',
            data={'file': (invalid_data, 'ملف خاطئ.xlsx')},
            content_type='multipart/form-data'
        )
        self.assertEqual(rejected.status_code, 400, rejected.get_json())
        remaining = app.get_db()
        self.assertEqual(remaining.execute('SELECT COUNT(*) FROM exam_roster').fetchone()[0], 1)
        remaining.close()

    def test_issue_grading_and_excel_export(self):
        conn = app.get_db()
        conn.executemany(
            '''INSERT INTO exam_roster
               (military_number, student_name, company, platoon, student_sequence)
               VALUES (?, ?, ?, ?, ?)''',
            [
                ('C1-S0001', 'طالب أول', 1, '1', 1),
                ('C1-S0002', 'طالب ثان', 1, '1', 2)
            ]
        )
        conn.commit()
        conn.close()

        model_a = [{'title': 'السؤال الأول', 'items': [{
            'type': 'mc', 'text': 'اختر الإجابة الصحيحة', 'options': ['صحيح', 'خطأ'], 'answer': 'أ'
        }]}]
        model_b = [{'title': 'السؤال الأول', 'items': [{
            'type': 'mc', 'text': 'اختر الإجابة الصحيحة', 'options': ['خطأ', 'صحيح'], 'answer': 'ب'
        }]}]
        issue_payload = {
            'course': 'دورة تجريبية',
            'cohort': 'A',
            'subject': 'الحاسبات الالكترونية',
            'year': 1,
            'semester': 1,
            'exam_date': '2026-10-15',
            'exam_time': '09:00',
            'counts': {'1': 2, '2': 0, '3': 0, '4': 0},
            'modelA': model_a,
            'modelB': model_b,
        }
        over_limit_payload = dict(issue_payload)
        over_limit_payload['counts'] = {'1': 3, '2': 0, '3': 0, '4': 0}
        over_limit = self.client.post('/api/exams/issues', json=over_limit_payload)
        self.assertEqual(over_limit.status_code, 400)
        self.assertIn('أكبر من عدد الطلاب', over_limit.get_json()['message'])

        issue_response = self.client.post('/api/exams/issues', json=issue_payload)
        self.assertEqual(issue_response.status_code, 200, issue_response.get_json())
        issue_id = issue_response.get_json()['issue_id']
        issued_booklets = issue_response.get_json()['booklets']
        first_booklet_qr = issued_booklets[0]['qr_data']
        qr_bytes = base64.b64decode(first_booklet_qr.partition(',')[2])
        qr_image = app.cv2.imdecode(app.np.frombuffer(qr_bytes, dtype=app.np.uint8), app.cv2.IMREAD_COLOR)
        qr_text, _, _ = app.cv2.QRCodeDetector().detectAndDecode(qr_image)
        self.assertTrue(qr_text, 'Generated booklet QR should be readable')
        conn = app.get_db()
        signing_secret = app.qr_signing_secret(conn)
        conn.close()
        payload = app.decode_exam_qr(qr_text, signing_secret)
        self.assertEqual((payload['c'], payload['s'], payload['m']), (1, 1, 'A'))

        issued = self.client.get(f'/api/exams/issues/{issue_id}').get_json()['issue']
        self.assertEqual([(row['serial_number'], row['model']) for row in issued['booklets']], [(1, 'A'), (2, 'B')])

        transfer_response = self.client.post(
            f'/api/exams/issues/{issue_id}/transfer-key',
            json={'printed': True, 'transfer': True}
        )
        self.assertEqual(transfer_response.status_code, 200, transfer_response.get_json())

        grading = {'A': [{'points': 10, 'manual': []}], 'B': [{'points': 10, 'manual': []}]}
        grading_response = self.client.post(
            f'/api/exams/issues/{issue_id}/grading', json={'grading': grading}
        )
        self.assertEqual(grading_response.status_code, 200, grading_response.get_json())

        first_scan = self.scan_booklet(issued_booklets[0], selected_option=0)
        self.assertEqual(first_scan.status_code, 200, first_scan.get_json())
        self.assertEqual(len(first_scan.get_json()['booklets']), 1, first_scan.get_json())
        self.assertEqual(first_scan.get_json()['booklets'][0]['review'], [])
        first_token = issued_booklets[0]['qr_token']
        finalized = self.client.post(
            f'/api/exams/booklets/{first_token}/finalize',
            json={'decisions': {}, 'manual_scores': {}}
        )
        self.assertEqual(finalized.status_code, 200, finalized.get_json())
        self.assertEqual(finalized.get_json()['score'], 10)

        export = self.client.get(f'/api/exams/issues/{issue_id}/grades.xlsx')
        self.assertEqual(export.status_code, 409, export.get_json())

        second_scan = self.scan_booklet(issued_booklets[1], selected_option=1)
        self.assertEqual(second_scan.status_code, 200, second_scan.get_json())
        self.assertEqual(len(second_scan.get_json()['booklets']), 1, second_scan.get_json())
        self.assertEqual(second_scan.get_json()['booklets'][0]['review'], [])
        second_token = issued_booklets[1]['qr_token']
        finalized_second = self.client.post(
            f'/api/exams/booklets/{second_token}/finalize',
            json={'decisions': {}, 'manual_scores': {}}
        )
        self.assertEqual(finalized_second.status_code, 200, finalized_second.get_json())

        export = self.client.get(f'/api/exams/issues/{issue_id}/grades.xlsx')
        self.assertEqual(export.status_code, 200)
        workbook = app.openpyxl.load_workbook(io.BytesIO(export.data), data_only=True)
        sheet = workbook['الدرجات']
        self.assertEqual(sheet.max_row, 3)
        self.assertEqual(sheet.cell(row=2, column=9).value, 10)
        response_details = app.json.loads(sheet.cell(row=2, column=10).value)
        self.assertTrue(response_details['0:0']['correct'])


if __name__ == '__main__':
    unittest.main()
