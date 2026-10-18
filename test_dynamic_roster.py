import io
import os
import tempfile
import unittest

import app


class TestDynamicRosterImport(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = app.DB_PATH
        self.original_qr_path = app.QR_DATABASE_PATH
        app.DB_PATH = os.path.join(self.temp_dir.name, 'exam_bank.db')
        app.QR_DATABASE_PATH = os.path.join(self.temp_dir.name, 'qr_tokens.db')
        app.init_db()
        self.client = app.app.test_client()

    def tearDown(self):
        app.DB_PATH = self.original_db_path
        app.QR_DATABASE_PATH = self.original_qr_path
        self.temp_dir.cleanup()

    def workbook(self, sequence, name, company):
        workbook = app.openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(['التسلسل', 'الاسم الكامل', 'السرية', 'الفصيل'])
        sheet.append([sequence, name, company, 'الأول'])
        data = io.BytesIO()
        workbook.save(data)
        data.seek(0)
        return data

    def test_multiple_dynamic_files_replace_roster_and_create_tokens(self):
        response = self.client.post(
            '/api/exams/roster',
            data={'files': [
                (self.workbook(1, 'طالب أول', 'سرية أولى'), 'قائمة 1.xlsx'),
                (self.workbook(1, 'طالب ثان', 'سرية ثانية'), 'قائمة 2.xlsx')
            ]},
            content_type='multipart/form-data'
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        self.assertEqual(response.get_json()['total'], 2)
        conn = app.get_db()
        rows = conn.execute(
            '''SELECT military_number, student_sequence, qr_jwt
               FROM exam_roster ORDER BY company, student_sequence'''
        ).fetchall()
        conn.close()
        self.assertEqual(len(rows), 2)
        self.assertEqual(
            [(row['military_number'], row['student_sequence']) for row in rows],
            [('C1-S0001', 1), ('C2-S0001', 1)]
        )
        self.assertTrue(all(row['qr_jwt'].count('.') == 2 for row in rows))


if __name__ == '__main__':
    unittest.main()
