import json
import app

print('APP_IMPORT_OK')
print('HAS_DOCX=', app.HAS_DOCX)
print('APP_NAME=', app.app.name)
print('ROUTES=', sorted(rule.rule for rule in app.app.url_map.iter_rules()))

client = app.app.test_client()
home = client.get('/')
questions = client.get('/api/questions')

print('HOME_STATUS=', home.status_code)
print('QUESTIONS_STATUS=', questions.status_code)
print('QUESTIONS_LEN=', len(questions.get_json()))
