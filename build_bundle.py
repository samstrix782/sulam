"""בונה את index.html מתוך app/: התבנית, השאלות ותיוג המפלגות.

    python3 build_bundle.py
"""
import json, os

ROOT = os.path.dirname(os.path.abspath(__file__))


def load(name):
    with open(os.path.join(ROOT, 'app', name), encoding='utf-8') as f:
        return json.load(f)


with open(os.path.join(ROOT, 'app', 'index.html'), encoding='utf-8') as f:
    template = f.read()
assert template.count('/*DATA*/') == 1, 'app/index.html must contain exactly one /*DATA*/ placeholder'

data = {'questions': load('questions.json'), 'parties': load('parties.json')}
out = template.replace('/*DATA*/', json.dumps(data, ensure_ascii=False))
with open(os.path.join(ROOT, 'index.html'), 'w', encoding='utf-8') as f:
    f.write(out)
print(f'index.html: {len(out.encode()) // 1024}KB')
