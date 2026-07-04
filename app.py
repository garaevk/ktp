# -*- coding: utf-8 -*-
"""
Flask web application for KTP Generator.
Multi-step: upload → preview table → configure dates → download.
"""

import os
import uuid
import json
import shutil
import threading
import time
from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from werkzeug.utils import secure_filename
from ktp_logic import *

app = Flask(__name__)
app.secret_key = 'ktp-generator-secret-key-2025-2026'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(
    os.environ.get('TEMP', os.path.join(BASE_DIR, 'temp')),
    'ktp_sessions'
)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Cleanup old sessions every hour
CLEANUP_AGE = 3600  # 1 hour


def cleanup_old_sessions():
    while True:
        time.sleep(CLEANUP_AGE)
        now = time.time()
        try:
            for entry in os.listdir(UPLOAD_FOLDER):
                path = os.path.join(UPLOAD_FOLDER, entry)
                if os.path.isdir(path) and now - os.path.getmtime(path) > CLEANUP_AGE:
                    shutil.rmtree(path, ignore_errors=True)
        except Exception:
            pass


threading.Thread(target=cleanup_old_sessions, daemon=True).start()


def get_session_dir(sid):
    path = os.path.join(UPLOAD_FOLDER, sid)
    os.makedirs(path, exist_ok=True)
    return path


def generate_sid():
    return uuid.uuid4().hex[:12]


# ===== ROUTES =====

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/', methods=['POST'])
def upload():
    if 'ktp_file' not in request.files:
        flash('Пожалуйста, выберите файл КТП (.docx)', 'danger')
        return render_template('index.html')

    ktp_file = request.files['ktp_file']
    if not ktp_file.filename:
        flash('Файл не выбран', 'danger')
        return render_template('index.html')

    if not ktp_file.filename.lower().endswith('.docx'):
        flash('Файл КТП должен быть в формате .docx', 'danger')
        return render_template('index.html')

    sid = generate_sid()
    session_dir = get_session_dir(sid)

    ktp_path = os.path.join(session_dir, 'uploaded.docx')
    ktp_file.save(ktp_path)

    calendar_path = None
    if 'calendar_file' in request.files and request.files['calendar_file'].filename:
        cal_file = request.files['calendar_file']
        if cal_file.filename.lower().endswith('.pdf'):
            calendar_path = os.path.join(session_dir, 'calendar.pdf')
            cal_file.save(calendar_path)

    meta = {
        'grade': request.form.get('grade', '').strip(),
        'subject': request.form.get('subject', '').strip(),
        'level': request.form.get('level', '').strip(),
    }
    with open(os.path.join(session_dir, 'meta.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False)

    return redirect(url_for('preview', sid=sid))


@app.route('/preview/<sid>')
def preview(sid):
    session_dir = get_session_dir(sid)
    ktp_path = os.path.join(session_dir, 'uploaded.docx')

    if not os.path.exists(ktp_path):
        flash('Сессия не найдена. Начните заново.', 'warning')
        return redirect(url_for('index'))

    try:
        with open(os.path.join(session_dir, 'meta.json'), 'r', encoding='utf-8') as f:
            meta = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        meta = {}

    tables = get_tables_info(ktp_path)
    if not tables:
        flash('В документе не найдено таблиц.', 'warning')
        return redirect(url_for('index'))

    auto_index = auto_detect_table(ktp_path)

    return render_template('preview.html', sid=sid, tables=tables,
                           meta=meta, auto_index=auto_index,
                           school_name=meta.get('school_name', ''))


@app.route('/preview/<sid>', methods=['POST'])
def select_table(sid):
    session_dir = get_session_dir(sid)
    table_index = int(request.form.get('table_index', 0))

    with open(os.path.join(session_dir, 'selection.json'), 'w', encoding='utf-8') as f:
        json.dump({'table_index': table_index}, f)

    return redirect(url_for('configure', sid=sid))


@app.route('/configure/<sid>')
def configure(sid):
    session_dir = get_session_dir(sid)
    ktp_path = os.path.join(session_dir, 'uploaded.docx')

    if not os.path.exists(ktp_path):
        flash('Сессия не найдена. Начните заново.', 'warning')
        return redirect(url_for('index'))

    try:
        with open(os.path.join(session_dir, 'meta.json'), 'r', encoding='utf-8') as f:
            meta = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        meta = {}

    try:
        with open(os.path.join(session_dir, 'selection.json'), 'r', encoding='utf-8') as f:
            sel = json.load(f)
        table_index = sel.get('table_index', 0)
    except (FileNotFoundError, json.JSONDecodeError):
        return redirect(url_for('preview', sid=sid))

    lessons = extract_lessons(ktp_path, table_index)
    if not lessons:
        flash('Не удалось извлечь уроки из таблицы. Попробуйте другую.', 'warning')
        return redirect(url_for('preview', sid=sid))

    # Cache lessons
    with open(os.path.join(session_dir, 'lessons.json'), 'w', encoding='utf-8') as f:
        json.dump(lessons, f, ensure_ascii=False)

    # Try to parse calendar
    calendar_path = os.path.join(session_dir, 'calendar.pdf')
    parsed = None
    if os.path.exists(calendar_path):
        parsed = parse_calendar_pdf(calendar_path)

    date_config = {
        'quarters': parsed['quarters'] if parsed and parsed.get('quarters') else DEFAULT_QUARTERS,
        'days_of_week': [0, 1, 2, 3, 5],  # Mon-Fri + Sat default
        'holiday_ranges': parsed['holiday_ranges'] if parsed and parsed.get('holiday_ranges') else DEFAULT_HOLIDAY_RANGES,
        'specific_holidays': parsed['specific_holidays'] if parsed and parsed.get('specific_holidays') else DEFAULT_SPECIFIC_HOLIDAYS,
    }

    weekdays_ru = WEEKDAYS_FULL

    return render_template('configure.html', sid=sid, meta=meta,
                           lessons=lessons, date_config=date_config,
                           weekdays=weekdays_ru,
                           school_name=meta.get('school_name', ''),
                           lessons_per_day={})


@app.route('/configure/<sid>', methods=['POST'])
def save_configure(sid):
    session_dir = get_session_dir(sid)
    ktp_path = os.path.join(session_dir, 'uploaded.docx')

    try:
        with open(os.path.join(session_dir, 'meta.json'), 'r', encoding='utf-8') as f:
            meta = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        meta = {}

    try:
        with open(os.path.join(session_dir, 'lessons.json'), 'r', encoding='utf-8') as f:
            lessons = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        flash('Данные уроков не найдены. Начните заново.', 'warning')
        return redirect(url_for('index'))

    # Parse form data
    quarters = []
    q_names = request.form.getlist('quarter_name[]')
    q_starts = request.form.getlist('quarter_start[]')
    q_ends = request.form.getlist('quarter_end[]')
    for i in range(len(q_names)):
        if q_starts[i] and q_ends[i]:
            quarters.append({
                'name': q_names[i],
                'start': q_starts[i],
                'end': q_ends[i],
            })

    days_of_week = [int(d) for d in request.form.getlist('days[]')]
    lessons_per_day_map = {}
    for i in range(6):
        try:
            val = int(request.form.get(f'lessons_per_day_{i}', 0))
        except (ValueError, TypeError):
            val = 0
        if val > 0:
            lessons_per_day_map[i] = val

    holiday_ranges = []
    hr_names = request.form.getlist('hr_name[]')
    hr_starts = request.form.getlist('hr_start[]')
    hr_ends = request.form.getlist('hr_end[]')
    for i in range(len(hr_names)):
        if hr_starts[i] and hr_ends[i]:
            holiday_ranges.append({
                'name': hr_names[i],
                'start': hr_starts[i],
                'end': hr_ends[i],
            })

    specific_holidays = []
    sh_dates = request.form.getlist('sh_date[]')
    sh_names = request.form.getlist('sh_name[]')
    for i in range(len(sh_dates)):
        if sh_dates[i]:
            specific_holidays.append({
                'date': sh_dates[i],
                'name': sh_names[i] if i < len(sh_names) else '',
            })

    if not quarters:
        flash('Добавьте хотя бы один учебный период.', 'danger')
        return redirect(url_for('configure', sid=sid))

    if not days_of_week:
        flash('Выберите хотя бы один день недели.', 'danger')
        return redirect(url_for('configure', sid=sid))

    dates = generate_dates(quarters, days_of_week, holiday_ranges,
                           specific_holidays, len(lessons),
                           lessons_per_day_map=lessons_per_day_map)

    total_with_dates = sum(1 for d in dates if d)
    total_empty = len(dates) - total_with_dates

    # Save config + dates for result page
    config = {
        'quarters': quarters,
        'days_of_week': days_of_week,
        'holiday_ranges': holiday_ranges,
        'specific_holidays': specific_holidays,
        'dates': dates,
        'meta': meta,
        'total_with_dates': total_with_dates,
        'total_empty': total_empty,
        'lessons_per_day': lessons_per_day_map,
    }
    with open(os.path.join(session_dir, 'config.json'), 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False)

    return redirect(url_for('result', sid=sid))


@app.route('/result/<sid>')
def result(sid):
    session_dir = get_session_dir(sid)

    try:
        with open(os.path.join(session_dir, 'config.json'), 'r', encoding='utf-8') as f:
            config = json.load(f)
        with open(os.path.join(session_dir, 'lessons.json'), 'r', encoding='utf-8') as f:
            lessons = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        flash('Данные не найдены. Начните заново.', 'warning')
        return redirect(url_for('index'))

    return render_template('result.html', sid=sid, lessons=lessons,
                           total=len(lessons),
                           school_name=config.get('meta', {}).get('school_name', ''),
                           **config)


@app.route('/download/<sid>')
def download(sid):
    session_dir = get_session_dir(sid)

    try:
        with open(os.path.join(session_dir, 'config.json'), 'r', encoding='utf-8') as f:
            config = json.load(f)
        with open(os.path.join(session_dir, 'lessons.json'), 'r', encoding='utf-8') as f:
            lessons = json.load(f)
        with open(os.path.join(session_dir, 'meta.json'), 'r', encoding='utf-8') as f:
            meta = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        flash('Данные не найдены. Начните заново.', 'warning')
        return redirect(url_for('index'))

    dates = config.get('dates', [])

    total_hours = sum(int(l['hours']) for l in lessons if l['hours'].isdigit())
    control_count = sum(1 for l in lessons if l['control'] and l['control'] != '0')
    practical_count = sum(1 for l in lessons if l['practical'] and l['practical'] != '0')

    doc = build_ktp(
        lessons, dates,
        grade=meta.get('grade', ''),
        subject=meta.get('subject', ''),
        level=meta.get('level', ''),
        school_name=meta.get('school_name', ''),
        control_count=control_count or 8,
        practical_count=practical_count or 0,
        total_hours=total_hours,
    )

    output_name = f"КТП_{meta.get('subject', 'Предмет')}_{meta.get('grade', '')}кл.docx"
    output_path = os.path.join(session_dir, 'result.docx')
    doc.save(output_path)

    return send_file(
        output_path,
        as_attachment=True,
        download_name=output_name,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@app.errorhandler(404)
def not_found(e):
    return render_template('error.html', message='Страница не найдена'), 404


@app.errorhandler(500)
def server_error(e):
    return render_template('error.html', message='Внутренняя ошибка сервера'), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') == 'development'
    print(f"КТП-генератор запущен на http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=debug)
