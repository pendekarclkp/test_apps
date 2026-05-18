from flask import Flask, render_template, request, redirect, url_for, session, flash
import json, os
from datetime import datetime
import openpyxl

app = Flask(__name__)
app.secret_key = 'iirs_secret_key_2026'

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
DATA_FILE     = os.path.join(os.path.dirname(__file__), 'data', 'current_data.json')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)

USERS = {
    'admin':    {'password': 'admin',    'role': 'admin'},
    'uploader': {'password': 'upload123','role': 'uploader'},
}

PLAYBOOKS = {
    'Playbook 1': {
        'title': 'Playbook 1 — Kategori Ringan',
        'range': 'IIRS 0–35', 'color': 'success', 'icon': '✅',
        'actions': [
            'Monitor rutin sentimen publik',
            'Respons standar melalui kanal komunikasi resmi',
            'Update data berkala (mingguan)',
            'Tidak diperlukan eskalasi khusus',
        ],
    },
    'Playbook 2': {
        'title': 'Playbook 2 — Kategori Sedang',
        'range': 'IIRS 36–49', 'color': 'warning', 'icon': '⚠️',
        'actions': [
            'Peningkatan frekuensi monitoring sentimen',
            'Koordinasi internal tim komunikasi',
            'Siapkan narasi klarifikasi & FAQ publik',
            'Evaluasi mingguan dengan pimpinan',
        ],
    },
    'Crisis Watch': {
        'title': 'Crisis Watch — Zona Pre-Krisis',
        'range': 'IIRS 50–59 (Klaster Strategis)', 'color': 'orange', 'icon': '🔶',
        'actions': [
            'Aktivasi tim respons krisis',
            'Monitoring real-time 24/7',
            'Siapkan pernyataan resmi & press release',
            'Koordinasi dengan stakeholder eksternal',
            'Eskalasi ke manajemen senior',
        ],
    },
    'Playbook 3': {
        'title': 'Playbook 3 — Kategori Crisis',
        'range': 'IIRS ≥ 60', 'color': 'danger', 'icon': '🚨',
        'actions': [
            'Aktivasi penuh protokol krisis komunikasi',
            'War room & command center aktif',
            'Pernyataan resmi pimpinan tertinggi',
            'Koordinasi lintas kementerian/lembaga',
            'Media briefing & press conference segera',
            'Laporan harian kepada Gubernur BI',
        ],
    },
}

DEFAULT_DATA = {
    'sentrix': [
        {'cluster': 'Rupiah',  'ns': 0.326, 'va': 0.101, 'er': 1.000, 'score': 46.07},
        {'cluster': 'Inflasi', 'ns': 0.140, 'va': 0.530, 'er': 0.840, 'score': 46.70},
        {'cluster': 'BI-Rate', 'ns': 0.250, 'va': 0.170, 'er': 1.000, 'score': 45.10},
    ],
    'nolimit': [
        {'cluster': 'Rupiah',  'ns': 0.350, 'va': 0.170, 'er': 1.000, 'score': 49.10},
        {'cluster': 'Inflasi', 'ns': 0.140, 'va': 0.000, 'er': 0.250, 'score': 13.10},
        {'cluster': 'BI-Rate', 'ns': 0.216, 'va': 0.304, 'er': 0.271, 'score': 25.89},
    ],
    'iirs': [
        {'cluster': 'Rupiah',  'score_nolimit': 49.10, 'score_sentrix': 46.07, 'iirs': 47.282},
        {'cluster': 'Inflasi', 'score_nolimit': 13.10, 'score_sentrix': 46.70, 'iirs': 34.940},
        {'cluster': 'BI-Rate', 'score_nolimit': 25.89, 'score_sentrix': 45.10, 'iirs': 38.3765},
    ],
    'uploaded_at': None,
    'uploaded_by': None,
    'filename': None,
}

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE) as f:
            return json.load(f)
    return DEFAULT_DATA.copy()

def save_data(data):
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def parse_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # ── Source Scoring sheet ──────────────────────────────────────────────────
    ws = wb['Source Scoring']
    rows = list(ws.iter_rows(values_only=True))

    sentrix_start = nolimit_start = None
    for i, row in enumerate(rows):
        val = str(row[0]).strip().lower() if row[0] else ''
        if 'sentrix' in val:
            sentrix_start = i
        if 'no limit' in val or 'nolimit' in val:
            nolimit_start = i

    def extract_scores(start_idx):
        result = []
        if start_idx is None:
            return result
        for row in rows[start_idx + 1:]:
            if not row[0] or str(row[0]).strip() in ('', 'Issue Cluster'):
                if result:
                    break
                continue
            if row[4] is None:
                continue
            try:
                result.append({
                    'cluster': str(row[0]).strip(),
                    'ns':    round(float(row[1] or 0), 4),
                    'va':    round(float(row[2] or 0), 4),
                    'er':    round(float(row[3] or 0), 4),
                    'score': round(float(row[4]), 4),
                })
            except (TypeError, ValueError):
                continue
        return result

    sentrix = extract_scores(sentrix_start)
    nolimit = extract_scores(nolimit_start)

    # ── IIRS Integration sheet ────────────────────────────────────────────────
    ws2  = wb['IIRS Integration']
    rows2 = list(ws2.iter_rows(values_only=True))

    header_idx = None
    for i, row in enumerate(rows2):
        if row[0] and str(row[0]).strip() == 'Issue Cluster':
            header_idx = i
            break

    iirs = []
    if header_idx is not None:
        for row in rows2[header_idx + 1:]:
            if not row[0] or row[3] is None:
                if iirs:
                    break
                continue
            try:
                iirs.append({
                    'cluster':       str(row[0]).strip(),
                    'score_nolimit': round(float(row[1] or 0), 4),
                    'score_sentrix': round(float(row[2] or 0), 4),
                    'iirs':          round(float(row[3]), 4),
                })
            except (TypeError, ValueError):
                continue

    if not sentrix or not nolimit or not iirs:
        raise ValueError('Format file Excel tidak sesuai. Pastikan sheet "Source Scoring" dan "IIRS Integration" ada dan terisi.')

    return sentrix, nolimit, iirs

def get_kategori(iirs_val):
    if iirs_val < 35:
        return 'Ringan',      'Playbook 1',  'success'
    elif iirs_val < 50:
        return 'Sedang',      'Playbook 2',  'warning'
    elif iirs_val < 60:
        return 'Crisis Watch','Crisis Watch', 'orange'
    else:
        return 'Crisis',      'Playbook 3',  'danger'

def require_login(role=None):
    if 'user' not in session:
        return redirect(url_for('login'))
    if role and USERS.get(session['user'], {}).get('role') != role:
        flash('Akses ditolak.', 'danger')
        return redirect(url_for('login'))
    return None

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
def login():
    if 'user' in session:
        role = USERS.get(session['user'], {}).get('role')
        return redirect(url_for('upload_page') if role == 'uploader' else url_for('dashboard'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        user = USERS.get(username)
        if user and user['password'] == password:
            session['user'] = username
            session['role'] = user['role']
            if user['role'] == 'uploader':
                return redirect(url_for('upload_page'))
            return redirect(url_for('dashboard'))
        flash('Username atau password salah.', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
def dashboard():
    err = require_login(role='admin')
    if err:
        return err

    data = load_data()
    sentrix = data['sentrix']
    nolimit = data['nolimit']

    iirs_enriched = []
    for row in data['iirs']:
        kategori, playbook_key, color = get_kategori(row['iirs'])
        iirs_enriched.append({
            **row,
            'kategori':    kategori,
            'playbook_key': playbook_key,
            'color':       color,
            'playbook':    PLAYBOOKS.get(playbook_key, {}),
        })

    return render_template(
        'dashboard.html',
        user=session['user'],
        role=session.get('role'),
        sentrix=sentrix,
        nolimit=nolimit,
        iirs_data=iirs_enriched,
        playbooks=PLAYBOOKS,
        uploaded_at=data.get('uploaded_at'),
        uploaded_by=data.get('uploaded_by'),
        filename=data.get('filename'),
        chart_labels =json.dumps([r['cluster'] for r in data['iirs']]),
        chart_sentrix=json.dumps([r['score_sentrix'] for r in data['iirs']]),
        chart_nolimit=json.dumps([r['score_nolimit'] for r in data['iirs']]),
        chart_iirs   =json.dumps([round(r['iirs'], 2) for r in data['iirs']]),
    )

@app.route('/upload', methods=['GET', 'POST'])
def upload_page():
    err = require_login()
    if err:
        return err

    data     = load_data()
    preview  = None
    parse_err = None

    if request.method == 'POST':
        f = request.files.get('excel_file')
        if not f or not f.filename.endswith(('.xlsx', '.xls')):
            flash('Harap upload file Excel (.xlsx atau .xls).', 'danger')
        else:
            save_path = os.path.join(UPLOAD_FOLDER, 'latest.xlsx')
            f.save(save_path)
            try:
                sentrix, nolimit, iirs = parse_excel(save_path)
                new_data = {
                    'sentrix':     sentrix,
                    'nolimit':     nolimit,
                    'iirs':        iirs,
                    'uploaded_at': datetime.now().strftime('%d %b %Y, %H:%M WIB'),
                    'uploaded_by': session['user'],
                    'filename':    f.filename,
                }
                save_data(new_data)
                data    = new_data
                preview = {'sentrix': sentrix, 'nolimit': nolimit, 'iirs': iirs}
                flash(f'Data berhasil diupload dari file "{f.filename}".', 'success')
            except Exception as e:
                parse_err = str(e)
                flash(f'Gagal membaca file: {parse_err}', 'danger')

    return render_template(
        'upload.html',
        user=session['user'],
        role=session.get('role'),
        data=data,
        preview=preview,
        parse_err=parse_err,
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
