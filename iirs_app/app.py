from flask import Flask, render_template, request, redirect, url_for, session, flash
import json

app = Flask(__name__)
app.secret_key = 'iirs_secret_key_2026'

USERS = {
    'admin': 'admin'
}

# Data dari Excel
SOURCE_SCORE_SENTRIX = [
    {'cluster': 'Rupiah',  'ns': 0.326, 'va': 0.101, 'er': 1.000, 'score': 46.07},
    {'cluster': 'Inflasi', 'ns': 0.140, 'va': 0.530, 'er': 0.840, 'score': 46.70},
    {'cluster': 'BI-Rate', 'ns': 0.250, 'va': 0.170, 'er': 1.000, 'score': 45.10},
]

SOURCE_SCORE_NOLIMIT = [
    {'cluster': 'Rupiah',  'ns': 0.350, 'va': 0.170, 'er': 1.000, 'score': 49.10},
    {'cluster': 'Inflasi', 'ns': 0.140, 'va': 0.000, 'er': 0.250, 'score': 13.10},
    {'cluster': 'BI-Rate', 'ns': 0.216, 'va': 0.304, 'er': 0.271, 'score': 25.89},
]

IIRS_DATA = [
    {
        'cluster': 'Rupiah',
        'score_nolimit': 49.10,
        'score_sentrix': 46.07,
        'iirs': 47.282,
    },
    {
        'cluster': 'Inflasi',
        'score_nolimit': 13.10,
        'score_sentrix': 46.70,
        'iirs': 34.94,
    },
    {
        'cluster': 'BI-Rate',
        'score_nolimit': 25.89,
        'score_sentrix': 45.10,
        'iirs': 38.3765,
    },
]

PLAYBOOKS = {
    'Playbook 1': {
        'title': 'Playbook 1 — Kategori Ringan',
        'range': 'IIRS 0–35',
        'color': 'success',
        'icon': '✅',
        'actions': [
            'Monitor rutin sentimen publik',
            'Respons standar melalui kanal komunikasi resmi',
            'Update data berkala (mingguan)',
            'Tidak diperlukan eskalasi khusus',
        ]
    },
    'Playbook 2': {
        'title': 'Playbook 2 — Kategori Sedang',
        'range': 'IIRS 36–49',
        'color': 'warning',
        'icon': '⚠️',
        'actions': [
            'Peningkatan frekuensi monitoring sentimen',
            'Koordinasi internal tim komunikasi',
            'Siapkan narasi klarifikasi & FAQ publik',
            'Evaluasi mingguan dengan pimpinan',
        ]
    },
    'Crisis Watch': {
        'title': 'Crisis Watch — Zona Pre-Krisis',
        'range': 'IIRS 50–59 (Klaster Strategis)',
        'color': 'orange',
        'icon': '🔶',
        'actions': [
            'Aktivasi tim respons krisis',
            'Monitoring real-time 24/7',
            'Siapkan pernyataan resmi & press release',
            'Koordinasi dengan stakeholder eksternal',
            'Eskalasi ke manajemen senior',
        ]
    },
    'Playbook 3': {
        'title': 'Playbook 3 — Kategori Crisis',
        'range': 'IIRS ≥ 60',
        'color': 'danger',
        'icon': '🚨',
        'actions': [
            'Aktivasi penuh protokol krisis komunikasi',
            'War room & command center aktif',
            'Pernyataan resmi pimpinan tertinggi',
            'Koordinasi lintas kementerian/lembaga',
            'Media briefing & press conference segera',
            'Laporan harian kepada Gubernur BI',
        ]
    },
}

def get_kategori(iirs):
    if iirs < 35:
        return 'Ringan', 'Playbook 1', 'success'
    elif iirs < 50:
        return 'Sedang', 'Playbook 2', 'warning'
    elif iirs < 60:
        return 'Crisis Watch', 'Crisis Watch', 'orange'
    else:
        return 'Crisis', 'Playbook 3', 'danger'

@app.route('/', methods=['GET', 'POST'])
def login():
    if 'user' in session:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        if username in USERS and USERS[username] == password:
            session['user'] = username
            return redirect(url_for('dashboard'))
        flash('Username atau password salah.', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))

    iirs_enriched = []
    for row in IIRS_DATA:
        kategori, playbook_key, color = get_kategori(row['iirs'])
        iirs_enriched.append({
            **row,
            'kategori': kategori,
            'playbook_key': playbook_key,
            'color': color,
            'playbook': PLAYBOOKS.get(playbook_key, {}),
        })

    chart_labels = [r['cluster'] for r in IIRS_DATA]
    chart_sentrix = [r['score_sentrix'] for r in IIRS_DATA]
    chart_nolimit = [r['score_nolimit'] for r in IIRS_DATA]
    chart_iirs = [round(r['iirs'], 2) for r in IIRS_DATA]

    return render_template(
        'dashboard.html',
        user=session['user'],
        sentrix=SOURCE_SCORE_SENTRIX,
        nolimit=SOURCE_SCORE_NOLIMIT,
        iirs_data=iirs_enriched,
        playbooks=PLAYBOOKS,
        chart_labels=json.dumps(chart_labels),
        chart_sentrix=json.dumps(chart_sentrix),
        chart_nolimit=json.dumps(chart_nolimit),
        chart_iirs=json.dumps(chart_iirs),
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
