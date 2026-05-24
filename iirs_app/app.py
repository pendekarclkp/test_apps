from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import json, os, uuid
from datetime import datetime
import openpyxl

app = Flask(__name__)
app.secret_key = 'iirs_secret_key_2026'

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
DATA_FILE     = os.path.join(BASE_DIR, 'data', 'current_data.json')
HISTORY_FILE  = os.path.join(BASE_DIR, 'data', 'history.json')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)

# ── Users: admin = dashboard only | uploader = dashboard + upload ─────────────
USERS = {
    'admin':    {'password': 'admin',    'role': 'admin'},
    'uploader': {'password': 'upload123','role': 'uploader'},
}

PLAYBOOKS = {
    'Tindakan Respons 1': {
        'title': 'Tindakan Respons 1 — Kategori Ringan',
        'range': 'IIRS 0–35', 'color': 'success', 'icon': '✅',
        'actions': [
            'Monitor rutin sentimen publik',
            'Respons standar melalui kanal komunikasi resmi',
            'Update data berkala (mingguan)',
            'Tidak diperlukan eskalasi khusus',
        ],
    },
    'Tindakan Respons 2': {
        'title': 'Tindakan Respons 2 — Kategori Sedang',
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
    'Tindakan Respons 3': {
        'title': 'Tindakan Respons 3 — Kategori Crisis',
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
    'tindak_lanjut': [],
    'uploaded_at': None,
    'uploaded_by': None,
    'filename': None,
    'periode_bulan': None,
    'periode_tahun': None,
}

# In-memory cache — single source of truth dalam satu proses gunicorn
_cache = {}

def load_data():
    if _cache:
        return _cache
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE) as f:
                _cache.update(json.load(f))
                return _cache
        except Exception:
            pass
    _cache.update(DEFAULT_DATA)
    return _cache

def save_data(new_data):
    _cache.clear()
    _cache.update(new_data)
    try:
        with open(DATA_FILE, 'w') as f:
            json.dump(new_data, f, indent=2, ensure_ascii=False)
    except Exception:
        pass  # disk write gagal tidak masalah, cache tetap update

# ── History helpers ─────────────────────────────────────────────────────────
RETENTION_DAYS = 7  # snapshot otomatis dihapus setelah 7 hari

def _prune_history(history):
    """Buang snapshot lebih lama dari RETENTION_DAYS hari."""
    from datetime import timedelta
    cutoff = datetime.now() - timedelta(days=RETENTION_DAYS)
    kept = []
    for snap in history:
        iso = snap.get('saved_at_iso')
        if not iso:
            # Snapshot lama tanpa timestamp ISO → simpan saja (backward compat)
            kept.append(snap)
            continue
        try:
            if datetime.fromisoformat(iso) >= cutoff:
                kept.append(snap)
        except (ValueError, TypeError):
            kept.append(snap)
    return kept

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE) as f:
                history = json.load(f)
        except Exception:
            return []
        pruned = _prune_history(history)
        if len(pruned) != len(history):
            # Persist hasil cleanup
            try:
                with open(HISTORY_FILE, 'w') as f:
                    json.dump(pruned, f, indent=2, ensure_ascii=False)
            except Exception:
                pass
        return pruned
    return []

def save_history(snap):
    history = load_history()
    history.insert(0, snap)
    history = _prune_history(history)[:20]  # max 20 entri & dalam retensi
    try:
        with open(HISTORY_FILE, 'w') as f:
            json.dump(history, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

BULAN_ID = {
    'januari':'Januari','februari':'Februari','maret':'Maret','april':'April',
    'mei':'Mei','juni':'Juni','juli':'Juli','agustus':'Agustus',
    'september':'September','oktober':'Oktober','november':'November','desember':'Desember',
    'january':'Januari','february':'Februari','march':'Maret','may':'Mei',
    'june':'Juni','july':'Juli','august':'Agustus','october':'Oktober',
}

def extract_period(wb):
    """Cari 'Periode <Bulan> <Tahun>' di semua sheet, kembalikan string mis. 'April 2026'."""
    import re
    pattern = re.compile(
        r'periode\s+([A-Za-z]+)\s+(\d{4})',
        re.IGNORECASE
    )
    for sheet in wb.sheetnames:
        for row in wb[sheet].iter_rows(values_only=True):
            for cell in row:
                if not cell:
                    continue
                m = pattern.search(str(cell))
                if m:
                    bulan_raw = m.group(1).lower()
                    tahun     = m.group(2)
                    bulan     = BULAN_ID.get(bulan_raw, m.group(1).capitalize())
                    return bulan, tahun
    return None, None

def parse_tindak_lanjut(ws):
    """Parse sheet Tindak Lanjut menjadi dict {headers, rows}.
    headers = label kolom tier dari baris header sheet.
    Row dengan kolom A = None adalah lanjutan aspek sebelumnya — gabung dengan newline.
    """
    result = []
    current = None
    headers = []

    for row in ws.iter_rows(values_only=True):
        # Baris header — ambil label tier dari kolom B/C/D
        if row[0] and str(row[0]).strip().lower() in ('aspek',):
            headers = [str(row[i]).strip() if row[i] is not None else '' for i in range(1, 4)]
            continue

        aspek = row[0]
        t1    = str(row[1]).strip() if row[1] is not None else ''
        t2    = str(row[2]).strip() if row[2] is not None else ''
        t3    = str(row[3]).strip() if row[3] is not None else ''

        if aspek is not None:
            # Baris baru — mulai aspek baru
            if current is not None:
                result.append(current)
            current = {
                'aspek': str(aspek).strip(),
                'tier1': t1,
                'tier2': t2,
                'tier3': t3,
            }
        else:
            # Lanjutan baris sebelumnya
            if current is not None:
                if t1:
                    current['tier1'] = (current['tier1'] + '\n' + t1).strip('\n')
                if t2:
                    current['tier2'] = (current['tier2'] + '\n' + t2).strip('\n')
                if t3:
                    current['tier3'] = (current['tier3'] + '\n' + t3).strip('\n')

    if current is not None:
        result.append(current)

    return {'headers': headers, 'rows': result}

def parse_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # Ekstrak periode dari file
    periode_bulan, periode_tahun = extract_period(wb)

    ws   = wb['Source Scoring']
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

    ws2    = wb['IIRS Integration']
    rows2  = list(ws2.iter_rows(values_only=True))
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
        raise ValueError('Format Excel tidak sesuai. Butuh sheet "Source Scoring" dan "IIRS Integration".')

    # Parse sheet Tindak Lanjut (opsional)
    tindak_lanjut = []
    tindak_lanjut_headers = []
    if 'Tindak Lanjut' in wb.sheetnames:
        try:
            parsed = parse_tindak_lanjut(wb['Tindak Lanjut'])
            tindak_lanjut = parsed['rows']
            tindak_lanjut_headers = parsed['headers']
        except Exception:
            tindak_lanjut = []
            tindak_lanjut_headers = []

    return sentrix, nolimit, iirs, periode_bulan, periode_tahun, tindak_lanjut, tindak_lanjut_headers

def get_kategori(iirs_val):
    if iirs_val < 35:
        return 'Ringan',       'Tindakan Respons 1',   'success'
    elif iirs_val < 50:
        return 'Sedang',       'Tindakan Respons 2',   'warning'
    elif iirs_val < 60:
        return 'Crisis Watch', 'Crisis Watch', 'orange'
    else:
        return 'Crisis',       'Tindakan Respons 3',   'danger'

def current_user():
    return session.get('user')

def current_role():
    return USERS.get(current_user(), {}).get('role')

def require_login():
    if not current_user():
        return redirect(url_for('login'))
    return None

def require_role(role):
    err = require_login()
    if err:
        return err
    if current_role() != role:
        flash('Akses ditolak untuk role Anda.', 'danger')
        return redirect(url_for('dashboard'))
    return None

def enrich_iirs(iirs_list):
    enriched = []
    for row in iirs_list:
        kategori, playbook_key, color = get_kategori(row['iirs'])
        enriched.append({
            **row,
            'kategori':     kategori,
            'playbook_key': playbook_key,
            'color':        color,
            'playbook':     PLAYBOOKS.get(playbook_key, {}),
        })
    return enriched

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
def login():
    if current_user():
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        user = USERS.get(username)
        if user and user['password'] == password:
            session['user'] = username
            session['role'] = user['role']
            return redirect(url_for('dashboard'))
        flash('Username atau password salah.', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
def dashboard():
    err = require_login()
    if err:
        return err

    snap_id   = request.args.get('snapshot')
    snap_data = None

    if snap_id:
        history = load_history()
        snap_data = next((s for s in history if s['id'] == snap_id), None)

    if snap_data:
        data         = snap_data
        iirs_list    = snap_data.get('iirs', [])
        tindak_lanjut = snap_data.get('tindak_lanjut', [])
        tindak_lanjut_headers = snap_data.get('tindak_lanjut_headers', [])
        is_snapshot  = True
    else:
        data         = load_data()
        iirs_list    = data['iirs']
        tindak_lanjut = data.get('tindak_lanjut', [])
        tindak_lanjut_headers = data.get('tindak_lanjut_headers', [])
        is_snapshot  = False
        snap_id      = None

    iirs_enriched = enrich_iirs(iirs_list)

    # Hitung tier per klaster (Tier 1: <35, Tier 2: 35-59, Tier 3: >=60)
    def to_tier(v):
        if v < 35:   return 1
        elif v < 60: return 2
        else:        return 3

    cluster_tiers = []
    for row in iirs_enriched:
        tier = to_tier(row['iirs'])
        row['tier'] = tier
        cluster_tiers.append({
            'cluster': row['cluster'],
            'iirs':    row['iirs'],
            'tier':    tier,
            'tier_label': {1:'Tier 1 Ringan',2:'Tier 2 Sedang',3:'Tier 3 Crisis'}[tier],
            'color':   {1:'success',2:'warning',3:'danger'}[tier],
        })

    # Tier aktif = tier paling parah dari semua klaster
    active_tier = max([c['tier'] for c in cluster_tiers], default=1)
    tiers_active = sorted({c['tier'] for c in cluster_tiers}) if cluster_tiers else [1]
    avg_iirs = (sum(r['iirs'] for r in iirs_list) / len(iirs_list)) if iirs_list else 0

    # Breakdown indikator NS/VA/ER rata-rata per kanal (internal vs eksternal)
    sentrix_list = data.get('sentrix', [])
    nolimit_list = data.get('nolimit', [])
    def avg_field(lst, field):
        vals = [r.get(field, 0) for r in lst]
        return round(sum(vals) / len(vals), 3) if vals else 0
    breakdown_internal = [avg_field(sentrix_list, 'ns'), avg_field(sentrix_list, 'va'), avg_field(sentrix_list, 'er')]
    breakdown_external = [avg_field(nolimit_list, 'ns'), avg_field(nolimit_list, 'va'), avg_field(nolimit_list, 'er')]

    # Align NS/VA/ER per cluster sesuai urutan iirs_list (pakai lookup by cluster name)
    sx_by_cluster = {r['cluster']: r for r in sentrix_list}
    nl_by_cluster = {r['cluster']: r for r in nolimit_list}
    clusters = [r['cluster'] for r in iirs_list]
    def pick(lookup, cluster, field, weight):
        row = lookup.get(cluster, {})
        return round(row.get(field, 0), 4)

    # Sinyal strategis dinamis berdasarkan data terkini
    signals = []
    tier_label = {1: 'Ringan', 2: 'Sedang', 3: 'Crisis'}

    saturated = [r['cluster'] for r in iirs_enriched
                 if sx_by_cluster.get(r['cluster'], {}).get('er', 0) >= 0.99]
    if saturated:
        signals.append({
            'type': 'warn',
            'title': 'ER Jenuh — ' + ', '.join(saturated) + ' (Sentrix)',
            'body': 'ER Sentrix pada ' + ', '.join(saturated) + ' sudah mencapai 100% (saturasi). '
                    'Kenaikan NS pada periode berikutnya langsung berdampak signifikan menaikkan IIRS — tidak ada ruang penyangga jangkauan.',
        })

    if iirs_enriched:
        top = max(iirs_enriched, key=lambda r: r['iirs'])
        if 45 <= top['iirs'] < 50:
            signals.append({
                'type': 'warn',
                'title': top['cluster'] + ': Waspadai Eskalasi ke Crisis Watch',
                'body': 'IIRS ' + top['cluster'] + ' (' + ('%.2f' % top['iirs']) + ') mendekati ambang Siaga (50). '
                        'Diperlukan langkah proaktif agar tidak melewati fase Siaga pada periode mendatang.',
            })

    gaps = sorted(iirs_enriched, key=lambda r: abs(r['score_nolimit'] - r['score_sentrix']), reverse=True)
    if gaps:
        g = gaps[0]
        diff = abs(g['score_nolimit'] - g['score_sentrix'])
        if diff >= 15:
            signals.append({
                'type': 'info',
                'title': g['cluster'] + ': Gap Dua Kanal (' + ('%.0f' % diff) + ' Poin)',
                'body': 'IIRS ' + g['cluster'] + ' dibentuk dari ketimpangan Sentrix ' + ('%.2f' % g['score_sentrix']) +
                        ' dan NoLimit ' + ('%.2f' % g['score_nolimit']) + '. Strategi bridging edukatif diperlukan.',
            })

    if active_tier < 3:
        crisis_watch = [r['cluster'] for r in iirs_enriched if 50 <= r['iirs'] < 60]
        if not crisis_watch:
            signals.append({
                'type': 'ok',
                'title': 'Tidak Ada Keyword pada Level Crisis',
                'body': 'Tidak ada keyword yang mencapai Crisis (≥60) atau Crisis Watch (50–59) pada periode ini. '
                        'Protokol Tier 3 dan fase Siaga tidak diaktifkan.',
            })

    # Panduan respon singkat per keyword sesuai tier & sheet Tindak Lanjut
    tl_by_aspek = {str(r.get('aspek', '')).strip().lower(): r for r in tindak_lanjut}
    esensi_row = tl_by_aspek.get('esensi', {})
    eskalasi_row = next((r for r in tindak_lanjut if 'eskalasi' in str(r.get('aspek', '')).lower()), {})
    panduan_respon = []
    for r in iirs_enriched:
        tk = 'tier' + str(r['tier'])
        panduan_respon.append({
            'cluster':  r['cluster'],
            'iirs':     r['iirs'],
            'tier':     r['tier'],
            'color':    r['color'],
            'kategori': r['kategori'],
            'esensi':   esensi_row.get(tk, ''),
            'eskalasi': eskalasi_row.get(tk, ''),
        })

    return render_template(
        'dashboard.html',
        user=current_user(),
        role=current_role(),
        sentrix=data.get('sentrix', []),
        nolimit=data.get('nolimit', []),
        iirs_data=iirs_enriched,
        playbooks=PLAYBOOKS,
        uploaded_at=data.get('uploaded_at'),
        uploaded_by=data.get('uploaded_by'),
        filename=data.get('filename'),
        periode_bulan=data.get('periode_bulan'),
        periode_tahun=data.get('periode_tahun'),
        tindak_lanjut=tindak_lanjut,
        tindak_lanjut_headers=tindak_lanjut_headers,
        active_tier=active_tier,
        cluster_tiers=cluster_tiers,
        tiers_active=tiers_active,
        avg_iirs=round(avg_iirs, 2),
        is_snapshot=is_snapshot,
        snapshot_id=snap_id,
        snapshot_saved_at=data.get('saved_at') if is_snapshot else None,
        chart_labels =json.dumps([r['cluster'] for r in iirs_list]),
        chart_sentrix=json.dumps([r['score_sentrix'] for r in iirs_list]),
        chart_nolimit=json.dumps([r['score_nolimit'] for r in iirs_list]),
        chart_iirs   =json.dumps([round(r['iirs'], 2) for r in iirs_list]),
        chart_sentrix_ns=json.dumps([pick(sx_by_cluster, c, 'ns', 40) for c in clusters]),
        chart_sentrix_va=json.dumps([pick(sx_by_cluster, c, 'va', 30) for c in clusters]),
        chart_sentrix_er=json.dumps([pick(sx_by_cluster, c, 'er', 30) for c in clusters]),
        chart_nolimit_ns=json.dumps([pick(nl_by_cluster, c, 'ns', 40) for c in clusters]),
        chart_nolimit_va=json.dumps([pick(nl_by_cluster, c, 'va', 30) for c in clusters]),
        chart_nolimit_er=json.dumps([pick(nl_by_cluster, c, 'er', 30) for c in clusters]),
        chart_breakdown_internal=json.dumps(breakdown_internal),
        chart_breakdown_external=json.dumps(breakdown_external),
        signals=signals,
        panduan_respon=panduan_respon,
    )

# Upload: hanya role 'uploader'
@app.route('/upload', methods=['GET', 'POST'])
def upload_page():
    err = require_role('uploader')
    if err:
        return err

    data     = load_data()
    preview  = None

    if request.method == 'POST':
        f = request.files.get('excel_file')
        if not f or not f.filename.endswith(('.xlsx', '.xls')):
            flash('Harap upload file Excel (.xlsx atau .xls).', 'danger')
        else:
            save_path = os.path.join(UPLOAD_FOLDER, 'latest.xlsx')
            f.save(save_path)
            try:
                sentrix, nolimit, iirs, periode_bulan, periode_tahun, tindak_lanjut, tindak_lanjut_headers = parse_excel(save_path)
                new_data = {
                    'sentrix':        sentrix,
                    'nolimit':        nolimit,
                    'iirs':           iirs,
                    'tindak_lanjut':  tindak_lanjut,
                    'tindak_lanjut_headers': tindak_lanjut_headers,
                    'uploaded_at':    datetime.now().strftime('%d %b %Y, %H:%M WIB'),
                    'uploaded_by':    current_user(),
                    'filename':       f.filename,
                    'periode_bulan':  periode_bulan,
                    'periode_tahun':  periode_tahun,
                }
                save_data(new_data)
                data    = load_data()
                preview = {'sentrix': sentrix, 'nolimit': nolimit, 'iirs': iirs}
                flash(f'✅ Data berhasil diupload dari "{f.filename}". Dashboard sudah terupdate.', 'success')
            except Exception as e:
                flash(f'❌ Gagal membaca file: {e}', 'danger')

    return render_template(
        'upload.html',
        user=current_user(),
        role=current_role(),
        data=data,
        preview=preview,
    )

# ── Admin: Save Snapshot ──────────────────────────────────────────────────────

@app.route('/save-snapshot', methods=['POST'])
def save_snapshot():
    err = require_role('admin')
    if err:
        return err

    data = load_data()
    now = datetime.now()
    snap = {
        'id':           str(uuid.uuid4()),
        'saved_at':     now.strftime('%d %b %Y, %H:%M WIB'),
        'saved_at_iso': now.isoformat(),
        'saved_by':     current_user(),
        'periode_bulan': data.get('periode_bulan'),
        'periode_tahun': data.get('periode_tahun'),
        'filename':     data.get('filename'),
        'sentrix':      data.get('sentrix', []),
        'nolimit':      data.get('nolimit', []),
        'iirs':         data.get('iirs', []),
        'tindak_lanjut': data.get('tindak_lanjut', []),
    }
    save_history(snap)
    flash(f'✅ Snapshot dashboard berhasil disimpan.', 'success')
    return redirect(url_for('dashboard'))

# ── Admin: History ────────────────────────────────────────────────────────────

@app.route('/history')
def history_page():
    err = require_role('admin')
    if err:
        return err

    from datetime import timedelta
    history = load_history()
    now = datetime.now()
    for snap in history:
        iso = snap.get('saved_at_iso')
        if iso:
            try:
                saved = datetime.fromisoformat(iso)
                expires = saved + timedelta(days=RETENTION_DAYS)
                remaining = expires - now
                days = max(0, remaining.days)
                hours = max(0, remaining.seconds // 3600) if remaining.total_seconds() > 0 else 0
                snap['expires_in'] = f'{days} hari {hours} jam' if days > 0 else f'{hours} jam'
            except (ValueError, TypeError):
                snap['expires_in'] = None
        else:
            snap['expires_in'] = None

    return render_template(
        'history.html',
        user=current_user(),
        role=current_role(),
        history=history,
        retention_days=RETENTION_DAYS,
    )

@app.route('/history/<snap_id>')
def history_detail(snap_id):
    err = require_role('admin')
    if err:
        return err
    return redirect(url_for('dashboard', snapshot=snap_id))

# ── Admin: Delete Snapshot ────────────────────────────────────────────────────

@app.route('/delete-snapshot/<snap_id>', methods=['POST'])
def delete_snapshot(snap_id):
    err = require_role('admin')
    if err:
        return err

    history = load_history()
    history = [s for s in history if s['id'] != snap_id]
    try:
        with open(HISTORY_FILE, 'w') as f:
            json.dump(history, f, indent=2, ensure_ascii=False)
    except Exception:
        pass
    flash('Snapshot berhasil dihapus.', 'success')
    return redirect(url_for('history_page'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
