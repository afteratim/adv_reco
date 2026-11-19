import os
import re
import io
import csv
import json
import html
import hmac
import uuid
import time
import zipfile
import secrets
import logging
import threading
from pathlib import Path
from datetime import date, datetime, timedelta, timezone
from collections import defaultdict, Counter
from functools import wraps
from difflib import SequenceMatcher

import psycopg
from psycopg.rows import dict_row
from flask import Flask, request, session, redirect, url_for, render_template, flash, send_file, abort, jsonify
from flask_wtf.csrf import CSRFProtect, generate_csrf
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_USERNAME = os.environ.get('APP_USERNAME', 'admin')
APP_PASSWORD = os.environ.get('APP_PASSWORD', '')
SECRET_KEY = os.environ.get('SECRET_KEY') or secrets.token_hex(32)
DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()
MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_REQUEST_BYTES = 30 * 1024 * 1024
LOCAL_DATA_DIR = Path(os.environ.get('LOCAL_DATA_DIR', 'instance/local_history'))
LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)
LOCAL_MIRROR = LOCAL_DATA_DIR / 'history_mirror.json'
BACKUP_MAX_BYTES = 25 * 1024 * 1024

app = Flask(__name__)
app.config.update(
    SECRET_KEY=SECRET_KEY,
    MAX_CONTENT_LENGTH=MAX_REQUEST_BYTES,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SECURE=os.environ.get('COOKIE_SECURE', 'true').lower() == 'true',
    SESSION_COOKIE_SAMESITE='Lax',
    PERMANENT_SESSION_LIFETIME=60 * 60 * 8,
    WTF_CSRF_TIME_LIMIT=3600,
)
CSRFProtect(app)
logging.basicConfig(level=os.environ.get('LOG_LEVEL', 'WARNING'))
logger = logging.getLogger('gst_smart_assistant')
_db_ready = False
_db_lock = threading.Lock()
_local_lock = threading.Lock()


def login_required(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if not session.get('authenticated'):
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapped


def db_connect():
    if not DATABASE_URL:
        raise RuntimeError('DATABASE_URL is not configured. Add your Neon PostgreSQL connection string.')
    return psycopg.connect(DATABASE_URL, row_factory=dict_row, connect_timeout=8)


def ensure_db():
    global _db_ready
    if not DATABASE_URL:
        return False
    if _db_ready:
        return True
    with _db_lock:
        if _db_ready:
            return True
        with db_connect() as conn:
            conn.execute('''CREATE TABLE IF NOT EXISTS gst_analyses (
                id TEXT PRIMARY KEY, month_label TEXT NOT NULL, source_filename TEXT NOT NULL,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), report_json JSONB NOT NULL,
                settings_json JSONB NOT NULL DEFAULT '{}'::jsonb)''')
            conn.execute('''CREATE TABLE IF NOT EXISTS gst_chat_sessions (
                id TEXT PRIMARY KEY, analysis_id TEXT, title TEXT NOT NULL,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW())''')
            conn.execute('''CREATE TABLE IF NOT EXISTS gst_chat_messages (
                id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES gst_chat_sessions(id) ON DELETE CASCADE,
                role TEXT NOT NULL CHECK (role IN ('user','assistant')), content TEXT NOT NULL,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW())''')
            conn.execute('''CREATE TABLE IF NOT EXISTS gst_settings (
                key TEXT PRIMARY KEY, value_json JSONB NOT NULL, updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW())''')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_gst_sessions_updated ON gst_chat_sessions(updated_at DESC)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_gst_messages_session ON gst_chat_messages(session_id, created_at)')
        _db_ready = True
    return True


def utc_now():
    return datetime.now(timezone.utc).isoformat()


def _json_default(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    raise TypeError(f'Not JSON serializable: {type(value).__name__}')


def load_local_data():
    if not LOCAL_MIRROR.exists():
        return {'analyses': [], 'sessions': [], 'messages': [], 'settings': []}
    try:
        data = json.loads(LOCAL_MIRROR.read_text(encoding='utf-8'))
        for key in ('analyses','sessions','messages','settings'):
            data.setdefault(key, [])
        return data
    except Exception:
        logger.warning('Local history mirror could not be read; starting with empty local cache.')
        return {'analyses': [], 'sessions': [], 'messages': [], 'settings': []}


def save_local_data(data):
    LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp = LOCAL_MIRROR.with_suffix('.tmp')
    with _local_lock:
        temp.write_text(json.dumps(data, ensure_ascii=False, default=_json_default), encoding='utf-8')
        temp.replace(LOCAL_MIRROR)


def fetch_all_data():
    if ensure_db():
        with db_connect() as conn:
            analyses = conn.execute('SELECT id, month_label, source_filename, created_at, report_json, settings_json FROM gst_analyses ORDER BY created_at DESC').fetchall()
            sessions = conn.execute('SELECT id, analysis_id, title, created_at, updated_at FROM gst_chat_sessions ORDER BY updated_at DESC').fetchall()
            messages = conn.execute('SELECT id, session_id, role, content, created_at FROM gst_chat_messages ORDER BY created_at').fetchall()
            settings = conn.execute('SELECT key, value_json, updated_at FROM gst_settings ORDER BY key').fetchall()
        data = {'analyses': analyses, 'sessions': sessions, 'messages': messages, 'settings': settings}
        save_local_data(data)
        return data
    return load_local_data()


def get_analysis(analysis_id):
    if ensure_db():
        with db_connect() as conn:
            row = conn.execute('SELECT id, month_label, source_filename, created_at, report_json, settings_json FROM gst_analyses WHERE id=%s', (analysis_id,)).fetchone()
        return row
    return next((x for x in load_local_data()['analyses'] if x.get('id') == analysis_id), None)


def save_analysis(analysis_id, month_label, filename, report, settings=None):
    settings = settings or {}
    if ensure_db():
        with db_connect() as conn:
            conn.execute('''INSERT INTO gst_analyses(id,month_label,source_filename,report_json,settings_json)
                VALUES(%s,%s,%s,%s::jsonb,%s::jsonb)
                ON CONFLICT(id) DO UPDATE SET month_label=EXCLUDED.month_label, source_filename=EXCLUDED.source_filename,
                report_json=EXCLUDED.report_json, settings_json=EXCLUDED.settings_json''',
                (analysis_id, month_label, filename, json.dumps(report, default=_json_default), json.dumps(settings)))
    data = fetch_all_data()
    if not ensure_db():
        data['analyses'] = [x for x in data['analyses'] if x.get('id') != analysis_id]
        data['analyses'].insert(0, {'id':analysis_id,'month_label':month_label,'source_filename':filename,
            'created_at':utc_now(),'report_json':report,'settings_json':settings})
        save_local_data(data)
    else:
        fetch_all_data()


def create_chat_session(analysis_id, title='New Chat'):
    sid = uuid.uuid4().hex
    if ensure_db():
        with db_connect() as conn:
            conn.execute('INSERT INTO gst_chat_sessions(id,analysis_id,title) VALUES(%s,%s,%s)', (sid, analysis_id, title))
            conn.execute('INSERT INTO gst_chat_messages(id,session_id,role,content) VALUES(%s,%s,%s,%s)',
                         (uuid.uuid4().hex, sid, 'assistant', 'Hello! Ask me about this GST 2B reconciliation, or choose a suggested question below.'))
        fetch_all_data()
    else:
        data=load_local_data(); now=utc_now()
        data['sessions'].insert(0, {'id':sid,'analysis_id':analysis_id,'title':title,'created_at':now,'updated_at':now})
        data['messages'].append({'id':uuid.uuid4().hex,'session_id':sid,'role':'assistant','content':'Hello! Ask me about this GST 2B reconciliation, or choose a suggested question below.','created_at':now})
        save_local_data(data)
    return sid


def get_session(session_id):
    if ensure_db():
        with db_connect() as conn:
            return conn.execute('SELECT * FROM gst_chat_sessions WHERE id=%s', (session_id,)).fetchone()
    return next((x for x in load_local_data()['sessions'] if x.get('id') == session_id), None)


def get_messages(session_id):
    if ensure_db():
        with db_connect() as conn:
            return conn.execute('SELECT id,session_id,role,content,created_at FROM gst_chat_messages WHERE session_id=%s ORDER BY created_at', (session_id,)).fetchall()
    return [x for x in load_local_data()['messages'] if x.get('session_id') == session_id]


def add_message(session_id, role, content):
    if ensure_db():
        with db_connect() as conn:
            conn.execute('INSERT INTO gst_chat_messages(id,session_id,role,content) VALUES(%s,%s,%s,%s)', (uuid.uuid4().hex,session_id,role,content))
            conn.execute('UPDATE gst_chat_sessions SET updated_at=NOW() WHERE id=%s', (session_id,))
        fetch_all_data()
    else:
        data=load_local_data(); now=utc_now()
        data['messages'].append({'id':uuid.uuid4().hex,'session_id':session_id,'role':role,'content':content,'created_at':now})
        for item in data['sessions']:
            if item.get('id')==session_id: item['updated_at']=now
        save_local_data(data)


def rename_session(session_id, title):
    title=(title or '').strip()[:80] or 'New Chat'
    if ensure_db():
        with db_connect() as conn: conn.execute('UPDATE gst_chat_sessions SET title=%s,updated_at=NOW() WHERE id=%s',(title,session_id))
        fetch_all_data()
    else:
        data=load_local_data()
        for item in data['sessions']:
            if item.get('id')==session_id: item['title']=title; item['updated_at']=utc_now()
        save_local_data(data)


def delete_session(session_id):
    if ensure_db():
        with db_connect() as conn: conn.execute('DELETE FROM gst_chat_sessions WHERE id=%s',(session_id,))
        fetch_all_data()
    else:
        data=load_local_data(); data['sessions']=[x for x in data['sessions'] if x.get('id')!=session_id]
        data['messages']=[x for x in data['messages'] if x.get('session_id')!=session_id]; save_local_data(data)


def month_from_purchase_rows(rows, fallback_rows=None):
    counts=Counter()
    for row in rows or []:
        d=parse_gst_date(gst_date(gst_field(row,'DATE')))
        if d: counts[(d.year,d.month)] += 1
    if not counts:
        for row in fallback_rows or []:
            d=parse_gst_date(gst_date(gst_field(row,'DATE')))
            if d: counts[(d.year,d.month)] += 1
    if not counts: return datetime.now().strftime('%b %Y')
    year, month=counts.most_common(1)[0][0]
    return datetime(year,month,1).strftime('%b %Y')


def month_filename(label):
    try:
        dt=datetime.strptime(label,'%b %Y')
        return f'Report of {dt.strftime("%b")} {dt.strftime("%y")}.xlsx'
    except Exception:
        return 'GST 2B Reconciliation Report.xlsx'


def month_chat_filename(label):
    return f'GST 2B Reco - {label.replace(" ", " ")}.html'


def report_summary(report):
    lines=report.get('reconciliation',[]); counts=report.get('counts',{})
    matched=counts.get('Matched',0)+counts.get('Fuzzy Matched',0)
    discrepancy=[x for x in lines if x.get('remark') not in ('Matched','Fuzzy Matched')]
    taxable=sum(abs(x.get('taxable_diff') or 0) for x in discrepancy)
    tax=sum(abs(x.get('igst_diff') or 0)+abs(x.get('cgst_diff') or 0)+abs(x.get('sgst_diff') or 0) for x in discrepancy)
    return {'pur':report.get('pur_count',0),'two_b':report.get('portal_count',0),'matched':matched,
        'discrepancies':len(discrepancy),'validation':len(report.get('validation',[])),
        'duplicate_groups':len(report.get('duplicates',[])),'taxable_abs_diff':taxable,'tax_abs_diff':tax,
        'counts':counts}


def assistant_answer(question, report, analysis_id):
    q=(question or '').strip(); low=q.casefold(); lines=report.get('reconciliation',[])
    if not q: return 'Please type a question or choose one of the suggested prompts.'
    summary=report_summary(report)
    if any(k in low for k in ('summary','overview','status','how many','total')):
        return (f"**GST 2B reconciliation summary – {report.get('month_label','')}**\n\n"
            f"- Purchase rows: {summary['pur']}\n- 2B rows: {summary['two_b']}\n- Matched (exact + fuzzy): {summary['matched']}\n"
            f"- Discrepancy / review rows: {summary['discrepancies']}\n- Validation warnings: {summary['validation']}\n"
            f"- Duplicate / mismatch audit entries: {summary['duplicate_groups']}\n"
            f"- Absolute taxable difference across discrepancy rows: ₹{summary['taxable_abs_diff']:,.2f}\n"
            f"- Absolute tax-component differences (summed): ₹{summary['tax_abs_diff']:,.2f}\n\n"
            "These are report calculations, not an independent GST compliance opinion.")
    if any(k in low for k in ('duplicate','repeated','same invoice')):
        rows=[x for x in lines if 'duplicate' in x.get('remark','').casefold() or 'duplicate' in x.get('detail','').casefold()]
        audit=report.get('duplicates',[])
        if not rows and not audit: return 'No duplicate or consolidation audit entries were found in this analysis.'
        preview='\n'.join(f"- {x.get('name','')} | GSTIN {x.get('gstin','')} | Pur {x.get('invoice_pur','')} | 2B {x.get('invoice_2b','')} | {x.get('remark','')}" for x in rows[:20])
        if not preview: preview='\n'.join('- '+' | '.join(map(str,x[:5])) for x in audit[:20])
        return f"Found {len(rows)} reconciliation row(s) and {len(audit)} duplicate/consolidation audit entry/entries.\n\n{preview}"
    if any(k in low for k in ('unmatched purchase','not on portal','missing in 2b','purchase missing')):
        rows=[x for x in lines if x.get('remark')=='Not on Portal']
        return format_rows('Purchase invoices not found in 2B',rows)
    if any(k in low for k in ('not in tally','missing in purchase','unmatched 2b','2b missing')):
        rows=[x for x in lines if x.get('remark')=='Not in Tally']
        return format_rows('2B invoices not found in Purchase',rows)
    if any(k in low for k in ('needs review','review required','fuzzy review','gstin mismatch','ambiguous')):
        rows=[x for x in lines if 'review' in x.get('remark','').casefold() or 'mismatch' in x.get('remark','').casefold() or 'duplicate' in x.get('remark','').casefold()]
        return format_rows('Invoices requiring manual review',rows)
    if any(k in low for k in ('tax difference','tax mismatch','igst','cgst','sgst','tax diff')):
        rows=[x for x in lines if any(abs(x.get(k) or 0)>report.get('tolerance',2) for k in ('igst_diff','cgst_diff','sgst_diff'))]
        return format_rows('Tax-component differences',rows)
    if any(k in low for k in ('high value','large mismatch','big mismatch','high-value')):
        rows=[x for x in lines if x.get('remark') not in ('Matched','Fuzzy Matched') and max([abs(x.get(k) or 0) for k in ('taxable_diff','invoice_value_diff','igst_diff','cgst_diff','sgst_diff')])>=10000]
        return format_rows('High-value discrepancies (at least ₹10,000 absolute difference)',rows)
    if any(k in low for k in ('party wise','party-wise','top party','supplier summary','vendor summary')):
        totals=defaultdict(float); counts=Counter()
        for x in lines:
            if x.get('remark') not in ('Matched','Fuzzy Matched'):
                name=x.get('name') or 'Unknown party'; totals[name]+=abs(x.get('taxable_diff') or 0); counts[name]+=1
        ordered=sorted(totals.items(),key=lambda x:x[1],reverse=True)[:10]
        if not ordered:return 'No discrepancy-bearing parties were found.'
        return 'Top parties by absolute taxable difference:\n\n'+'\n'.join(f"{i}. {name} — ₹{amount:,.2f} taxable difference across {counts[name]} row(s)" for i,(name,amount) in enumerate(ordered,1))
    if any(k in low for k in ('excel','download','export report','report file')):
        return f"Your complete Excel reconciliation report is ready: {url_for('download_analysis',analysis_id=analysis_id)}"
    terms=[x for x in re.split(r'\s+',q) if len(x)>=4]
    matches=[]
    for x in lines:
        hay=' '.join(str(x.get(k) or '') for k in ('name','gstin','invoice_pur','invoice_2b','detail','remark')).casefold()
        if any(t.casefold() in hay for t in terms): matches.append(x)
    if matches:
        return format_rows(f"Search results for: {q}",matches)
    return ("I can answer questions about this GST 2B analysis. Try one of these: summary, unmatched Purchase invoices, invoices missing in Tally, "
        "duplicate invoices, Needs Review, tax differences, high-value mismatches, party-wise differences, or search by invoice/GSTIN/party. "
        "I do not change reconciliation results or make accounting decisions.")


def format_rows(title, rows):
    if not rows: return f'{title}: none found in this analysis.'
    out=[f'{title}: {len(rows)} row(s).']
    for x in rows[:25]:
        inv=x.get('invoice_pur') or x.get('invoice_2b') or '—'
        diffs=[]
        for label,key in [('Taxable','taxable_diff'),('Invoice','invoice_value_diff'),('IGST','igst_diff'),('CGST','cgst_diff'),('SGST','sgst_diff')]:
            v=x.get(key)
            if v is not None and abs(v)>0: diffs.append(f'{label} {v:+,.2f}')
        out.append(f"- {x.get('name') or 'Unknown party'} | {inv} | {x.get('remark')} | GSTIN {x.get('gstin') or '—'} | {'; '.join(diffs) or 'No numeric difference'}")
    if len(rows)>25: out.append(f'…and {len(rows)-25} more. Use the report filters or Excel export for the complete list.')
    return '\n'.join(out)


def backup_payload():
    data=fetch_all_data()
    # Only history, analysis results, chat sessions/messages and settings are included.
    return {'format':'gst-smart-assistant-backup','version':1,'created_at':utc_now(),**data}


def restore_payload(payload):
    if not isinstance(payload,dict) or payload.get('format')!='gst-smart-assistant-backup' or payload.get('version')!=1:
        raise ValueError('This is not a supported GST Smart Assistant backup.')
    for key in ('analyses','sessions','messages','settings'):
        if not isinstance(payload.get(key,[]),list): raise ValueError(f'Backup section {key} is invalid.')
    if sum(len(payload.get(k,[])) for k in ('analyses','sessions','messages','settings'))>50000:
        raise ValueError('Backup contains too many records.')
    analyses=payload.get('analyses',[]); sessions=payload.get('sessions',[]); messages=payload.get('messages',[]); settings=payload.get('settings',[])
    if ensure_db():
        with db_connect() as conn:
            for a in analyses:
                conn.execute('''INSERT INTO gst_analyses(id,month_label,source_filename,created_at,report_json,settings_json)
                    VALUES(%s,%s,%s,COALESCE(%s::timestamptz,NOW()),%s::jsonb,%s::jsonb)
                    ON CONFLICT(id) DO UPDATE SET month_label=EXCLUDED.month_label,source_filename=EXCLUDED.source_filename,
                    report_json=EXCLUDED.report_json,settings_json=EXCLUDED.settings_json''',
                    (a['id'],a.get('month_label','Unknown period'),a.get('source_filename','Restored backup'),a.get('created_at'),json.dumps(a.get('report_json',{})),json.dumps(a.get('settings_json',{}))))
            for s in sessions:
                conn.execute('''INSERT INTO gst_chat_sessions(id,analysis_id,title,created_at,updated_at)
                    VALUES(%s,%s,%s,COALESCE(%s::timestamptz,NOW()),COALESCE(%s::timestamptz,NOW()))
                    ON CONFLICT(id) DO UPDATE SET analysis_id=EXCLUDED.analysis_id,title=EXCLUDED.title,updated_at=EXCLUDED.updated_at''',
                    (s['id'],s.get('analysis_id'),s.get('title','Restored chat'),s.get('created_at'),s.get('updated_at')))
            for m in messages:
                if any(s.get('id')==m.get('session_id') for s in sessions):
                    conn.execute('''INSERT INTO gst_chat_messages(id,session_id,role,content,created_at)
                        VALUES(%s,%s,%s,%s,COALESCE(%s::timestamptz,NOW())) ON CONFLICT(id) DO NOTHING''',
                        (m['id'],m['session_id'],m.get('role','assistant'),m.get('content',''),m.get('created_at')))
            for st in settings:
                conn.execute('''INSERT INTO gst_settings(key,value_json) VALUES(%s,%s::jsonb)
                    ON CONFLICT(key) DO UPDATE SET value_json=EXCLUDED.value_json,updated_at=NOW()''',
                    (st['key'],json.dumps(st.get('value_json'))))
        fetch_all_data()
    else:
        existing=load_local_data()
        for key, incoming in [('analyses',analyses),('sessions',sessions),('messages',messages),('settings',settings)]:
            by_id={str(x.get('id',x.get('key'))):x for x in existing.get(key,[])}
            for x in incoming: by_id[str(x.get('id',x.get('key')))]=x
            existing[key]=list(by_id.values())
        save_local_data(existing)


@app.after_request
def security_headers(response):
    response.headers['X-Content-Type-Options']='nosniff'
    response.headers['X-Frame-Options']='DENY'
    response.headers['Referrer-Policy']='no-referrer'
    response.headers['Cache-Control']='no-store, max-age=0'
    response.headers['Content-Security-Policy']="default-src 'self'; style-src 'self'; script-src 'self'; img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'self'; frame-ancestors 'none'; form-action 'self'"
    if request.is_secure or os.environ.get('RENDER'): response.headers['Strict-Transport-Security']='max-age=31536000; includeSubDomains'
    return response


@app.route('/health')
def health():
    return {'status':'ok','database_configured':bool(DATABASE_URL)},200


@app.route('/login',methods=['GET','POST'])
def login():
    if not APP_PASSWORD: return render_template('setup_required.html'),503
    if request.method=='POST':
        user=request.form.get('username',''); password=request.form.get('password','')
        if hmac.compare_digest(user.encode(),APP_USERNAME.encode()) and hmac.compare_digest(password.encode(),APP_PASSWORD.encode()):
            session.clear(); session['authenticated']=True; session.permanent=True
            return redirect(url_for('index'))
        flash('Login failed. Check your credentials.','error')
    return render_template('login.html')


@app.route('/logout',methods=['POST'])
@login_required
def logout():
    session.clear(); return redirect(url_for('login'))


@app.route('/',methods=['GET'])
@login_required
def index():
    try: data=fetch_all_data(); db_error=''
    except Exception as exc:
        logger.warning('History retrieval failed (%s).',type(exc).__name__); data=load_local_data(); db_error='Neon is temporarily unavailable; showing the local mirror.'
    analyses=sorted(data.get('analyses',[]),key=lambda x:str(x.get('created_at','')),reverse=True)
    return render_template('index.html',analyses=analyses[:30],db_enabled=bool(DATABASE_URL),db_error=db_error)


@app.route('/gst-reconcile',methods=['GET'])
@login_required
def gst_reconcile():
    return render_template('gst_reconciliation.html')


@app.route('/gst-reconcile/run',methods=['POST'])
@login_required
def run_gst_reconciliation():
    upload=request.files.get('workbook')
    if not upload or not upload.filename:
        flash('Choose one Excel workbook containing the Pur and 2B tables.','error'); return redirect(url_for('gst_reconcile'))
    filename=secure_filename(upload.filename)
    if not filename.lower().endswith('.xlsx'):
        flash('Please upload an .xlsx workbook.','error'); return redirect(url_for('gst_reconcile'))
    try:
        data=upload.read(MAX_FILE_BYTES+1)
        if not data: raise ValueError('The uploaded workbook is empty.')
        if len(data)>MAX_FILE_BYTES: raise ValueError('The workbook exceeds the 25 MB limit.')
        if not data.startswith(b'PK\x03\x04'): raise ValueError('The uploaded file is not a valid .xlsx workbook.')
        wb=load_workbook(io.BytesIO(data),data_only=True)
        try:
            pur_rows=read_named_table(wb,'Pur'); portal_rows=read_named_table(wb,'2B')
            fuzzy_enabled=request.form.get('fuzzy_enabled')=='on'
            report=reconcile_gst_tables(pur_rows,portal_rows,tolerance=2.0,fuzzy_enabled=fuzzy_enabled)
            month_label=month_from_purchase_rows(pur_rows,portal_rows)
        finally: wb.close()
        analysis_id=uuid.uuid4().hex
        save_analysis(analysis_id,month_label,filename,report,{'fuzzy_enabled':fuzzy_enabled,'tolerance':2.0})
        return redirect(url_for('analysis_view',analysis_id=analysis_id))
    except ValueError as exc:
        flash(str(exc),'error'); return redirect(url_for('gst_reconcile'))
    except Exception as exc:
        logger.error('GST reconciliation failed (%s).',type(exc).__name__)
        flash('GST reconciliation failed. Check table names, required columns and Neon connection, then try again.','error')
        return redirect(url_for('gst_reconcile'))


@app.route('/analysis/<analysis_id>')
@login_required
def analysis_view(analysis_id):
    entry=get_analysis(analysis_id)
    if not entry: abort(404)
    report=entry.get('report_json') or {}
    lines=[]
    for item in report.get('reconciliation',[]):
        row=dict(item); row['explanation'],row['action']=gst_explanation(row); lines.append(row)
    report['month_label']=entry.get('month_label','')
    counts=dict(report.get('counts',{})); counts['Matched']=counts.get('Matched',0)+counts.get('Fuzzy Matched',0)
    counts['All Discrepancies']=sum(1 for x in lines if x.get('remark') not in ('Matched','Fuzzy Matched'))
    counts['Validation warnings']=len(report.get('validation',[])); counts['Duplicate / mismatch review entries']=len(report.get('duplicates',[]))
    return render_template('gst_results.html',report=report,lines=lines,categories=sorted({x.get('remark','') for x in lines}),summary_counts=counts,analysis_id=analysis_id,month_label=entry.get('month_label',''))


@app.route('/analysis/<analysis_id>/download')
@login_required
def download_analysis(analysis_id):
    entry=get_analysis(analysis_id)
    if not entry: abort(404)
    report=entry.get('report_json') or {}
    output=build_gst_workbook(report)
    return send_file(io.BytesIO(output),mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name=month_filename(entry.get('month_label','')))


@app.route('/analysis/<analysis_id>/chat-export')
@login_required
def export_analysis_chat(analysis_id):
    entry=get_analysis(analysis_id)
    if not entry: abort(404)
    sessions=[s for s in fetch_all_data().get('sessions',[]) if s.get('analysis_id')==analysis_id]
    messages=[]
    for s in sessions:
        messages.extend(get_messages(s['id']))
    messages=sorted(messages,key=lambda x:str(x.get('created_at','')))
    body=''.join(f"<section><b>{html.escape(m.get('role','').title())}</b><p>{html.escape(m.get('content','')).replace(chr(10),'<br>')}</p></section>" for m in messages)
    page=f"<!doctype html><html><head><meta charset='utf-8'><title>GST 2B Chat</title><style>body{{font:15px Arial;max-width:900px;margin:40px auto;padding:0 18px;color:#172033}}section{{padding:12px;border-bottom:1px solid #ddd}}p{{white-space:normal;line-height:1.5}}</style></head><body><h1>GST 2B Reconciliation Chat - {html.escape(entry.get('month_label',''))}</h1>{body}</body></html>"
    return send_file(io.BytesIO(page.encode()),mimetype='text/html',as_attachment=True,download_name=month_chat_filename(entry.get('month_label','')))


@app.route('/assistant/<analysis_id>')
@login_required
def assistant_home(analysis_id):
    entry=get_analysis(analysis_id)
    if not entry: abort(404)
    sid=request.args.get('session_id')
    chat_session=get_session(sid) if sid else None
    if not chat_session or chat_session.get('analysis_id')!=analysis_id:
        sid=create_chat_session(analysis_id,f'GST 2B - {entry.get("month_label","")}')
        return redirect(url_for('assistant_home',analysis_id=analysis_id,session_id=sid))
    data=fetch_all_data(); sessions=[s for s in data.get('sessions',[]) if s.get('analysis_id')==analysis_id]
    sessions=sorted(sessions,key=lambda x:str(x.get('updated_at','')),reverse=True)
    return render_template('assistant.html',analysis=entry,analysis_id=analysis_id,chat_session=chat_session,sessions=sessions,messages=get_messages(sid),csrf_token_value=generate_csrf())


@app.route('/assistant/<analysis_id>/new',methods=['POST'])
@login_required
def assistant_new(analysis_id):
    entry=get_analysis(analysis_id)
    if not entry: abort(404)
    sid=create_chat_session(analysis_id,f'GST 2B - {entry.get("month_label","")}')
    return redirect(url_for('assistant_home',analysis_id=analysis_id,session_id=sid))


@app.route('/assistant/session/<session_id>/rename',methods=['POST'])
@login_required
def assistant_rename(session_id):
    s=get_session(session_id)
    if not s: abort(404)
    rename_session(session_id,request.form.get('title',''))
    return redirect(url_for('assistant_home',analysis_id=s.get('analysis_id'),session_id=session_id))


@app.route('/assistant/session/<session_id>/delete',methods=['POST'])
@login_required
def assistant_delete(session_id):
    s=get_session(session_id)
    if not s: abort(404)
    aid=s.get('analysis_id'); delete_session(session_id)
    return redirect(url_for('analysis_view',analysis_id=aid))


@app.route('/api/chat/<session_id>',methods=['POST'])
@login_required
def api_chat(session_id):
    s=get_session(session_id)
    if not s: abort(404)
    payload=request.get_json(silent=True) or request.form
    question=(payload.get('message') or '').strip()[:2000]
    entry=get_analysis(s.get('analysis_id'))
    if not entry: return jsonify({'error':'Analysis not found.'}),404
    add_message(session_id,'user',question)
    report=entry.get('report_json') or {}; report['month_label']=entry.get('month_label','')
    answer=assistant_answer(question,report,s.get('analysis_id'))
    add_message(session_id,'assistant',answer)
    return jsonify({'answer':answer,'messages':get_messages(session_id)})


@app.route('/history')
@login_required
def history():
    data=fetch_all_data()
    analyses=sorted(data.get('analyses',[]),key=lambda x:str(x.get('created_at','')),reverse=True)
    sessions=sorted(data.get('sessions',[]),key=lambda x:str(x.get('updated_at','')),reverse=True)
    return render_template('history.html',analyses=analyses,sessions=sessions)


@app.route('/backup')
@login_required
def backup_download():
    payload=backup_payload(); raw=json.dumps(payload,ensure_ascii=False,default=_json_default,indent=2).encode('utf-8')
    buffer=io.BytesIO()
    with zipfile.ZipFile(buffer,'w',zipfile.ZIP_DEFLATED) as zf: zf.writestr('gst-smart-assistant-backup.json',raw)
    buffer.seek(0)
    return send_file(buffer,mimetype='application/zip',as_attachment=True,download_name='GST_Smart_Assistant_Backup.zip')


@app.route('/restore',methods=['POST'])
@login_required
def backup_restore():
    upload=request.files.get('backup')
    if not upload or not upload.filename:
        flash('Choose a GST Smart Assistant backup ZIP or JSON file.','error'); return redirect(url_for('history'))
    try:
        raw=upload.read(BACKUP_MAX_BYTES+1)
        if len(raw)>BACKUP_MAX_BYTES: raise ValueError('Backup exceeds the 25 MB limit.')
        if upload.filename.lower().endswith('.zip'):
            with zipfile.ZipFile(io.BytesIO(raw)) as zf:
                names=[n for n in zf.namelist() if n.endswith('.json') and '/' not in n and not n.startswith('.')]
                if not names: raise ValueError('No backup JSON was found inside the ZIP.')
                raw=zf.read(names[0])
        payload=json.loads(raw.decode('utf-8-sig'))
        restore_payload(payload)
        flash('Backup restored. Existing matching IDs were updated; new records were added.','success')
    except Exception as exc:
        logger.warning('Backup restore failed (%s).',type(exc).__name__)
        flash(str(exc) if isinstance(exc,ValueError) else 'Backup could not be restored. Verify the backup file.','error')
    return redirect(url_for('history'))


@app.errorhandler(RequestEntityTooLarge)
def request_too_large(_error):
    flash('Upload is too large.','error'); return redirect(url_for('gst_reconcile'))


@app.errorhandler(413)
def too_large(_error):
    return {'error':'Request too large.'},413




def gst_explanation(line):
    """Create plain-language explanation and suggested action for a reconciliation line."""
    remark = line.get("remark", "")
    details = []
    actions = []
    if remark in ("Matched", "Fuzzy Matched"):
        if remark == "Fuzzy Matched":
            return (line.get("detail") or "Invoice pair was matched using fuzzy invoice-number comparison and passed the configured date and amount tolerances.",
                    "Retain the source invoice and review the fuzzy-match audit details in the report.")
        return ("The invoice is present in both records and the compared amounts are within the configured tolerance.",
                "No action normally required; retain supporting records.")
    if remark == "Old Bill Filed":
        return (line.get("detail") or "The 2B invoice date is earlier than the earliest valid Purchase date.",
                "Check the invoice date, accounting period, and whether the bill was recorded in a later period.")
    if remark == "Not in Tally":
        return ("The invoice appears in 2B, but no corresponding invoice was found in the Purchase table.",
                "Search the Purchase register by supplier GSTIN and invoice number; if genuinely absent, review the purchase entry.")
    if remark == "Not on Portal":
        return ("The invoice appears in Purchase, but no corresponding invoice was found in 2B.",
                "Verify the supplier's reporting, invoice number, GSTIN, and relevant return period.")
    if "GSTIN Mismatch" in remark:
        return ("The invoice number appears on both sides, but the GSTIN differs. It has not been automatically matched.",
                "Compare the GSTIN on the original supplier invoice and verify the correct supplier identity.")
    if "Duplicate" in remark or "Review" in remark:
        return (line.get("detail") or "Multiple or ambiguous records prevent a reliable automatic match.",
                "Review each source row and confirm whether the records are duplicates or legitimate separate entries.")
    if remark == "Date Difference":
        return ("The invoice was matched, but the dates differ between Purchase and 2B.",
                "Check the original invoice date and confirm whether the date difference is expected.")
    if remark == "Not Matched":
        fields = []
        for label, key in [("Taxable value", "taxable_diff"), ("Invoice value", "invoice_value_diff"),
                           ("IGST", "igst_diff"), ("CGST", "cgst_diff"), ("SGST", "sgst_diff")]:
            value = line.get(key)
            if value is not None and abs(value) > line.get("tolerance", 2.0):
                fields.append(f"{label} differs by ₹{value:,.2f} (2B minus Pur)")
        if fields:
            details.append("; ".join(fields))
        return (("Matched invoice, but " + "; ".join(details) + ".") if details else "One or more compared fields differ.",
                "Compare the supplier invoice with both records and correct the source that is inaccurate.")
    return (line.get("detail") or f"Review required: {remark}.",
            "Inspect the source documents and resolve the discrepancy before relying on the reconciliation.")

def read_named_table(workbook, table_name):
    """Read an Excel Table by its defined table name, not by worksheet name."""
    target = table_name.casefold()
    accepted_names = {target}
    if target == "2b":
        # Excel table names cannot begin with a digit in many Excel versions; accept practical aliases.
        accepted_names.update({"table2b", "t2b", "portal2b"})
    for ws in workbook.worksheets:
        for name in ws.tables:
            if name.casefold() in accepted_names:
                table = ws.tables[name]
                cells = ws[table.ref]
                values = [[cell.value for cell in row] for row in cells]
                if len(values) < 2:
                    raise ValueError(f"The Excel table '{table_name}' has headers but no data rows.")
                headers = [str(v).strip() if v is not None else "" for v in values[0]]
                normalized = [normalize_header(h) for h in headers]
                missing = [label for label, aliases in REQUIRED_GST_COLUMNS.items() if not any(h in aliases for h in normalized)]
                if missing:
                    raise ValueError(f"Table '{table_name}' is missing required columns: {', '.join(missing)}.")
                records = []
                for excel_row, values_row in enumerate(values[1:], start=2):
                    if not any(v is not None and str(v).strip() for v in values_row):
                        continue
                    record = {normalize_header(headers[i]): values_row[i] if i < len(values_row) else None for i in range(len(headers))}
                    record["_source_row"] = excel_row
                    record["_table"] = table_name
                    records.append(record)
                return records
    raise ValueError(f"Could not find an Excel Table named '{table_name}'. Check Table Design → Table Name.")

REQUIRED_GST_COLUMNS = {
    "DATE": {"date", "invoice date", "bill date"},
    "Particulars": {"particulars", "party", "party name", "supplier name", "name"},
    "Supplier Invoice No.": {"supplier invoice no", "supplier invoice number", "invoice no", "invoice number", "inv no", "bill no"},
    "GSTIN/UIN": {"gstin uin", "gstin", "gstin number", "supplier gstin", "uin"},
    "TAXABLE VALUE": {"taxable value", "taxable amount", "taxable"},
    "INV VALUE": {"inv value", "invoice value", "invoice amount", "total invoice value"},
    "IGST": {"igst", "igst amount"},
    "CGST": {"cgst", "cgst amount"},
    "SGST": {"sgst", "sgst amount"},
}

def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).strip()

def gst_field(record, label):
    aliases = REQUIRED_GST_COLUMNS[label]
    for key, value in record.items():
        if key in aliases:
            return value
    return None

def gst_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ")).strip()

def normalize_gstin(value):
    return re.sub(r"\s+", "", gst_text(value)).upper()

def normalize_invoice(value):
    # Conservative normalization: ignore case and whitespace only; preserve punctuation and leading zeros.
    return re.sub(r"\s+", "", gst_text(value)).upper()

def gst_number(value):
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("₹", "")
    try:
        return float(text)
    except ValueError:
        return None

def gst_date(value):
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, (int, float)):
        # Excel's default 1900 date system. openpyxl normally converts date-formatted cells already.
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).strftime("%d-%m-%Y")
        except (OverflowError, ValueError):
            return str(value)
    text = str(value).strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d-%m-%Y")
        except ValueError:
            pass
    return text

def parse_gst_date(value):
    """Parse normalized DD-MM-YYYY dates for chronological comparison."""
    if not value:
        return None
    try:
        return datetime.strptime(str(value).strip(), "%d-%m-%Y").date()
    except (TypeError, ValueError):
        return None

def gst_record_view(record):
    return {
        "date": gst_date(gst_field(record, "DATE")),
        "name": gst_text(gst_field(record, "Particulars")),
        "invoice": gst_text(gst_field(record, "Supplier Invoice No.")),
        "gstin": normalize_gstin(gst_field(record, "GSTIN/UIN")),
        "taxable": gst_number(gst_field(record, "TAXABLE VALUE")),
        "invoice_value": gst_number(gst_field(record, "INV VALUE")),
        "igst": gst_number(gst_field(record, "IGST")),
        "cgst": gst_number(gst_field(record, "CGST")),
        "sgst": gst_number(gst_field(record, "SGST")),
        "source_row": record.get("_source_row", ""),
    }

def consolidate_gst_duplicates(rows, side):
    """Combine rows sharing GSTIN + normalized invoice number + invoice date.

    Original source row numbers are retained for audit. Rows missing GSTIN or
    invoice number are deliberately kept separate and are never consolidated.
    """
    grouped = defaultdict(list)
    for row in rows:
        gstin = normalize_gstin(row.get("gstin"))
        invoice = normalize_invoice(row.get("invoice"))
        if gstin and invoice and row.get("date"):
            grouped[(gstin, invoice, row["date"])].append(row)
        else:
            # Missing key/date information is insufficient to safely consolidate.
            grouped[("__single__", side, len(grouped), row.get("source_row"))].append(row)

    consolidated, audit = [], []
    amount_fields = ("taxable", "invoice_value", "igst", "cgst", "sgst")
    for group_key, members in grouped.items():
        if group_key[0] == "__single__" or len(members) == 1:
            consolidated.append(members[0])
            continue
        combined = dict(members[0])
        for field in amount_fields:
            values = [member.get(field) for member in members if member.get(field) is not None]
            combined[field] = sum(values) if values else None
        source_rows = [str(member.get("source_row", "")) for member in members if member.get("source_row") not in (None, "")]
        combined["source_row"] = ", ".join(source_rows)
        combined["source_rows"] = source_rows
        combined["_consolidated_count"] = len(members)
        consolidated.append(combined)
        audit.append([group_key[0], group_key[1], side, len(members),
                      f"Consolidated same-date duplicate rows; source Excel rows: {', '.join(source_rows)}"])
    return consolidated, audit

def reconcile_gst_tables(pur_records, portal_records, tolerance=2.0, fuzzy_enabled=False):
    pur_raw_count, portal_raw_count = len(pur_records), len(portal_records)
    pur_raw = [gst_record_view(r) for r in pur_records]
    portal_raw = [gst_record_view(r) for r in portal_records]
    pur, pur_consolidation_log = consolidate_gst_duplicates(pur_raw, "Pur")
    portal, portal_consolidation_log = consolidate_gst_duplicates(portal_raw, "2B")
    def key(row):
        inv = normalize_invoice(row["invoice"])
        gst = normalize_gstin(row["gstin"])
        return (gst, inv) if gst and inv else None
    pur_index, portal_index = defaultdict(list), defaultdict(list)
    for i, row in enumerate(pur):
        k = key(row)
        if k: pur_index[k].append(i)
    for i, row in enumerate(portal):
        k = key(row)
        if k: portal_index[k].append(i)
    pur_used, portal_used = set(), set()
    reconciliation, duplicates, validation = [], pur_consolidation_log + portal_consolidation_log, []

    # All records involved in duplicated GSTIN+invoice keys are sent to review; never guess a pairing.
    duplicate_keys = {k for k in set(pur_index) | set(portal_index) if len(pur_index.get(k, [])) > 1 or len(portal_index.get(k, [])) > 1}
    for k in sorted(duplicate_keys):
        for side, index, rows in (("Pur", pur_index, pur), ("2B", portal_index, portal)):
            if len(index.get(k, [])) > 1:
                duplicates.append([k[0], k[1], side, len(index[k]), "Duplicate GSTIN + invoice key; manual review required"] )
        for i in pur_index.get(k, []): pur_used.add(i)
        for i in portal_index.get(k, []): portal_used.add(i)
        # Show each source record as an individual review line.
        for i in pur_index.get(k, []):
            reconciliation.append(make_gst_line(pur[i], None, "Duplicate / Review Required", tolerance, "Duplicate key in Pur"))
        for i in portal_index.get(k, []):
            reconciliation.append(make_gst_line(None, portal[i], "Duplicate / Review Required", tolerance, "Duplicate key in 2B"))

    unique_common = (set(pur_index) & set(portal_index)) - duplicate_keys
    for k in sorted(unique_common):
        i, j = pur_index[k][0], portal_index[k][0]
        pur_used.add(i); portal_used.add(j)
        consolidation_notes = []
        if pur[i].get("_consolidated_count", 1) > 1:
            consolidation_notes.append(f"Purchase: {pur[i]['_consolidated_count']} rows consolidated (source rows {pur[i].get('source_row', '')})")
        if portal[j].get("_consolidated_count", 1) > 1:
            consolidation_notes.append(f"2B: {portal[j]['_consolidated_count']} rows consolidated (source rows {portal[j].get('source_row', '')})")
        line_detail = "; ".join(consolidation_notes)
        line = make_gst_line(pur[i], portal[j], "Matched", tolerance, line_detail)
        if line["taxable_diff_abs"] > tolerance or any(line[f"{tax.lower()}_diff_abs"] > tolerance for tax in ("IGST", "CGST", "SGST")) or line["invoice_value_diff_abs"] > tolerance:
            line["remark"] = "Not Matched"
        elif line["date_pur"] != line["date_2b"]:
            line["remark"] = "Date Difference"
        else:
            line["remark"] = "Matched"
        reconciliation.append(line)

    # Earliest valid date in Purchase register, used for Old Bill Filed classification.
    pur_dates = [parsed for row in pur if (parsed := parse_gst_date(row["date"])) is not None]
    min_pur_date = min(pur_dates) if pur_dates else None

    # Check invoice-only overlaps for possible GSTIN mismatch, but do not auto-match them.
    remaining_pur = [i for i in range(len(pur)) if i not in pur_used]
    remaining_portal = [j for j in range(len(portal)) if j not in portal_used]
    pur_inv, portal_inv = defaultdict(list), defaultdict(list)
    for i in remaining_pur:
        inv = normalize_invoice(pur[i]["invoice"])
        if inv: pur_inv[inv].append(i)
    for j in remaining_portal:
        inv = normalize_invoice(portal[j]["invoice"])
        if inv: portal_inv[inv].append(j)
    possible_mismatch_pur, possible_mismatch_portal = set(), set()
    for inv in set(pur_inv) & set(portal_inv):
        for i in pur_inv[inv]:
            for j in portal_inv[inv]:
                if pur[i]["gstin"] != portal[j]["gstin"]:
                    possible_mismatch_pur.add(i); possible_mismatch_portal.add(j)
                    duplicates.append([f"{pur[i]['gstin']} / {portal[j]['gstin']}", inv, "Pur + 2B", 2, "Invoice number matches but GSTIN differs; verify supplier and invoice"] )
    fuzzy_suggestions = []
    fuzzy_review_pur, fuzzy_review_portal = {}, {}
    fuzzy_matched_pairs = []
    if fuzzy_enabled:
        # Fuzzy second pass: same GSTIN, >=75% invoice similarity, dates within 3 days.
        # Automatic Fuzzy Matched status additionally requires taxable <=₹10 and
        # the combined tax amount (IGST + CGST + SGST) to match exactly.
        eligible_pur = [i for i in remaining_pur if i not in possible_mismatch_pur]
        eligible_portal = [j for j in remaining_portal if j not in possible_mismatch_portal]
        candidates = []
        amount_fields = ("taxable", "igst", "cgst", "sgst")
        for i in eligible_pur:
            left = pur[i]
            left_inv = re.sub(r"[^A-Z0-9]", "", normalize_invoice(left["invoice"]))
            left_gstin = normalize_gstin(left["gstin"])
            left_date = parse_gst_date(left["date"])
            if not left_inv or not left_gstin or not left_date:
                continue
            for j in eligible_portal:
                right = portal[j]
                if left_gstin != normalize_gstin(right["gstin"]):
                    continue
                right_inv = re.sub(r"[^A-Z0-9]", "", normalize_invoice(right["invoice"]))
                right_date = parse_gst_date(right["date"])
                if not right_inv or not right_date:
                    continue
                date_gap = abs((left_date - right_date).days)
                if date_gap > 3:
                    continue
                similarity = SequenceMatcher(None, left_inv, right_inv).ratio() * 100
                if similarity < 75:
                    continue

                # Retain the agreed taxable-value tolerance of ₹10.
                # Tax is compared as one combined amount: IGST + CGST + SGST.
                missing_amounts = []
                failed_amounts = []
                taxable_left, taxable_right = left.get("taxable"), right.get("taxable")
                if taxable_left is None or taxable_right is None:
                    missing_amounts.append("taxable")
                    taxable_difference = None
                else:
                    taxable_difference = abs(taxable_right - taxable_left)
                    if taxable_difference > 10.0:
                        failed_amounts.append(("taxable", taxable_difference, 10.0))

                invoice_value_left, invoice_value_right = left.get("invoice_value"), right.get("invoice_value")
                if invoice_value_left is None or invoice_value_right is None:
                    missing_amounts.append("invoice value")
                elif abs(invoice_value_right - invoice_value_left) > 10.0:
                    failed_amounts.append(("invoice value", abs(invoice_value_right - invoice_value_left), 10.0))
                left_tax_parts = [left.get(field) for field in ("igst", "cgst", "sgst")]
                right_tax_parts = [right.get(field) for field in ("igst", "cgst", "sgst")]
                # In GST registers, blank tax-component cells commonly mean zero
                # when another component is populated (e.g. IGST invoice with
                # blank CGST/SGST). Treat those blanks as zero, but keep the total
                # tax as missing when all three components are blank on either side.
                if all(value is None for value in left_tax_parts) or all(value is None for value in right_tax_parts):
                    missing_amounts.append("total tax")
                    total_tax_left = total_tax_right = None
                    total_tax_difference = None
                else:
                    total_tax_left = sum(value or 0.0 for value in left_tax_parts)
                    total_tax_right = sum(value or 0.0 for value in right_tax_parts)
                    total_tax_difference = abs(total_tax_right - total_tax_left)
                    if round(total_tax_difference, 2) > 2.00:
                        failed_amounts.append(("total tax", total_tax_difference, 2.0))

                amounts_pass = not missing_amounts and not failed_amounts
                amount_score = 100.0 if amounts_pass else 0.0
                date_score = max(0.0, 100.0 - (date_gap / 3.0 * 100.0))
                confidence = round(similarity * 0.70 + amount_score * 0.20 + date_score * 0.10, 1)
                taxable_diff = (right["taxable"] - left["taxable"]) if left["taxable"] is not None and right["taxable"] is not None else None
                invoice_value_diff = (right["invoice_value"] - left["invoice_value"]) if left["invoice_value"] is not None and right["invoice_value"] is not None else None
                tax_detail = ", ".join(f"{field.upper()} ₹{diff:,.2f} (limit ₹{limit:,.2f})" for field, diff, limit in failed_amounts)
                missing_detail = ", ".join(x.upper() for x in missing_amounts)
                amount_note = ("Taxable difference is within ₹10 and combined tax (IGST + CGST + SGST) difference is within ₹2" if amounts_pass else "; ".join(x for x in [f"Over tolerance / mismatch: {tax_detail}" if tax_detail else "", f"Missing amounts: {missing_detail}" if missing_detail else ""] if x))
                candidates.append({
                    "i": i, "j": j, "gstin": left_gstin, "name": left["name"] or right["name"],
                    "invoice_pur": left["invoice"], "invoice_2b": right["invoice"],
                    "date_pur": left["date"], "date_2b": right["date"],
                    "invoice_similarity": round(similarity, 1), "confidence": confidence,
                    "date_gap": date_gap, "amounts_agree": amounts_pass,
                    "total_tax_pur": total_tax_left, "total_tax_2b": total_tax_right,
                    "total_tax_diff": (total_tax_right - total_tax_left) if total_tax_left is not None and total_tax_right is not None else None,
                    "taxable_diff": taxable_diff, "invoice_value_diff": invoice_value_diff,
                    "igst_diff": (right["igst"] - left["igst"]) if left["igst"] is not None and right["igst"] is not None else None,
                    "cgst_diff": (right["cgst"] - left["cgst"]) if left["cgst"] is not None and right["cgst"] is not None else None,
                    "sgst_diff": (right["sgst"] - left["sgst"]) if left["sgst"] is not None and right["sgst"] is not None else None,
                    "amount_note": amount_note,
                    "purchase_source_row": left["source_row"], "portal_source_row": right["source_row"],
                })

        # Only auto-match a pair when it is the unique best candidate from BOTH sides.
        by_pur, by_portal = defaultdict(list), defaultdict(list)
        for candidate in candidates:
            by_pur[candidate["i"]].append(candidate)
            by_portal[candidate["j"]].append(candidate)
        def unique_best(items):
            ordered = sorted(items, key=lambda x: (x["confidence"], x["invoice_similarity"], -x["date_gap"]), reverse=True)
            if not ordered:
                return None
            if len(ordered) > 1 and (ordered[0]["confidence"], ordered[0]["invoice_similarity"], ordered[0]["date_gap"]) == (ordered[1]["confidence"], ordered[1]["invoice_similarity"], ordered[1]["date_gap"]):
                return None
            return ordered[0]
        auto_pairs = []
        for candidate in candidates:
            if candidate["invoice_similarity"] < 85 or not candidate["amounts_agree"]:
                continue
            best_p = unique_best(by_pur[candidate["i"]])
            best_2b = unique_best(by_portal[candidate["j"]])
            if best_p is candidate and best_2b is candidate:
                auto_pairs.append(candidate)
        # Guard one-to-one even in unusual tie/duplicate situations.
        claimed_p, claimed_2b = set(), set()
        auto_ids = set()
        for candidate in sorted(auto_pairs, key=lambda x: (-x["confidence"], -x["invoice_similarity"], x["date_gap"])):
            if candidate["i"] in claimed_p or candidate["j"] in claimed_2b:
                continue
            claimed_p.add(candidate["i"]); claimed_2b.add(candidate["j"])
            auto_ids.add((candidate["i"], candidate["j"]))
            fuzzy_matched_pairs.append(candidate)
            line = make_gst_line(pur[candidate["i"]], portal[candidate["j"]], "Fuzzy Matched", tolerance,
                f"Fuzzy match accepted: invoice numbers {candidate['invoice_similarity']:.1f}% similar; same GSTIN; dates {candidate['date_gap']} day(s) apart (maximum 3); taxable difference within ₹10; combined tax (IGST + CGST + SGST) difference is within ₹2. Manual verification is still recommended.")
            line["invoice_similarity"] = candidate["invoice_similarity"]
            line["confidence"] = candidate["confidence"]
            reconciliation.append(line)
            pur_used.add(candidate["i"]); portal_used.add(candidate["j"])

        # Keep every candidate visible. Candidates not safely auto-matched remain review-only.
        for candidate in candidates:
            pair = (candidate["i"], candidate["j"])
            auto_matched = pair in auto_ids
            if auto_matched:
                status = "Fuzzy Matched"
                explanation = f"Invoice numbers are {candidate['invoice_similarity']:.1f}% similar; same GSTIN; dates differ by {candidate['date_gap']} day(s). {candidate['amount_note']}."
                action = "Fuzzy matched by configured rules. Retain invoice evidence and verify the pair during review."
            else:
                status = "Review Required"
                explanation = f"Invoice numbers are {candidate['invoice_similarity']:.1f}% similar; same GSTIN; dates differ by {candidate['date_gap']} day(s). {candidate['amount_note'] or 'Candidate is ambiguous because another invoice is similarly ranked.'}"
                if candidate["amounts_agree"]:
                    explanation += " Candidate was not auto-matched because the best pairing is ambiguous."
                action = "Do not treat as matched. Compare original invoice and choose the correct one-to-one pair manually."
                if candidate["i"] not in claimed_p:
                    fuzzy_review_pur[candidate["i"]] = explanation
                if candidate["j"] not in claimed_2b:
                    fuzzy_review_portal[candidate["j"]] = explanation
            fuzzy_suggestions.append({**{k:v for k,v in candidate.items() if k not in ("i", "j", "amount_note")},
                "status": status, "explanation": explanation, "action": action})
        fuzzy_suggestions.sort(key=lambda x: (x["status"] != "Fuzzy Matched", -x["confidence"], x["date_gap"]))

    for i in remaining_pur:
        if i in pur_used:
            continue
        if i in possible_mismatch_pur:
            reconciliation.append(make_gst_line(pur[i], None, "GSTIN Mismatch - Review", tolerance, "Same invoice number found in 2B with a different GSTIN"))
            pur_used.add(i)
        elif i in fuzzy_review_pur:
            reconciliation.append(make_gst_line(pur[i], None, "Fuzzy Match - Review", tolerance, fuzzy_review_pur[i]))
            pur_used.add(i)
        else:
            reconciliation.append(make_gst_line(pur[i], None, "Not on Portal", tolerance, "No corresponding invoice found in 2B"))
            pur_used.add(i)
    for j in remaining_portal:
        if j in portal_used:
            continue
        if j in possible_mismatch_portal:
            reconciliation.append(make_gst_line(None, portal[j], "GSTIN Mismatch - Review", tolerance, "Same invoice number found in Pur with a different GSTIN"))
            portal_used.add(j)
        elif j in fuzzy_review_portal:
            reconciliation.append(make_gst_line(None, portal[j], "Fuzzy Match - Review", tolerance, fuzzy_review_portal[j]))
            portal_used.add(j)
        else:
            portal_date = parse_gst_date(portal[j]["date"])
            if min_pur_date and portal_date and portal_date < min_pur_date:
                reconciliation.append(make_gst_line(
                    None, portal[j], "Old Bill Filed", tolerance,
                    f"Invoice date {portal[j]['date']} is earlier than the earliest Purchase date {min_pur_date.strftime('%d-%m-%Y')}. Check whether this is an old bill filed in a later period."
                ))
            else:
                reconciliation.append(make_gst_line(None, portal[j], "Not in Tally", tolerance, "No corresponding invoice found in Pur"))
            portal_used.add(j)

    for side, rows in (("Pur", pur), ("2B", portal)):
        for n, row in enumerate(rows, start=1):
            if not row["invoice"]: validation.append([side, row["source_row"], row["gstin"], "", "Supplier Invoice No. is blank"] )
            if not row["gstin"]: validation.append([side, row["source_row"], row["gstin"], row["invoice"], "GSTIN/UIN is blank; reliable matching may not be possible"] )
            elif len(row["gstin"]) != 15: validation.append([side, row["source_row"], row["gstin"], row["invoice"], f"GSTIN length is {len(row['gstin'])}; expected 15 characters. Verify manually."])
            if not row["date"]: validation.append([side, row["source_row"], row["gstin"], row["invoice"], "DATE is blank"] )
            for field in ("taxable", "invoice_value", "igst", "cgst", "sgst"):
                original = row[field]
                if original is None:
                    validation.append([side, row["source_row"], row["gstin"], row["invoice"], f"{field.replace('_', ' ').title()} is blank or not numeric"] )
            if all(row[x] is not None for x in ("taxable", "igst", "cgst", "sgst", "invoice_value")):
                expected = row["taxable"] + row["igst"] + row["cgst"] + row["sgst"]
                if abs(expected - row["invoice_value"]) > tolerance:
                    validation.append([side, row["source_row"], row["gstin"], row["invoice"], f"Taxable + IGST + CGST + SGST = {expected:.2f}; invoice value = {row['invoice_value']:.2f}; difference = {row['invoice_value'] - expected:+.2f}"])
    # Summary counts refer to reconciliation lines; duplicate records are review lines, not automatic matches.
    counts = defaultdict(int)
    for line in reconciliation: counts[line["remark"]] += 1
    return {"reconciliation": reconciliation, "duplicates": duplicates, "validation": validation, "fuzzy_suggestions": fuzzy_suggestions, "fuzzy_enabled": fuzzy_enabled, "fuzzy_matched_count": len(fuzzy_matched_pairs), "pur_count": pur_raw_count, "portal_count": portal_raw_count, "pur_consolidated_count": len(pur), "portal_consolidated_count": len(portal), "tolerance": tolerance, "counts": dict(counts)}

def make_gst_line(pur, portal, remark, tolerance, detail):
    def val(row, field): return row[field] if row and row[field] is not None else None
    def diff(field):
        # For a one-sided invoice, treat the missing source amount as zero so
        # the web report shows the full exposure instead of an em dash.
        a, b = val(pur, field), val(portal, field)
        if a is None and b is None:
            return None
        return (b or 0.0) - (a or 0.0)
    taxable_diff, inv_diff = diff("taxable"), diff("invoice_value")
    tax_diffs = {tax: diff(tax.lower()) for tax in ("IGST", "CGST", "SGST")}
    name = (pur["name"] if pur and pur["name"] else (portal["name"] if portal else ""))
    return {
        "gstin": val(pur, "gstin") or val(portal, "gstin") or "", "name": name,
        "invoice_pur": val(pur, "invoice") or "", "invoice_2b": val(portal, "invoice") or "",
        "date_pur": val(pur, "date") or "", "date_2b": val(portal, "date") or "",
        "taxable_pur": val(pur, "taxable"), "taxable_2b": val(portal, "taxable"), "taxable_diff": taxable_diff, "taxable_diff_abs": abs(taxable_diff) if taxable_diff is not None else 0,
        "invoice_value_pur": val(pur, "invoice_value"), "invoice_value_2b": val(portal, "invoice_value"), "invoice_value_diff": inv_diff, "invoice_value_diff_abs": abs(inv_diff) if inv_diff is not None else 0,
        "igst_pur": val(pur, "igst"), "igst_2b": val(portal, "igst"), "igst_diff": tax_diffs["IGST"], "igst_diff_abs": abs(tax_diffs["IGST"]) if tax_diffs["IGST"] is not None else 0,
        "cgst_pur": val(pur, "cgst"), "cgst_2b": val(portal, "cgst"), "cgst_diff": tax_diffs["CGST"], "cgst_diff_abs": abs(tax_diffs["CGST"]) if tax_diffs["CGST"] is not None else 0,
        "sgst_pur": val(pur, "sgst"), "sgst_2b": val(portal, "sgst"), "sgst_diff": tax_diffs["SGST"], "sgst_diff_abs": abs(tax_diffs["SGST"]) if tax_diffs["SGST"] is not None else 0,
        "remark": remark, "detail": detail, "tolerance": tolerance,
    }

def build_gst_workbook(report):
    wb = Workbook()
    ws = wb.active; ws.title = "Reconciliation"
    columns = [
        ("FinalGST", "gstin"), ("FinalName", "name"), ("Pur Invoice No.", "invoice_pur"), ("2B Invoice No.", "invoice_2b"),
        ("Pur Date", "date_pur"), ("2B Date", "date_2b"),
        ("2B TAXABLE VALUE", "taxable_2b"), ("2B INV VALUE", "invoice_value_2b"), ("2B IGST", "igst_2b"), ("2B CGST", "cgst_2b"), ("2B SGST", "sgst_2b"),
        ("PUR TAXABLE VALUE", "taxable_pur"), ("PUR INV VALUE", "invoice_value_pur"), ("PUR IGST", "igst_pur"), ("PUR CGST", "cgst_pur"), ("PUR SGST", "sgst_pur"),
        ("TAXABLE DIFF (2B - Pur)", "taxable_diff"), ("INV VALUE DIFF (2B - Pur)", "invoice_value_diff"), ("IGST DIFF (2B - Pur)", "igst_diff"), ("CGST DIFF (2B - Pur)", "cgst_diff"), ("SGST DIFF (2B - Pur)", "sgst_diff"), ("REMARKS", "remark"), ("DETAIL", "detail")
    ]
    write_sheet(ws, [c[0] for c in columns], [[line.get(c[1]) for c in columns] for line in report["reconciliation"]])
    missing_2b = [x for x in report["reconciliation"] if x["remark"] == "Not on Portal"]
    missing_pur = [x for x in report["reconciliation"] if x["remark"] == "Not in Tally"]
    old_bills = [x for x in report["reconciliation"] if x["remark"] == "Old Bill Filed"]
    tax_diffs = [x for x in report["reconciliation"] if any(x[k] is not None and x[k] > report["tolerance"] for k in ("taxable_diff_abs", "invoice_value_diff_abs", "igst_diff_abs", "cgst_diff_abs", "sgst_diff_abs"))]
    review = [x for x in report["reconciliation"] if "Review" in x["remark"] or x["remark"] == "Duplicate / Review Required"]
    all_discrepancies = [x for x in report["reconciliation"] if x["remark"] not in ("Matched", "Fuzzy Matched")]
    write_gst_subset(wb, "Missing in 2B", columns, missing_2b)
    write_gst_subset(wb, "Missing in Purchase", columns, missing_pur)
    write_gst_subset(wb, "Old Bill Filed", columns, old_bills)
    write_gst_subset(wb, "Tax Differences", columns, tax_diffs)
    write_gst_subset(wb, "Duplicates & Review", columns, review)
    write_gst_subset(wb, "All Discrepancies", columns, all_discrepancies)
    write_sheet(wb.create_sheet("Validation"), ["Source", "Excel Row in Source Table", "GSTIN", "Invoice No.", "Validation Warning"], report["validation"])
    write_sheet(wb.create_sheet("Duplicate Keys"), ["GSTIN / GSTINs", "Normalized Invoice No.", "Source", "Count", "Review Note"], report["duplicates"])
    fuzzy_columns = [
        ("GSTIN", "gstin"), ("Party", "name"), ("Purchase Invoice", "invoice_pur"), ("2B Invoice", "invoice_2b"),
        ("Purchase Date", "date_pur"), ("2B Date", "date_2b"), ("Invoice Similarity %", "invoice_similarity"),
        ("Confidence %", "confidence"), ("Date Gap (days)", "date_gap"), ("Total Tax - Purchase", "total_tax_pur"), ("Total Tax - 2B", "total_tax_2b"), ("Total Tax Difference", "total_tax_diff"), ("Amounts Pass (Taxable ₹10 / Combined Tax Difference ≤₹2)", "amounts_agree"), ("Status", "status"),
        ("Taxable Diff (2B - Pur)", "taxable_diff"), ("Invoice Value Diff (2B - Pur)", "invoice_value_diff"),
        ("IGST Diff (2B - Pur)", "igst_diff"), ("CGST Diff (2B - Pur)", "cgst_diff"), ("SGST Diff (2B - Pur)", "sgst_diff"),
        ("Explanation", "explanation"), ("Recommended Action", "action"),
        ("Purchase Source Row", "purchase_source_row"), ("2B Source Row", "portal_source_row")]
    write_gst_subset(wb, "Fuzzy Suggestions", fuzzy_columns, report.get("fuzzy_suggestions", []))
    summary = wb.create_sheet("Summary", 0)
    summary_rows = [["GST 2B and Purchase Reconciliation", "Count / Value"], ["Records in Pur", report["pur_count"]], ["Records in 2B", report["portal_count"]], ["Exact Matched", report["counts"].get("Matched", 0)], ["Fuzzy Matched", report["counts"].get("Fuzzy Matched", 0)], ["Total Matched (Exact + Fuzzy)", report["counts"].get("Matched", 0) + report["counts"].get("Fuzzy Matched", 0)], ["Exact-match amount tolerance (₹ per field)", report["tolerance"]]]
    for label in sorted(report["counts"]):
        summary_rows.append([label, report["counts"][label]])
    summary_rows.extend([["Old Bill Filed", len(old_bills)], ["All Discrepancies", len(all_discrepancies)],
                         ["Validation warnings", len(report["validation"])], ["Duplicate / mismatch review entries", len(report["duplicates"])],
                         ["Difference convention", "2B minus Pur; positive means 2B is higher"],
                         ["Matching rule", "Rows with same GSTIN + normalized invoice number + invoice date are consolidated separately in Pur and 2B before matching; same invoice key on different dates remains under review"], ["Fuzzy matching enabled", "Yes" if report.get("fuzzy_enabled") else "No"], ["Fuzzy suggestions", len(report.get("fuzzy_suggestions", []))], ["Fuzzy matching rule", "Same GSTIN, invoice similarity >=75%, dates within 3 days; auto-match only when taxable difference <=₹10, combined IGST+CGST+SGST difference is within ₹2, and pairing is unambiguous; otherwise review"]])
    for row in summary_rows:
        summary.append(row)
    style_sheet(summary)
    sheet_links = {
        "Not on Portal": "Missing in 2B",
        "Not in Tally": "Missing in Purchase",
        "Old Bill Filed": "Old Bill Filed",
        "Not Matched": "Tax Differences",
        "Date Difference": "All Discrepancies",
        "GSTIN Mismatch - Review": "Duplicates & Review",
        "Duplicate / Review Required": "Duplicates & Review",
        "Validation warnings": "Validation",
        "Duplicate / mismatch review entries": "Duplicate Keys",
        "All Discrepancies": "All Discrepancies",
    }
    for r in range(2, summary.max_row + 1):
        label = str(summary.cell(r, 1).value or "")
        target = sheet_links.get(label)
        if not target and label in {"Taxable Value Difference", "Tax Amount Difference", "Invoice Value Difference", "Tax Differences"}:
            target = "Tax Differences"
        if target and target in wb.sheetnames:
            summary.cell(r, 1).hyperlink = f"#'{target}'!A1"
            summary.cell(r, 1).style = "Hyperlink"
    for row in summary.iter_rows(min_row=2, min_col=1, max_col=2):
        row[0].font = Font(bold=True)
    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
    buffer = io.BytesIO(); wb.save(buffer); return buffer.getvalue()

def write_gst_subset(wb, title, columns, lines):
    ws = wb.create_sheet(title)
    write_sheet(ws, [c[0] for c in columns], [[line.get(c[1]) for c in columns] for line in lines])

def write_sheet(ws, headers, rows):
    ws.append(headers)
    for row in rows: ws.append(row)
    style_sheet(ws)

def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    if ws.max_row > 1 and ws.max_column > 0:
        ws.auto_filter.ref = ws.dimensions
    for col in range(1, ws.max_column + 1):
        values = [len(str(ws.cell(row=r, column=col).value or "")) for r in range(1, min(ws.max_row, 200) + 1)]
        ws.column_dimensions[get_column_letter(col)].width = min(max(max(values or [10]) + 2, 12), 42)
    ws.sheet_view.showGridLines = False
