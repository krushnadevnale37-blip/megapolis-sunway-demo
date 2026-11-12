from flask import Flask, request, render_template_string, send_file, flash, redirect, url_for, session, abort
from openpyxl import load_workbook
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from werkzeug.utils import secure_filename
from datetime import datetime
from decimal import Decimal, InvalidOperation
import os, io, re, sqlite3, secrets, time, zipfile, xml.etree.ElementTree as ET
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "0") == "1",
    PERMANENT_SESSION_LIFETIME=60 * 60 * 8,
)
DB_PATH = os.environ.get("DATABASE_PATH", os.path.join(os.path.dirname(__file__), "society_users.db"))

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_auth_db():
    with db() as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin','staff')),
            created_at TEXT NOT NULL
        )""")
        conn.execute("""CREATE TABLE IF NOT EXISTS activity_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            action TEXT NOT NULL,
            created_at TEXT NOT NULL
        )""")

def log_activity(username, action):
    with db() as conn:
        conn.execute("INSERT INTO activity_log(username,action,created_at) VALUES(?,?,?)",
                     (username, action, datetime.utcnow().isoformat(timespec="seconds")))

def user_count():
    with db() as conn:
        return conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped

def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login", next=request.path))
        if session.get("role") != "admin":
            abort(403)
        return view(*args, **kwargs)
    return wrapped

init_auth_db()
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB per request
ALLOWED_EXCEL = {".xlsx", ".xlsm"}
ALLOWED_WORD = {".docx"}


MAINTENANCE_PAGE = """
<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Maintenance Calculator</title><style>
body{margin:0;background:#f3f6fb;font-family:Segoe UI,Arial;color:#203047}.wrap{max-width:1000px;margin:30px auto;padding:0 18px}
.panel{background:white;border:1px solid #e2e8f0;border-radius:15px;padding:22px;margin:14px 0}
h1{color:#142b49;margin:0 0 6px}.muted{font-size:12px;color:#68778b;line-height:1.5}
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
label{display:block;font-size:12px;font-weight:700;margin-bottom:5px}input,select{width:100%;padding:10px;border:1px solid #d5deea;border-radius:8px;font-size:13px}
table{width:100%;border-collapse:collapse;margin-top:10px;font-size:12px}td,th{padding:9px;border-bottom:1px solid #edf1f6;text-align:left}button,.btn{display:inline-block;border:0;border-radius:8px;background:#285ca8;color:white;padding:11px 15px;font-weight:700;text-decoration:none;cursor:pointer}
.result{background:#edf4fc;padding:15px;border-radius:10px;margin-top:15px}.big{font-size:24px;font-weight:800;color:#174477}
@media(max-width:700px){.grid{grid-template-columns:1fr}}

/* Visual refresh: vibrant but professional society-software palette */
:root{--navy:#102a43;--blue:#2563eb;--teal:#0f766e;--violet:#6d4aff;--amber:#d97706;--ink:#1e293b;--muted:#64748b;--line:#dbe5f1;--soft:#f5f8ff}
body{background:radial-gradient(circle at 8% 0%,#e0edff 0,transparent 30%),radial-gradient(circle at 96% 10%,#e3f8f4 0,transparent 25%),#f4f7fc!important;color:var(--ink)}
.top{background:linear-gradient(115deg,#102a43 0%,#174a7e 55%,#2563a6 100%)!important;border-bottom:3px solid #43c6ac}
.logo{background:linear-gradient(135deg,#ffffff30,#ffffff10)!important;box-shadow:0 5px 16px #061b3333}
.local,.version{background:#ffffff1a!important;border-color:#ffffff35!important;color:#e5f1ff!important}
.panel{border:1px solid #dce6f2!important;border-radius:18px!important;box-shadow:0 12px 32px #173b6412,0 2px 5px #173b6408!important}
.upload{background:linear-gradient(110deg,#f8fbff,#f4f8ff)!important;border-color:#d7e5f6!important;transition:.18s}
.upload:hover{border-color:#8eb5ed!important;box-shadow:0 6px 18px #2563eb12;transform:translateY(-1px)}
.stepno{background:linear-gradient(135deg,#dbeafe,#e0f2fe)!important;color:#1d4ed8!important;box-shadow:inset 0 0 0 1px #bfdbfe}
button,.btn{background:linear-gradient(110deg,#2458c5,#397de0)!important;box-shadow:0 7px 18px #2458c533!important}
button:hover,.btn:hover{filter:brightness(1.06);transform:translateY(-1px)}
.section-title,.info-title{color:#173b64!important}
.tick{color:#0f9f78!important}
.note{background:linear-gradient(120deg,#eef6ff,#effaf8)!important;border-color:#d8e9f5!important}
input:focus,select:focus{outline:3px solid #bfdbfe;border-color:#3b82f6}
a{transition:.15s}
@media(max-width:760px){.panel{padding:18px!important}}

</style></head><body><div class="wrap">
<a href="/" class="btn">← Dashboard</a><h1 style="margin-top:18px">Maintenance Calculation</h1>
<p class="muted">Enter approved charge amounts and select the allocation method stated in your Society's registered bye-laws / General Body resolutions. This tool calculates only from the values and method you select; it does not decide the legally applicable method.</p>
<form method="post" action="/maintenance">
<div class="panel"><h3>Society / Billing Details</h3><div class="grid">
<div><label>Total Units</label><input type="number" name="units" min="1" value="{{units or 100}}" required></div>
<div><label>Total chargeable area (sq. ft.)</label><input type="number" step="0.01" name="total_area" min="0" value="{{total_area or 0}}"></div>
<div><label>Billing period (months)</label><select name="months"><option value="1" {% if months==1 %}selected{% endif %}>1 - Monthly</option><option value="3" {% if months==3 %}selected{% endif %}>3 - Quarterly</option><option value="6" {% if months==6 %}selected{% endif %}>6 - Half-yearly</option><option value="12" {% if months==12 %}selected{% endif %}>12 - Yearly</option></select></div>
</div></div>
<div class="panel"><h3>Charge Components</h3><p class="muted">For each charge enter the society-wide monthly amount and select how it is to be distributed.</p>
<table><tr><th>Charge Head</th><th>Total Monthly Amount (₹)</th><th>Allocation Method</th></tr>
{% for i in range(6) %}<tr><td><input name="head{{i}}" value="{{heads[i] if heads else default_heads[i]}}"></td><td><input type="number" step="0.01" min="0" name="amount{{i}}" value="{{amounts[i] if amounts else 0}}"></td><td><select name="method{{i}}"><option value="equal" {% if methods and methods[i]=='equal' %}selected{% endif %}>Equally per unit</option><option value="area" {% if methods and methods[i]=='area' %}selected{% endif %}>By area (sq. ft.)</option></select></td></tr>{% endfor %}
</table>
<div style="margin-top:16px"><label>Selected Flat Number (optional reference)</label><input name="flat" placeholder="e.g. A7-401"></div>
<div style="margin-top:14px"><label>Selected Flat Area (sq. ft.) - required for area-based charges</label><input type="number" step="0.01" min="0" name="flat_area" value="{{flat_area or 0}}"></div>
<button type="submit" style="margin-top:16px">Calculate Maintenance</button>
</div></form>
{% if result is not none %}<div class="panel"><h3>Calculation Result {% if flat %}— {{flat}}{% endif %}</h3>
<div class="result"><div class="muted">Estimated bill for selected billing period</div><div class="big">₹ {{ "{:,.2f}".format(result) }}</div><div class="muted">Monthly equivalent: ₹ {{ "{:,.2f}".format(monthly_result) }}</div></div>
<table><tr><th>Charge Head</th><th>Allocation</th><th>Monthly Share (₹)</th><th>Period Amount (₹)</th></tr>
{% for r in breakdown %}<tr><td>{{r.head}}</td><td>{{r.method}}</td><td>{{"{:,.2f}".format(r.monthly)}}</td><td>{{"{:,.2f}".format(r.period)}}</td></tr>{% endfor %}
</table></div>{% endif %}
<div class="muted">Note: Tax, interest, penalties, one-time charges, caps, exemptions and other special cases are not included unless separately modelled and approved. Verify the allocation basis and charge heads with your registered bye-laws, resolutions and current law.</div>
</div></body></html>
"""

HELP_PAGE = """
<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Society Help Centre</title><style>
body{margin:0;background:#f3f6fb;font-family:Segoe UI,Arial;color:#203047}.wrap{max-width:1000px;margin:30px auto;padding:0 18px}
.panel{background:white;border:1px solid #e2e8f0;border-radius:15px;padding:22px;margin:14px 0}
h1{color:#142b49;margin:0 0 6px}.muted{font-size:12px;color:#68778b;line-height:1.6}
input{width:100%;padding:13px;border:1px solid #d5deea;border-radius:8px;font-size:14px}button,.btn{display:inline-block;border:0;border-radius:8px;background:#285ca8;color:white;padding:11px 15px;font-weight:700;text-decoration:none;cursor:pointer;margin-top:10px}
.item{padding:15px 0;border-bottom:1px solid #edf1f6}.item:last-child{border:0}.tag{font-size:10px;background:#edf4fc;color:#285ca8;padding:4px 8px;border-radius:20px}

/* Visual refresh: vibrant but professional society-software palette */
:root{--navy:#102a43;--blue:#2563eb;--teal:#0f766e;--violet:#6d4aff;--amber:#d97706;--ink:#1e293b;--muted:#64748b;--line:#dbe5f1;--soft:#f5f8ff}
body{background:radial-gradient(circle at 8% 0%,#e0edff 0,transparent 30%),radial-gradient(circle at 96% 10%,#e3f8f4 0,transparent 25%),#f4f7fc!important;color:var(--ink)}
.top{background:linear-gradient(115deg,#102a43 0%,#174a7e 55%,#2563a6 100%)!important;border-bottom:3px solid #43c6ac}
.logo{background:linear-gradient(135deg,#ffffff30,#ffffff10)!important;box-shadow:0 5px 16px #061b3333}
.local,.version{background:#ffffff1a!important;border-color:#ffffff35!important;color:#e5f1ff!important}
.panel{border:1px solid #dce6f2!important;border-radius:18px!important;box-shadow:0 12px 32px #173b6412,0 2px 5px #173b6408!important}
.upload{background:linear-gradient(110deg,#f8fbff,#f4f8ff)!important;border-color:#d7e5f6!important;transition:.18s}
.upload:hover{border-color:#8eb5ed!important;box-shadow:0 6px 18px #2563eb12;transform:translateY(-1px)}
.stepno{background:linear-gradient(135deg,#dbeafe,#e0f2fe)!important;color:#1d4ed8!important;box-shadow:inset 0 0 0 1px #bfdbfe}
button,.btn{background:linear-gradient(110deg,#2458c5,#397de0)!important;box-shadow:0 7px 18px #2458c533!important}
button:hover,.btn:hover{filter:brightness(1.06);transform:translateY(-1px)}
.section-title,.info-title{color:#173b64!important}
.tick{color:#0f9f78!important}
.note{background:linear-gradient(120deg,#eef6ff,#effaf8)!important;border-color:#d8e9f5!important}
input:focus,select:focus{outline:3px solid #bfdbfe;border-color:#3b82f6}
a{transition:.15s}
@media(max-width:760px){.panel{padding:18px!important}}

</style></head><body><div class="wrap"><a href="/" class="btn">← Dashboard</a>
<h1 style="margin-top:18px">Housing Society Help Centre</h1>
<p class="muted">Search the built-in starter reference topics. This is an offline keyword reference—not an AI legal opinion. For current requirements, open the official sources and confirm your Society's registered bye-laws and applicable resolutions.</p>
<div class="panel"><form method="get" action="/society-help"><input name="q" value="{{q}}" placeholder="Search: maintenance, sinking fund, AGM, audit, parking, NOC, bonus..."><button type="submit">Search Knowledge Base</button></form>
{% if q %}<h3>Search results for “{{q}}”</h3>{% endif %}
{% for item in results %}<div class="item"><span class="tag">{{item.category}}</span><h3>{{item.title}}</h3><p class="muted">{{item.answer}}</p><p class="muted"><b>What to verify:</b> {{item.verify}}</p></div>{% else %}<p class="muted">No matching topic found. Try a broader keyword or use the official research links below.</p>{% endfor %}
</div>
<div class="panel"><h3>Official research starting points</h3>
<p><a href="https://sahakarayukta.maharashtra.gov.in/" target="_blank">Maharashtra Cooperation Commissioner & Registrar</a></p>
<p><a href="https://mahasahakar.maharashtra.gov.in/" target="_blank">Maharashtra Cooperation Department</a></p>
<p><a href="https://www.google.com/search?q=Maharashtra+cooperative+housing+society+model+bye+laws+official+PDF" target="_blank">Find official Model Bye-laws PDF</a></p>
<p class="muted">The model bye-laws are a reference. The Society's registered bye-laws, amendments, resolutions, current Acts/Rules and Government circulars may affect the answer.</p>
</div></div></body></html>
"""

PAGE = """
<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Society Analytics</title><style>
body{margin:0;background:#f3f6fb;font-family:Segoe UI,Arial;color:#203047}.wrap{max-width:1100px;margin:30px auto;padding:0 18px}
h1{color:#142b49;margin-bottom:4px}.muted{color:#68778b;font-size:13px}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:20px 0}
.card{background:white;border:1px solid #e2e8f0;border-radius:13px;padding:17px;box-shadow:0 5px 18px #18314d09}
.k{font-size:11px;color:#68778b}.v{font-size:23px;font-weight:800;color:#18395f;margin-top:7px}
.cols{display:grid;grid-template-columns:1.1fr .9fr;gap:14px}.card h3{font-size:14px;margin:0 0 14px}
table{width:100%;border-collapse:collapse;font-size:12px}td,th{text-align:left;padding:9px 7px;border-bottom:1px solid #edf1f6}th{color:#5a6a7f}
.bar{height:9px;background:#e8eef7;border-radius:8px;overflow:hidden}.bar span{display:block;height:100%;background:#3476c7}
.btn{display:inline-block;background:#285ca8;color:white;text-decoration:none;padding:10px 14px;border-radius:8px;font-size:12px;font-weight:700;margin-top:15px}
@media(max-width:750px){.grid{grid-template-columns:repeat(2,1fr)}.cols{grid-template-columns:1fr}}

/* Visual refresh: vibrant but professional society-software palette */
:root{--navy:#102a43;--blue:#2563eb;--teal:#0f766e;--violet:#6d4aff;--amber:#d97706;--ink:#1e293b;--muted:#64748b;--line:#dbe5f1;--soft:#f5f8ff}
body{background:radial-gradient(circle at 8% 0%,#e0edff 0,transparent 30%),radial-gradient(circle at 96% 10%,#e3f8f4 0,transparent 25%),#f4f7fc!important;color:var(--ink)}
.top{background:linear-gradient(115deg,#102a43 0%,#174a7e 55%,#2563a6 100%)!important;border-bottom:3px solid #43c6ac}
.logo{background:linear-gradient(135deg,#ffffff30,#ffffff10)!important;box-shadow:0 5px 16px #061b3333}
.local,.version{background:#ffffff1a!important;border-color:#ffffff35!important;color:#e5f1ff!important}
.panel{border:1px solid #dce6f2!important;border-radius:18px!important;box-shadow:0 12px 32px #173b6412,0 2px 5px #173b6408!important}
.upload{background:linear-gradient(110deg,#f8fbff,#f4f8ff)!important;border-color:#d7e5f6!important;transition:.18s}
.upload:hover{border-color:#8eb5ed!important;box-shadow:0 6px 18px #2563eb12;transform:translateY(-1px)}
.stepno{background:linear-gradient(135deg,#dbeafe,#e0f2fe)!important;color:#1d4ed8!important;box-shadow:inset 0 0 0 1px #bfdbfe}
button,.btn{background:linear-gradient(110deg,#2458c5,#397de0)!important;box-shadow:0 7px 18px #2458c533!important}
button:hover,.btn:hover{filter:brightness(1.06);transform:translateY(-1px)}
.section-title,.info-title{color:#173b64!important}
.tick{color:#0f9f78!important}
.note{background:linear-gradient(120deg,#eef6ff,#effaf8)!important;border-color:#d8e9f5!important}
input:focus,select:focus{outline:3px solid #bfdbfe;border-color:#3b82f6}
a{transition:.15s}
@media(max-width:760px){.panel{padding:18px!important}}

</style></head><body><div class="wrap">
<h1>Defaulter Analytics</h1><p class="muted">Summary calculated from the uploaded Excel file. Verify totals with your accounting records before action.</p>
<div class="grid">
<div class="card"><div class="k">TOTAL OUTSTANDING</div><div class="v">₹ {{total}}</div></div>
<div class="card"><div class="k">TOTAL FLATS IN FILE</div><div class="v">{{rows}}</div></div>
<div class="card"><div class="k">BLOCKS COVERED</div><div class="v">{{blocks}}</div></div>
<div class="card"><div class="k">AVERAGE DUES / FLAT</div><div class="v">₹ {{avg}}</div></div>
</div>
<div class="cols">
<div class="card"><h3>Block-wise Outstanding Analysis</h3><table><tr><th>Block</th><th>Flats</th><th>Outstanding</th><th>Share</th></tr>
{% for b in by_block %}<tr><td>{{b.block}}</td><td>{{b.flats}}</td><td>₹ {{ "{:,.2f}".format(b.amount) }}</td><td style="width:25%"><div class="bar"><span style="width:{{ (b.amount / (by_block[0].amount if by_block and by_block[0].amount else 1) * 100)|round(1) }}%"></span></div></td></tr>{% endfor %}</table></div>
<div class="card"><h3>Overdue Ageing (number of flats)</h3><table><tr><th>Overdue Days</th><th>Flats</th></tr>{% for k,v in buckets.items() %}<tr><td>{{k}}</td><td>{{v}}</td></tr>{% endfor %}</table>
<h3 style="margin-top:22px">Top 10 Outstanding Flats</h3><table><tr><th>Block / Flat</th><th>Days</th><th>Amount</th></tr>{% for x in top %}<tr><td>{{x.block}} / {{x.unit}}</td><td>{{x.days}}</td><td>₹ {{ "{:,.2f}".format(x.amount) }}</td></tr>{% endfor %}</table></div>
</div><a class="btn" href="/">← Back to Report Studio</a></div></body></html>
"""

PAGE = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Megapolis Sunway | Defaulter Report Studio</title>
<style>
:root{--navy:#132b49;--blue:#285ca8;--ink:#202d40;--muted:#66768b;--line:#e2e8f0;--soft:#f5f8fc;--green:#18794e}
*{box-sizing:border-box}body{margin:0;background:#f3f6fb;color:var(--ink);font-family:Inter,"Segoe UI",Arial,sans-serif}
.top{background:linear-gradient(110deg,#132b49,#1e416c);color:white;padding:19px 24px;box-shadow:0 5px 22px #122b4920}
.topin{max-width:1060px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:15px}
.brand{display:flex;align-items:center;gap:13px}.logo{width:44px;height:44px;border-radius:12px;background:#ffffff18;border:1px solid #ffffff35;display:grid;place-items:center;font-size:22px}
.brandname{font-weight:800;letter-spacing:.35px;font-size:16px}.brandname small{display:block;font-size:10px;font-weight:500;color:#c9d7e8;letter-spacing:1.3px;margin-top:4px}
.local{font-size:11px;color:#dce8f5;border:1px solid #ffffff35;padding:8px 11px;border-radius:30px}
.wrap{max-width:1060px;margin:34px auto;padding:0 20px 40px}
.hero{display:flex;justify-content:space-between;align-items:end;gap:20px;margin-bottom:22px}.hero h1{margin:0 0 7px;color:var(--navy);font-size:28px}.hero p{margin:0;color:var(--muted);font-size:14px}.version{font-size:11px;color:#66768b;background:white;border:1px solid var(--line);padding:7px 10px;border-radius:20px;white-space:nowrap}
.layout{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(260px,.85fr);gap:20px;align-items:start}
.panel{background:white;border:1px solid var(--line);border-radius:17px;padding:24px;box-shadow:0 10px 28px #1b35500b}
.section-title{font-size:15px;font-weight:800;color:#243a57;margin:0 0 5px}.section-sub{font-size:12px;color:var(--muted);margin:0 0 18px}
.upload{display:flex;gap:13px;padding:17px;border:1px solid var(--line);border-radius:13px;margin-bottom:13px;background:#fcfdff}
.stepno{width:30px;height:30px;flex:0 0 30px;border-radius:9px;background:#eaf1fb;color:var(--blue);display:grid;place-items:center;font-weight:800;font-size:13px}
.field{flex:1;min-width:0}.field label{display:block;font-weight:750;font-size:13px;margin:1px 0 6px}.field p{font-size:11px;color:var(--muted);margin:0 0 10px}
input[type=file]{width:100%;font-size:12px;border:1px solid #d5deea;background:white;border-radius:8px;padding:9px}
input[type=file]::file-selector-button{border:0;border-radius:6px;background:#edf2f9;color:#29466d;padding:7px 10px;margin-right:9px;font-weight:700;cursor:pointer}
.filename{font-size:11px;color:var(--green);margin-top:7px;min-height:14px;overflow-wrap:anywhere}
button{width:100%;border:0;border-radius:10px;background:linear-gradient(110deg,#285ca8,#3476c7);color:white;font-size:14px;font-weight:800;padding:14px;cursor:pointer;margin-top:6px;box-shadow:0 7px 15px #285ca82b}
button:disabled{opacity:.65;cursor:wait}.alert{padding:12px 14px;background:#fff6e8;border:1px solid #f2d6a6;border-radius:9px;margin-bottom:14px;color:#80551b;font-size:12px}
.side{display:grid;gap:14px}.info-title{font-size:13px;font-weight:800;margin:0 0 12px;color:#263b58}
.check{display:flex;gap:9px;align-items:flex-start;margin:10px 0;font-size:12px;color:#4c5c71;line-height:1.45}.tick{color:var(--green);font-weight:900}
.schema{width:100%;border-collapse:collapse;font-size:11px}.schema td{padding:8px 5px;border-bottom:1px solid #edf1f6}.schema td:first-child{font-weight:700;color:#354a66}
.note{background:#f1f6fd;border:1px solid #e1ebf8;border-radius:11px;padding:13px;font-size:11px;line-height:1.55;color:#52647b}
.privacy{display:flex;gap:9px;align-items:flex-start;font-size:11px;color:#63748a;line-height:1.5;margin-top:16px}.privacy b{color:#354a66}
footer{text-align:center;color:#93a0b1;font-size:10px;margin-top:24px}
@media(max-width:760px){.layout{grid-template-columns:1fr}.hero{align-items:start;flex-direction:column}.local{display:none}}

/* Visual refresh: vibrant but professional society-software palette */
:root{--navy:#102a43;--blue:#2563eb;--teal:#0f766e;--violet:#6d4aff;--amber:#d97706;--ink:#1e293b;--muted:#64748b;--line:#dbe5f1;--soft:#f5f8ff}
body{background:radial-gradient(circle at 8% 0%,#e0edff 0,transparent 30%),radial-gradient(circle at 96% 10%,#e3f8f4 0,transparent 25%),#f4f7fc!important;color:var(--ink)}
.top{background:linear-gradient(115deg,#102a43 0%,#174a7e 55%,#2563a6 100%)!important;border-bottom:3px solid #43c6ac}
.logo{background:linear-gradient(135deg,#ffffff30,#ffffff10)!important;box-shadow:0 5px 16px #061b3333}
.local,.version{background:#ffffff1a!important;border-color:#ffffff35!important;color:#e5f1ff!important}
.panel{border:1px solid #dce6f2!important;border-radius:18px!important;box-shadow:0 12px 32px #173b6412,0 2px 5px #173b6408!important}
.upload{background:linear-gradient(110deg,#f8fbff,#f4f8ff)!important;border-color:#d7e5f6!important;transition:.18s}
.upload:hover{border-color:#8eb5ed!important;box-shadow:0 6px 18px #2563eb12;transform:translateY(-1px)}
.stepno{background:linear-gradient(135deg,#dbeafe,#e0f2fe)!important;color:#1d4ed8!important;box-shadow:inset 0 0 0 1px #bfdbfe}
button,.btn{background:linear-gradient(110deg,#2458c5,#397de0)!important;box-shadow:0 7px 18px #2458c533!important}
button:hover,.btn:hover{filter:brightness(1.06);transform:translateY(-1px)}
.section-title,.info-title{color:#173b64!important}
.tick{color:#0f9f78!important}
.note{background:linear-gradient(120deg,#eef6ff,#effaf8)!important;border-color:#d8e9f5!important}
input:focus,select:focus{outline:3px solid #bfdbfe;border-color:#3b82f6}
a{transition:.15s}
@media(max-width:760px){.panel{padding:18px!important}}

</style></head>
<body>
<header class="top"><div class="topin"><div class="brand"><div class="logo">▦</div><div class="brandname">MEGAPOLIS SUNWAY<small>CO-OPERATIVE HOUSING SOCIETY LTD.</small></div></div><div class="local">● LOCAL COMPUTER PROCESSING</div></div></header>
<main class="wrap">
<div class="hero"><div><h1>Defaulter Report Studio</h1><p>Prepare a clean, block-wise member circulation report from your Excel data.</p></div><div class="version">REPORT GENERATOR · LOCAL VERSION</div></div>
<div style="display:flex;gap:9px;flex-wrap:wrap;margin:0 0 18px">
<a href="/maintenance" style="background:linear-gradient(110deg,#2458c5,#397de0);color:white;text-decoration:none;padding:10px 14px;border-radius:9px;box-shadow:0 5px 13px #2458c526;font-size:12px;font-weight:700">Maintenance Calculator</a>
<a href="/society-help" style="background:linear-gradient(110deg,#ffffff,#eff6ff);color:#2458a8;border:1px solid #c9d8eb;text-decoration:none;padding:10px 14px;border-radius:9px;font-size:12px;font-weight:700">Society Help Centre</a>
</div>
<div class="layout">
<section class="panel">
<h2 class="section-title">Create a new report</h2><p class="section-sub">Upload both files, confirm the selected names and download your ZIP report.</p>
{% with messages = get_flashed_messages() %}{% if messages %}{% for m in messages %}<div class="alert">{{m}}</div>{% endfor %}{% endif %}{% endwith %}
<form id="reportForm" method="post" action="/generate" enctype="multipart/form-data">
<div class="upload"><div class="stepno">01</div><div class="field"><label for="excel">Defaulter Excel File</label><p>Choose the latest member outstanding list (.xlsx / .xlsm).</p><input id="excel" type="file" name="excel" accept=".xlsx,.xlsm" required><div class="filename" id="excelName">No Excel file selected</div></div></div>
<div class="upload"><div class="stepno">02</div><div class="field"><label for="template">Society Word Template</label><p>Choose the approved Society template (.docx).</p><input id="template" type="file" name="template" accept=".docx" required><div class="filename" id="wordName">No Word template selected</div></div></div>
<button id="submitBtn" type="submit">Generate &amp; Download ZIP Report&nbsp; →</button>
<div class="privacy"><span>🔒</span><div><b>Your files stay on this computer.</b> This local tool processes the selected Excel and Word files on your PC; it does not upload them to a public website.</div></div>
</form>
<form method="post" action="/analyze" enctype="multipart/form-data" style="border-top:1px solid #e8edf4;margin-top:17px;padding-top:15px">
<label style="display:block;font-size:12px;font-weight:750;margin-bottom:8px;color:#314866">Want to analyse the Excel before generating the report?</label>
<input type="file" name="excel" accept=".xlsx,.xlsm" required>
<button type="submit" style="background:#fff;color:#285ca8;border:1px solid #b8cce8;box-shadow:none;margin-top:10px">View Defaulter Analytics</button>
</form></section>
<aside class="side">
<div class="panel"><h3 class="info-title">Report features</h3>
<div class="check"><span class="tick">✓</span><span>Separate page for each Block</span></div>
<div class="check"><span class="tick">✓</span><span>Flat Number, Overdue Days and Pending Amount</span></div>
<div class="check"><span class="tick">✓</span><span>No Block Total or Grand Total rows</span></div>
<div class="check"><span class="tick">✓</span><span>Payment reminder on every page</span></div>
<div class="check"><span class="tick">✓</span><span>Separate Accountant and Treasurer signature spaces on every page</span></div>
<div class="check"><span class="tick">✓</span><span>Society Word template header and footer retained</span></div>
</div>
<div class="panel"><h3 class="info-title">Required Excel columns</h3>
<table class="schema"><tr><td>Block</td><td>Building / block code</td></tr><tr><td>Unit #</td><td>Flat number</td></tr><tr><td>Overdue Days</td><td>Pending days</td></tr><tr><td>Pending Amount</td><td>Outstanding amount</td></tr></table>
<div class="note" style="margin-top:12px">The column headings must match these names. The first worksheet in the Excel file is used.</div></div>
<div class="panel"><h3 class="info-title">Maintenance Calculator</h3>
<p style="font-size:11px;color:#68778b;line-height:1.5">Calculate monthly, quarterly, half-yearly or yearly charges with equal-per-unit or area-based allocation selected by you.</p>
<a href="/maintenance" style="display:inline-block;background:linear-gradient(110deg,#2458c5,#397de0);color:white;text-decoration:none;padding:10px 14px;border-radius:9px;box-shadow:0 5px 13px #2458c526;font-size:12px;font-weight:700">Open Maintenance Calculator →</a></div>
<div class="panel"><h3 class="info-title">Society Rules &amp; Knowledge</h3>
<p style="font-size:11px;color:#68778b;line-height:1.5">Search the built-in Housing Society topic library and open official Maharashtra reference sources. No API key required; this is not generative AI.</p>
<a href="/society-help" style="display:inline-block;background:linear-gradient(110deg,#2458c5,#397de0);color:white;text-decoration:none;padding:10px 14px;border-radius:9px;box-shadow:0 5px 13px #2458c526;font-size:12px;font-weight:700">Open Society Help Centre →</a></div>
<div class="note"><b>Before circulation</b><br>Review the generated Word report, verify member dues with Society records and obtain required committee approval before sharing.</div>
</aside></div>
<footer>MEGAPOLIS SUNWAY · INTERNAL REPORTING TOOL · LOCAL COMPUTER USE</footer>
</main>
<script>
const excel=document.getElementById('excel'), word=document.getElementById('template');
excel.addEventListener('change',()=>document.getElementById('excelName').textContent=excel.files[0]?.name||'No Excel file selected');
word.addEventListener('change',()=>document.getElementById('wordName').textContent=word.files[0]?.name||'No Word template selected');
document.getElementById('reportForm').addEventListener('submit',()=>{const b=document.getElementById('submitBtn');b.disabled=true;b.textContent='Generating report… Please wait';});
</script></body></html>
"""

def money(v):
    try:
        if v is None or str(v).strip() == "":
            return 0.0
        return float(str(v).replace(",", "").replace("₹", "").strip())
    except (ValueError, TypeError):
        return 0.0

def fmt_money(v):
    return f"{v:,.2f}"

def shade(cell, fill="D9E2F3"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)

def set_cell_text(cell, value, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    r = p.add_run(str(value))
    r.bold = bold
    r.font.name = "Arial"
    r.font.size = Pt(9)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)

def repair_excel_styles(file_stream):
    """Repair common malformed XLSX/XLSM style records before openpyxl reads them.

    Some Excel exports contain empty <fill/> entries. openpyxl expects each
    fill entry to deserialize into a Fill object and otherwise raises errors
    such as: expected <class 'openpyxl.styles.fills.Fill'>.
    This function repairs those style records in memory without changing the
    uploaded original file or its worksheet data.
    """
    file_stream.seek(0)
    raw = file_stream.read()
    file_stream.seek(0)
    try:
        source = io.BytesIO(raw)
        output_files = []
        changed = False
        with zipfile.ZipFile(source, 'r') as zin:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == 'xl/styles.xml':
                    root = ET.fromstring(data)
                    ns_uri = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                    ns = {'m': ns_uri}
                    fills = root.find('m:fills', ns)
                    if fills is not None:
                        for fill in list(fills):
                            pattern = fill.find('m:patternFill', ns)
                            gradient = fill.find('m:gradientFill', ns)
                            if pattern is None and gradient is None:
                                ET.SubElement(fill, '{%s}patternFill' % ns_uri,
                                              {'patternType': 'none'})
                                changed = True
                        fills.set('count', str(len(list(fills))))

                    # Guard against style records pointing to non-existent fills.
                    if fills is not None:
                        fill_count = len(list(fills))
                        for xf_group_name in ('cellXfs', 'cellStyleXfs'):
                            group = root.find('m:' + xf_group_name, ns)
                            if group is None:
                                continue
                            for xf in list(group):
                                fill_id = xf.get('fillId')
                                if fill_id is not None:
                                    try:
                                        invalid = int(fill_id) >= fill_count or int(fill_id) < 0
                                    except ValueError:
                                        invalid = True
                                    if invalid:
                                        xf.set('fillId', '0')
                                        changed = True
                    if changed:
                        ET.register_namespace('', ns_uri)
                        ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
                        ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
                        data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                output_files.append((info, data))

        if not changed:
            return io.BytesIO(raw)
        repaired = io.BytesIO()
        with zipfile.ZipFile(repaired, 'w', zipfile.ZIP_DEFLATED) as zout:
            for info, data in output_files:
                zout.writestr(info, data)
        repaired.seek(0)
        return repaired
    except zipfile.BadZipFile:
        raise ValueError('This is not a valid .xlsx/.xlsm Excel file, or its ZIP structure is damaged. Open it in Excel and use Save As.')
    except ET.ParseError:
        raise ValueError('Excel formatting information could not be read. Open the file in Excel and save it as a new .xlsx file.')


def read_rows(file_stream):
    safe_stream = repair_excel_styles(file_stream)
    try:
        wb = load_workbook(safe_stream, data_only=True)
    except Exception as exc:
        raise ValueError('Could not read the Excel file. It may be password-protected, unsupported, or severely corrupted. Details: ' + str(exc))
    ws = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    required = ["Block", "Unit #", "Overdue Days", "Pending Amount"]
    norm = {h.lower(): i for i, h in enumerate(headers)}
    missing = [h for h in required if h.lower() not in norm]
    if missing:
        raise ValueError("Required columns were not found in the Excel file: " + ", ".join(missing))
    ix = {h: norm[h.lower()] for h in required}
    grouped = {}
    for vals in ws.iter_rows(min_row=2, values_only=True):
        block = vals[ix["Block"]]
        unit = vals[ix["Unit #"]]
        if block is None or unit is None:
            continue
        block = str(block).strip()
        unit = str(unit).strip()
        overdue = vals[ix["Overdue Days"]]
        amount = money(vals[ix["Pending Amount"]])
        grouped.setdefault(block, []).append((unit, overdue if overdue is not None else "", amount))
    if not grouped:
        raise ValueError("No report data was found in the Excel file.")
    return dict(sorted(grouped.items(), key=lambda x: x[0]))

def clear_document_body(doc):
    body = doc._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

def generate_docx(excel_file, template_file):
    grouped = read_rows(excel_file)
    doc = Document(template_file)
    clear_document_body(doc)

    # Page setup and body style
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

    all_total = 0.0
    for page_idx, (block, rows) in enumerate(grouped.items()):
        if page_idx:
            doc.add_page_break()

        # Add a little top spacing so the report content sits lower on the page.
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(10)
        spacer.paragraph_format.line_spacing = 1.0

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(5)
        r = p.add_run("MAINTENANCE DEFAULTER LIST")
        r.bold = True; r.font.name = "Arial"; r.font.size = Pt(16); r.font.color.rgb = RGBColor(20,43,74)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
        r = p.add_run(f"BLOCK {block}")
        r.bold = True; r.font.name = "Arial"; r.font.size = Pt(13); r.font.color.rgb = RGBColor(36,87,167)

        table = doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Apply borders directly so this also works with Word templates
        # that do not contain the built-in "Table Grid" style.
        tblPr = table._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            element = OxmlElement(f"w:{edge}")
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "5")
            element.set(qn("w:space"), "0")
            element.set(qn("w:color"), "808080")
            borders.append(element)
        tblPr.append(borders)
        headers = ["Flat Number", "Overdue Days", "Pending Amount (₹)"]
        for i, h in enumerate(headers):
            set_cell_text(table.rows[0].cells[i], h, True)
            shade(table.rows[0].cells[i], "D9E2F3")
        set_repeat_table_header(table.rows[0])

        block_total = 0.0
        for unit, days, amount in rows:
            cells = table.add_row().cells
            set_cell_text(cells[0], unit)
            set_cell_text(cells[1], days)
            set_cell_text(cells[2], fmt_money(amount), align=WD_ALIGN_PARAGRAPH.CENTER)
            block_total += amount
        all_total += block_total

        # Payment reminder and approval signatures on EVERY block page.
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(10)
        r = p.add_run("Members are requested to verify their outstanding dues and arrange payment at the earliest. If payment has already been made, please share the payment details with the Society office for reconciliation.")
        r.font.name = "Arial"
        r.font.size = Pt(9)

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run("For Megapolis Sunway Co-op. Housing Society Ltd.")
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(10)

        # Two independent signature spaces side-by-side.
        sign_table = doc.add_table(rows=1, cols=2)
        sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sign_table.autofit = False
        for cell in sign_table.rows[0].cells:
            cell.width = Inches(3.0)
        for cell, label in zip(sign_table.rows[0].cells, ("Accountant", "Treasurer")):
            cell.text = ""
            p_line = cell.paragraphs[0]
            p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_line.paragraph_format.space_before = Pt(30)
            p_line.paragraph_format.space_after = Pt(4)
            run = p_line.add_run("____________________________")
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(10)
            p_label = cell.add_paragraph()
            p_label.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_label.paragraph_format.space_after = Pt(0)
            label_run = p_label.add_run(label)
            label_run.bold = True
            label_run.font.name = "Arial"
            label_run.font.size = Pt(10)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM

        # Hide signature table borders.
        tblPr = sign_table._tbl.tblPr
        no_borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            element = OxmlElement(f"w:{edge}")
            element.set(qn("w:val"), "nil")
            no_borders.append(element)
        tblPr.append(no_borders)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output, len(grouped), sum(len(v) for v in grouped.values()), all_total


AUTH_STYLE = """
<style>
body{margin:0;background:radial-gradient(circle at 10% 5%,#dbeafe,transparent 30%),#f2f6fc;font-family:Segoe UI,Arial;color:#1e293b}
.authwrap{max-width:440px;margin:7vh auto;padding:18px}.authcard{background:#fff;border:1px solid #dce6f2;border-radius:20px;padding:28px;box-shadow:0 18px 45px #12385d18}
.brand{color:#173b64;font-size:13px;font-weight:800;letter-spacing:1px}.authcard h1{font-size:24px;color:#132f50;margin:20px 0 6px}
.sub{color:#64748b;font-size:12px;line-height:1.6;margin-bottom:22px}
label{display:block;font-size:12px;font-weight:700;margin:13px 0 6px}
input{width:100%;box-sizing:border-box;padding:12px;border:1px solid #cbd8e8;border-radius:9px;font-size:14px}
button{width:100%;margin-top:20px;border:0;border-radius:9px;padding:13px;background:linear-gradient(110deg,#2458c5,#397de0);color:white;font-weight:800;cursor:pointer}
.alert{background:#fff4e5;border:1px solid #f2d3a1;color:#80551b;border-radius:8px;padding:10px;font-size:12px;margin:10px 0}
small{color:#64748b;font-size:11px}
</style>
"""
LOGIN_PAGE = """<!doctype html><html><head><meta charset="utf-8"><title>Society Login</title>""" + AUTH_STYLE + """</head><body><div class="authwrap"><div class="authcard">
<div class="brand">MEGAPOLIS SUNWAY · SOCIETY PORTAL</div><h1>Welcome back</h1><div class="sub">Sign in with your authorised Society account.</div>
{% with messages=get_flashed_messages() %}{% for m in messages %}<div class="alert">{{m}}</div>{% endfor %}{% endwith %}
<form method="post"><label>Username</label><input name="username" autocomplete="username" required autofocus>
<label>Password</label><input type="password" name="password" autocomplete="current-password" required>
<button type="submit">Secure Login</button></form><p><small>Access is restricted to authorised Society staff.</small></p></div></div></body></html>"""
SETUP_PAGE = """<!doctype html><html><head><meta charset="utf-8"><title>Initial Admin Setup</title>""" + AUTH_STYLE + """</head><body><div class="authwrap"><div class="authcard">
<div class="brand">MEGAPOLIS SUNWAY · FIRST-TIME SETUP</div><h1>Create Administrator</h1><div class="sub">Create the first Admin account. Use a unique password of at least 12 characters. This setup page disables itself after the first Admin is created.</div>
{% with messages=get_flashed_messages() %}{% for m in messages %}<div class="alert">{{m}}</div>{% endfor %}{% endwith %}
<form method="post"><label>Admin Username</label><input name="username" minlength="4" maxlength="40" required>
<label>Admin Password</label><input type="password" name="password" minlength="12" required>
<label>Confirm Password</label><input type="password" name="confirm" minlength="12" required>
<button type="submit">Create Admin Account</button></form></div></div></body></html>"""
ADMIN_PAGE = """<!doctype html><html><head><meta charset="utf-8"><title>Admin Control</title>""" + AUTH_STYLE + """</head><body><div class="authwrap" style="max-width:650px"><div class="authcard">
<div class="brand">MEGAPOLIS SUNWAY · ADMIN CONTROL</div><h1>Staff Access Management</h1><div class="sub">You selected one shared Staff login. Admin can create or reset the shared Staff username and password here.</div>
{% with messages=get_flashed_messages() %}{% for m in messages %}<div class="alert">{{m}}</div>{% endfor %}{% endwith %}
<form method="post"><label>Shared Staff Username</label><input name="username" minlength="4" maxlength="40" required>
<label>Shared Staff Password (minimum 12 characters)</label><input type="password" name="password" minlength="12" required>
<label>Confirm Password</label><input type="password" name="confirm" minlength="12" required>
<button type="submit">Save / Update Shared Staff Login</button></form>
<h3>Current Accounts</h3><p><small>{% for u in users %}{{u.username}} — {{u.role}}<br>{% endfor %}</small></p>
<h3>Recent Activity</h3><p><small>{% for a in logs %}{{a.created_at}} · {{a.username}} · {{a.action}}<br>{% endfor %}</small></p>
<p><a href="/">Open Dashboard</a> · <a href="/logout">Logout</a></p></div></div></body></html>"""

@app.before_request
def enforce_login():
    # Public demo mode: no Admin setup, login, or session gate.
    return None


@app.route("/setup", methods=["GET","POST"])
def setup():
    abort(404)
    if user_count():
        abort(404)
    if request.method == "POST":
        username=request.form.get("username","").strip()
        password=request.form.get("password","")
        confirm=request.form.get("confirm","")
        if len(username)<4 or len(password)<12 or password!=confirm:
            flash("Username must contain at least 4 characters and password must contain at least 12 characters. Password confirmation must match.")
            return redirect(url_for("setup"))
        with db() as conn:
            cur=conn.execute("INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)",
                (username,generate_password_hash(password), "admin",datetime.utcnow().isoformat(timespec="seconds")))
            uid=cur.lastrowid
        session.clear();session["user_id"]=uid;session["username"]=username;session["role"]="admin";session.permanent=True
        log_activity(username,"Created initial Admin account")
        return redirect(url_for("admin"))
    return render_template_string(SETUP_PAGE)

@app.route("/login", methods=["GET","POST"])
def login():
    abort(404)
    if not user_count():
        return redirect(url_for("setup"))
    locked_until=session.get("locked_until",0)
    if time.time()<locked_until:
        flash("Too many failed attempts. Please wait 5 minutes before trying again.")
        return render_template_string(LOGIN_PAGE), 429
    if request.method=="POST":
        username=request.form.get("username","").strip()
        password=request.form.get("password","")
        with db() as conn:
            user=conn.execute("SELECT * FROM users WHERE username=?",(username,)).fetchone()
        if user and check_password_hash(user["password_hash"],password):
            session.clear();session["user_id"]=user["id"];session["username"]=user["username"];session["role"]=user["role"];session.permanent=True
            log_activity(username,"Successful login")
            target=request.args.get("next","/")
            if not target.startswith("/") or target.startswith("//"): target="/"
            return redirect(target)
        fails=session.get("fails",0)+1
        session["fails"]=fails
        if fails>=5:
            session["fails"]=0;session["locked_until"]=time.time()+300
        flash("Incorrect username or password.")
    return render_template_string(LOGIN_PAGE)

@app.route("/logout")
def logout():
    abort(404)
    if session.get("username"):
        log_activity(session["username"],"Logout")
    session.clear()
    return redirect(url_for("login"))

@app.route("/admin", methods=["GET","POST"])
def admin():
    abort(404)
    if request.method=="POST":
        username=request.form.get("username","").strip()
        password=request.form.get("password","")
        confirm=request.form.get("confirm","")
        if len(username)<4 or len(password)<12 or password!=confirm:
            flash("Username must contain at least 4 characters and password must contain at least 12 characters. Please check the confirmation.")
            return redirect(url_for("admin"))
        with db() as conn:
            existing=conn.execute("SELECT id FROM users WHERE role='staff' LIMIT 1").fetchone()
            if existing:
                conn.execute("UPDATE users SET username=?,password_hash=?,created_at=? WHERE id=?",
                    (username,generate_password_hash(password),datetime.utcnow().isoformat(timespec="seconds"),existing["id"]))
                action="Updated shared Staff login"
            else:
                conn.execute("INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)",
                    (username,generate_password_hash(password),"staff",datetime.utcnow().isoformat(timespec="seconds")))
                action="Created shared Staff login"
        log_activity(session["username"],action)
        flash("Shared Staff login saved successfully.")
        return redirect(url_for("admin"))
    with db() as conn:
        users=conn.execute("SELECT username,role FROM users ORDER BY role").fetchall()
        logs=conn.execute("SELECT username,action,created_at FROM activity_log ORDER BY id DESC LIMIT 20").fetchall()
    return render_template_string(ADMIN_PAGE,users=users,logs=logs)


@app.route("/", methods=["GET"])
def home():
    return render_template_string(PAGE)



@app.route("/maintenance", methods=["GET", "POST"])
def maintenance():
    default_heads = ["Service Charges", "Repairs & Maintenance Fund", "Sinking Fund", "Water Charges", "Lift Charges", "Other Charges"]
    context = dict(units=100, total_area=0, months=1, heads=None, amounts=None, methods=None,
                   flat_area=0, flat="", result=None, monthly_result=0, breakdown=[], default_heads=default_heads)
    if request.method == "POST":
        try:
            units = max(1, int(request.form.get("units", "1")))
            total_area = float(request.form.get("total_area", "0") or 0)
            months = int(request.form.get("months", "1"))
            flat_area = float(request.form.get("flat_area", "0") or 0)
            if months not in (1,3,6,12): raise ValueError("Billing period invalid.")
            if total_area < 0 or flat_area < 0: raise ValueError("Area cannot be negative.")
            breakdown=[]; monthly=0.0
            for i in range(6):
                head=request.form.get(f"head{i}", default_heads[i]).strip() or default_heads[i]
                amount=float(request.form.get(f"amount{i}", "0") or 0)
                method=request.form.get(f"method{i}", "equal")
                if amount < 0: raise ValueError("Charge amount cannot be negative.")
                if method=="area":
                    if total_area <= 0: raise ValueError("Area-based allocation requires total chargeable area.")
                    if flat_area > total_area: raise ValueError("Selected flat area cannot exceed total chargeable area.")
                    share=amount*flat_area/total_area
                    method_label="By area"
                else:
                    share=amount/units
                    method_label="Equally per unit"
                monthly += share
                breakdown.append(dict(head=head,method=method_label,monthly=share,period=share*months))
            context.update(units=units,total_area=total_area,months=months,flat_area=flat_area,
                flat=request.form.get("flat",""),heads=[request.form.get(f"head{i}",default_heads[i]) for i in range(6)],
                amounts=[float(request.form.get(f"amount{i}","0") or 0) for i in range(6)],
                methods=[request.form.get(f"method{i}","equal") for i in range(6)],
                result=monthly*months,monthly_result=monthly,breakdown=breakdown)
        except Exception as e:
            flash("Calculation error: "+str(e))
    return render_template_string(MAINTENANCE_PAGE, **context)

KNOWLEDGE = [
 dict(category="Maintenance", title="Maintenance charge calculation",
 answer="Maintenance bills may contain multiple approved charge heads. Allocation can differ by charge type; do not assume every charge is divided equally or by area. Use the Society's registered bye-laws, applicable Act/Rules, resolutions and supporting budgets to confirm each basis.",
 verify="Charge head, approved allocation basis, carpet/built-up area definition if relevant, General Body resolution, budget and applicable current rules."),
 dict(category="Funds", title="Sinking Fund",
 answer="A sinking fund is a reserve intended for major repairs, replacement or structural works as permitted by the applicable bye-laws and law. Keep a separate ledger, approved rate/basis, bank/investment records and utilisation approvals.",
 verify="Registered bye-law provision, rate/basis, AGM/GB approval, purpose, investment and utilisation conditions."),
 dict(category="Charges", title="Non-Occupancy Charges (NOC)",
 answer="Non-occupancy charges are subject to the applicable Maharashtra cooperative housing framework and the Society's registered bye-laws. Do not apply a charge solely on an informal rule or without verifying the relevant conditions and limits.",
 verify="Current Government orders/circulars, applicable bye-law clause, occupancy/letting facts and approved billing policy."),
 dict(category="Governance", title="AGM and General Body approvals",
 answer="Important Society decisions should be placed before the competent committee or General Body as required. Maintain agenda, notice, attendance, minutes, resolutions and supporting documents.",
 verify="Applicable Act/Rules, registered bye-laws, notice period, quorum, agenda and decision-making authority."),
 dict(category="Accounts", title="Audit and accounting records",
 answer="Maintain complete books, vouchers, invoices, bank statements, member ledgers, asset/fund ledgers and reconciliation records. Statutory audit and filing obligations depend on the applicable cooperative law and current directions.",
 verify="Current audit requirements, auditor appointment, statutory due dates, prescribed returns and Society records."),
 dict(category="Parking", title="Parking allotment and charges",
 answer="Parking use, allotment and charges should follow the registered bye-laws, approved Society policy and applicable law. Keep a transparent allotment register and record resolutions and waiting lists where applicable.",
 verify="Parking rules in registered bye-laws, sanctioned parking plan, allotment resolution and applicable court/government directions."),
 dict(category="Staff", title="Staff bonus and payments",
 answer="Staff bonus and other benefits depend on the worker's legal coverage, employment terms, applicable labour legislation, contracts and Society approvals. A Society should not assume that every vendor worker is its direct employee.",
 verify="Employer relationship, vendor contract, employee eligibility, applicable labour law, wage records and competent approval."),
 dict(category="Recovery", title="Defaulter notices and interest",
 answer="Member recovery steps should be supported by accurate ledger statements, invoice/charge basis, payment allocation and notices. Interest or penalty must be authorised by the applicable bye-laws/resolution and current law.",
 verify="Registered bye-law clause, valid resolution, rate, start date, payment credits, notice procedure and statutory recovery process."),
 dict(category="Tax", title="TDS and GST on vendor invoices",
 answer="TDS/GST treatment depends on the nature of supply, vendor status, thresholds, payment/credit timing and current tax law. Check each invoice and contract; do not apply one rate to every vendor.",
 verify="Current Income-tax provisions, GST law/notifications, PAN, invoice, threshold and service classification.")
]

@app.route("/society-help", methods=["GET"])
def society_help():
    q=request.args.get("q","").strip()
    results=[]
    if q:
        terms=[x.lower() for x in re.split(r"\\s+",q) if x]
        for item in KNOWLEDGE:
            hay=" ".join([item["category"],item["title"],item["answer"],item["verify"]]).lower()
            if any(t in hay for t in terms):
                results.append(item)
    else:
        results=KNOWLEDGE
    return render_template_string(HELP_PAGE,q=q,results=results)


@app.route("/analyze", methods=["POST"])
def analyze():
    excel = request.files.get("excel")
    if not excel or not excel.filename:
        flash("Please select an Excel file for analysis.")
        return redirect(url_for("home"))
    ext = os.path.splitext(secure_filename(excel.filename))[1].lower()
    if ext not in ALLOWED_EXCEL:
        flash("Please select an .xlsx or .xlsm Excel file.")
        return redirect(url_for("home"))
    try:
        grouped = read_rows(excel)
        rows = []
        for block, items in grouped.items():
            for unit, days, amount in items:
                try:
                    day_num = int(float(days)) if str(days).strip() else 0
                except Exception:
                    day_num = 0
                rows.append({"block": block, "unit": unit, "days": day_num, "amount": amount})
        total = sum(x["amount"] for x in rows)
        avg = total / len(rows) if rows else 0
        buckets = {"0–30 days": 0, "31–60 days": 0, "61–90 days": 0, "90+ days": 0}
        for x in rows:
            if x["days"] <= 30: buckets["0–30 days"] += 1
            elif x["days"] <= 60: buckets["31–60 days"] += 1
            elif x["days"] <= 90: buckets["61–90 days"] += 1
            else: buckets["90+ days"] += 1
        by_block = []
        for block, items in grouped.items():
            amount = sum(x[2] for x in items)
            by_block.append({"block": block, "flats": len(items), "amount": amount})
        by_block.sort(key=lambda x: x["amount"], reverse=True)
        top = sorted(rows, key=lambda x: x["amount"], reverse=True)[:10]
        import json
        return render_template_string(ANALYTICS_PAGE, rows=len(rows), blocks=len(grouped),
            total=fmt_money(total), avg=fmt_money(avg), buckets=buckets,
            by_block=by_block, top=top, chart_json=json.dumps(by_block))
    except Exception as e:
        flash("Error while analyzing the Excel file: " + str(e))
        return redirect(url_for("home"))


@app.route("/generate", methods=["POST"])
def generate():
    excel = request.files.get("excel")
    template = request.files.get("template")
    if not excel or not template or not excel.filename or not template.filename:
        flash("Please select both the Excel file and the Word template.")
        return redirect(url_for("home"))
    ex_ext = os.path.splitext(secure_filename(excel.filename))[1].lower()
    wd_ext = os.path.splitext(secure_filename(template.filename))[1].lower()
    if ex_ext not in ALLOWED_EXCEL or wd_ext not in ALLOWED_WORD:
        flash("Use an .xlsx/.xlsm file for Excel and a .docx file for the Word template.")
        return redirect(url_for("home"))
    try:
        result, blocks, records, total = generate_docx(excel, template)
        docx_filename = "Megapolis_Member_Circulation_Defaulter_Report.docx"

        # Package the generated Word report inside a ZIP archive for download.
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
            result.seek(0)
            archive.writestr(docx_filename, result.read())
        zip_buffer.seek(0)

        zip_filename = "Megapolis_Member_Circulation_Defaulter_Report.zip"
        return send_file(zip_buffer, as_attachment=True, download_name=zip_filename,
                         mimetype="application/zip")
    except Exception as e:
        flash("Error while generating the report: " + str(e))
        return redirect(url_for("home"))

if __name__ == "__main__":
    print("\\nOpen your browser at: http://127.0.0.1:5000")
    app.run(host=os.environ.get("HOST", "0.0.0.0"), port=int(os.environ.get("PORT", "5000")), debug=False)
