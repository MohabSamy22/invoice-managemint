import os, re, csv, sqlite3, hashlib, threading, subprocess, sys, shutil, calendar, secrets, hmac, socket, json, html
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    import pdfplumber
except ImportError: pdfplumber = None
# Pandas is only needed when exporting to Excel. Importing it during startup
# can make the desktop app appear frozen on some Python installations, so keep
# startup lightweight and load it only when Export Excel is used.
pd = None
# OCR dependencies are loaded on demand. Some installations spend a long time
# initializing pytesseract, which otherwise prevents the window from opening.
pytesseract = convert_from_path = None
try:
    import pymupdf as fitz
except ImportError:
    try: import fitz
    except ImportError: fitz = None

# When packaged as a one-file executable, PyInstaller runs code from a
# temporary extraction directory. Keep the database beside the executable so
# invoice data survives restarts and can travel with the app.
BASE = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
DB = BASE / 'invoices.db'
DEFAULT_ROOT = r'D:\Invoice' if Path(r'D:\Invoice').exists() else r'D:\Invoices'

def clean(s): return re.sub(r'\s+', ' ', s or '').strip()
def money(s):
    try: return float(str(s).replace(',', '').replace('$',''))
    except Exception: return None
def parse_date(s):
    if not s: return None
    for f in ('%m/%d/%Y','%m-%d-%Y','%m.%d.%Y','%m/%d/%y','%m-%d-%y','%m.%d.%y','%Y-%m-%d','%B %d, %Y','%b %d, %Y'):
        try: return datetime.strptime(s.strip(), f).date().isoformat()
        except ValueError: pass
    return None

def normalize_reference(value):
    """Return a safe invoice/reference token, or the UI sentinel."""
    if not value: return 'Not Detected'
    value = clean(str(value)).strip(' :#-.,;')
    if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9/_-]*', value): return 'Not Detected'
    if value.casefold() in {'date','number','no','invoice','reference','ref','account','phone','po','purchase'}: return 'Not Detected'
    return value

def folder_info(path, root):
    try: rel = Path(path).parent.relative_to(Path(root))
    except ValueError: rel = Path()
    parts = rel.parts
    if parts and parts[0].lower() == 'nursing notes':
        vendor = parts[1] if len(parts)>1 else 'Unknown Vendor'; category = 'Nursing Services'; tail = parts[2:]
    else:
        vendor = parts[0] if parts else 'Unknown Vendor'; category = 'General'; tail = parts[2:] if len(parts)>1 and parts[1].lower() in ('done','invoices','paid','unpaid') else parts[1:]
    month = tail[0] if tail else None
    folder_date = tail[1] if len(tail)>1 else None
    if parts and parts[0].lower() == 'nursing notes': category = 'Nursing Services'
    return vendor, category, month, folder_date

def extract(path, root, force_ocr=False, enhanced_ocr=False):
    global pytesseract, convert_from_path
    text = ''
    try: nursing_file = any(part.lower() == 'nursing notes' for part in Path(path).relative_to(Path(root)).parts)
    except Exception: nursing_file = False
    try:
        # Some large scanned Nursing Notes PDFs contain malformed embedded images
        # that make pdfplumber hang. OCR page 1 directly for those files.
        use_text_parser = pdfplumber and not force_ocr and not (nursing_file and os.path.getsize(path) > 5 * 1024 * 1024)
        if use_text_parser:
            # PyMuPDF is substantially safer on malformed PDFs. Prefer it
            # before pdfplumber, which can hang while decoding damaged pages.
            if fitz:
                doc=fitz.open(path)
                parts=Path(path).relative_to(Path(root)).parts
                pages=doc[:1] if parts and parts[0].lower()=='nursing notes' else doc
                text='\n'.join((p.get_text() or '') for p in pages)
                doc.close()
            if len(clean(text)) < 20:
                with pdfplumber.open(path) as pdf:
                    parts=Path(path).relative_to(Path(root)).parts
                    pages=pdf.pages[:1] if parts and parts[0].lower()=='nursing notes' else pdf.pages
                    text='\n'.join((p.extract_text() or '') for p in pages)
    except Exception: pass
    if force_ocr or len(clean(text)) < 20:
        if pytesseract is None or convert_from_path is None:
            try:
                import pytesseract as _pytesseract
                from pdf2image import convert_from_path as _convert_from_path
                from PIL import ImageOps, ImageEnhance, ImageFilter
                pytesseract, convert_from_path = _pytesseract, _convert_from_path
                if os.path.exists(r'C:\Program Files\Tesseract-OCR\tesseract.exe'):
                    pytesseract.pytesseract.tesseract_cmd=r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            except ImportError:
                pytesseract = convert_from_path = None
        if pytesseract and convert_from_path:
            try:
                nursing = nursing_file
                if fitz:
                    doc=fitz.open(path); page=doc.load_page(0); scale=(300 if enhanced_ocr else 180)/72; pix=page.get_pixmap(matrix=fitz.Matrix(scale,scale),alpha=False); from PIL import Image; import io; images=[Image.open(io.BytesIO(pix.tobytes('png')))]
                else:
                    images = convert_from_path(path, dpi=300 if enhanced_ocr else 180, first_page=1, last_page=1)
                ocr_text=[]
                for im in images:
                    if enhanced_ocr:
                        im=ImageOps.grayscale(im); im=ImageEnhance.Contrast(im).enhance(2.0); im=im.filter(ImageFilter.SHARPEN)
                        ocr_text.extend([pytesseract.image_to_string(im,config='--psm 6'),pytesseract.image_to_string(im,config='--psm 11')])
                    else: ocr_text.append(pytesseract.image_to_string(im))
                text='\n'.join(ocr_text)
            except Exception: pass
    vendor, category, month, folder_date = folder_info(path, root)
    # Autola invoices are sometimes stored in a short drive-named folder or
    # OCR renders the vendor as "Atuola". Normalize both cases so the UI and
    # download destination use the real vendor name.
    vendor_hint = clean(f'{text} {Path(path).stem}').lower()
    if (str(vendor).strip().lower() in ('d', 'atuola', 'autola') or
            re.search(r'\b(?:autola|atuola|atola)\b', vendor_hint, re.I)):
        vendor = 'Autola'
    def find(pattern):
        m = re.search(pattern, text, re.I); return clean(m.group(1)) if m else None
    # Accept the label variants used by vendor PDFs and by OCR output
    # (e.g. "Reference Number", "Ref #", and "Invoice No.").
    ref = find(r'(?:invoice\s*)?(?:ref(?:erence)?|reference\s+number)\s*(?:number|no\.?|#)?\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-_/]*)')
    inv = find(r'(?:invoice\s*(?:number|no\.?|#)|inv\.?\s*(?:number|no\.?|#))\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-_/]*)')
    if inv and inv.upper() in ('DATE','NO','NUMBER'): inv = find(r'INVOICE\s*#\s*([0-9][A-Z0-9\-_/]+)')
    ref = ref or inv
    # CuraScript credit/debit memos often expose a generic QYT value in the
    # surrounding reference area. The memo number is the correct reference.
    memo_ref = find(r'(?:credit|debit)\s+memo\s*(?:(?:reference|ref(?:erence)?|number|no\.?)\s*)?(?:#|number|no\.?)?\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-_/]+)')
    if not memo_ref:
        memo_ref = find(r'\b(?:CM|DM)\s*(?:#|number|no\.?)?\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-_/]+)')
    if memo_ref:
        ref = memo_ref
    document_type = 'Credit Memo' if re.search(r'credit\s+memo', text, re.I) else ('Debit Memo' if re.search(r'debit\s+memo', text, re.I) else 'Bill')
    # CuraScript debit memos identify the transaction with a QYT number.
    # Prefer that value over OCR text captured from the surrounding label.
    qyt_ref = find(r'\bQYT\s*(?:#|number|no\.?)?\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-_/]+)')
    if document_type == 'Debit Memo' and qyt_ref:
        ref = qyt_ref
    terms = find(r'\b(NET\s*\d{1,3}|Due\s+on\s+receipt)\b')
    memo_date_label = 'credit' if document_type == 'Credit Memo' else ('debit' if document_type == 'Debit Memo' else '')
    memo_date = find(rf'{memo_date_label}\s+memo\s+date\.?\s*[:#-]?\s*(\d{{1,2}}[/-]\d{{1,2}}[/-]\d{{2,4}}|[A-Za-z]+\s+\d{{1,2}},?\s+\d{{4}})') if memo_date_label else None
    # Do not let the generic date fallback start at the word "date" inside
    # "due date"; CuraScript documents often print due date before invoice date.
    date_token=r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2}|[A-Za-z]+\s+\d{1,2},?\s+\d{4})'
    d = memo_date or find(r'(?:invoice\s+date|(?<!due\s)\bdate)\s*[:#-]?\s*'+date_token)
    check = find(r'TOTAL\s+IF\s+PAID\s+BY\s+CHECK\s+OR\s+EFT[\s\S]{0,120}?\$([\d,]+\.\d{2})')
    aps_total = re.search(r'INVOICE\s+TOTAL\s*:\s*[^\n]*\$([\d,]+\.\d{2})', text, re.I)
    labeled_total = re.search(r'(?:DOCUMENT\s+TOTAL|TOTAL\s+AMT\s*\(?(?:USD)?\)?|TOTAL\s*\(USD\)|NET\s+AMOUNT|TOTAL\s+AMOUNT\s*\(USD\)|BALANCE\s+DUE|AMOUNT\s+DUE\s+BY)\s*[:$]*\s*\$?([\d,]+\.\d{2})', text, re.I)
    document_total = re.search(r'(?:DOCUMENT\s+TOTAL|TOTAL\s+AMT|TOTAL\s+AMOUNT|NET\s+AMOUNT)[\s\S]{0,100}?([\d,]+\.\d{2})', text, re.I)
    amount = check or (aps_total.group(1) if aps_total else None) or (labeled_total.group(1) if labeled_total else None) or (document_total.group(1) if document_total else None) or find(r'(?:total\s*(?:amount|due)?|amount\s*due(?:\s+by)?|balance\s*due|invoice\s+total)\s*[:$-]*\s*\$?([\d,]+\.\d{2})')
    if vendor.lower() == 'autola':
        vals=re.findall(r'\$\s*([\d,]+\.\d{2})',text); amount=vals[-1] if vals else None
    # Credit memos must be recorded as negative amounts. Apply this only
    # after the document total has been extracted.
    if document_type == 'Credit Memo' and amount is not None:
        memo_amount=money(amount)
        amount = -abs(memo_amount) if memo_amount is not None else amount
    if not ref:
        file_ref=re.search(r'(?:invoice|inv)[ _#-]*(\d{5,})\b',Path(path).stem,re.I)
        ref=file_ref.group(1) if file_ref else ref
    if not ref and vendor.lower() == 'autola':
        autola_ref=re.search(r'\b(20\d{2}-\d{5,})\b', text+' '+Path(path).stem)
        ref=autola_ref.group(1) if autola_ref else ref
    # OCR often turns CuraScript labels into short words such as "lect",
    # "Sales", or "Debit". These are not invoice references; prefer the
    # invoice number or filename fallback instead.
    junk_refs={'lect','sales','debit','credit','date','number','no','invoice','reference','ref','rigerated','refrigerated'}
    if ref and str(ref).strip().casefold() in junk_refs:
        ref=inv if inv and str(inv).strip().casefold() not in junk_refs else None
    if not ref and vendor.lower() == 'curascript':
        file_ref=re.search(r'(?:invoice|inv)[ _#-]*(\d{5,})\b',Path(path).stem,re.I)
        ref=file_ref.group(1) if file_ref else ref
    if vendor.lower() == 'amerisource' and ref and str(ref).strip().casefold() in junk_refs:
        ref=inv if inv and str(inv).strip().casefold() not in junk_refs else None
    if not ref and vendor.lower() == 'amerisource':
        file_ref=re.search(r'(?:invoice|inv)[ _#-]*(\d{5,})\b',Path(path).stem,re.I)
        ref=file_ref.group(1) if file_ref else ref
    ref = normalize_reference(ref)
    date_source = 'invoice' if parse_date(d) else ''
    inv_date = parse_date(d)
    if not inv_date and folder_date:
        for x in (folder_date, f'{folder_date}-{datetime.now().year}'):
            inv_date = parse_date(x)
            if inv_date: date_source='folder'
    if not inv_date:
        try:
            inv_date=datetime.fromtimestamp(Path(path).stat().st_mtime).date().isoformat(); date_source='file timestamp'
        except OSError: pass
    due = find(r'(?:due\s+date|due)\s*[:#-]?\s*'+date_token)
    due = parse_date(due)
    if not due and terms and inv_date:
        m = re.search(r'(\d+)', terms); due = (date.fromisoformat(inv_date)+timedelta(days=int(m.group(1)))).isoformat() if m else None
    # Some vendor templates omit the printed terms. Only infer them when the
    # document did not provide terms: first from invoice/due dates, then from
    # the known vendor defaults.
    if not terms and inv_date and due:
        days=(date.fromisoformat(due)-date.fromisoformat(inv_date)).days
        if 0 <= days <= 365: terms=f'Net {days}'
    if not terms:
        vendor_terms={
            'make invoices': 'Net 0',
            'make invoice': 'Net 0',
            'autola': 'Net 30',
            'treasure care home health': 'Net 15',
            'treasure care': 'Net 15',
            'aps': 'Net 15',
        }
        terms=vendor_terms.get(str(vendor).strip().casefold())
    if terms and not due and inv_date:
        m=re.search(r'\d+',terms)
        if m: due=(date.fromisoformat(inv_date)+timedelta(days=int(m.group(0)))).isoformat()
    confidence = 0.9 if text and amount and (ref or inv) else (0.55 if text else 0.2)
    lower=(text+' '+str(vendor)).lower()
    normalized_address = re.sub(r'[^a-z0-9]', '', lower)
    # Company ownership rules: Amerisource is always NHW. CuraScript can serve
    # either company, so Lotus/Nature Health signals must take precedence.
    if nursing_file: company='SPX'
    elif str(vendor).lower() == 'amerisource': company='NHW'
    elif 'lotus' in lower or 'nature health' in lower: company='NHW'
    elif 'infusion' in lower or 'curascript' in lower: company='SPX'
    else: company='Unknown'
    # Lotus invoices do not always print the city as a separate site label.
    # Match the known street/ZIP combinations as well as the city names so
    # OCR can still populate the site when line breaks or punctuation vary.
    site_patterns=[('LEAGUE CITY',r'league\s*city|77573|1108\s+gulf\s+fwy\s+s\s+(?:ste|suite)\s*200'),('Lotus El Paso',r'(?:lotus\s*(?:health\s*)?)?(?:el\s*paso|elpaso)|1699\s+north\s+zaragoza|79936'),('Lotus Allen',r'(?:lotus\s*(?:health\s*)?)?allen|997\s+raintree|75013'),('Lotus Nacogdoches',r'(?:lotus\s*(?:health\s*)?)?nacogdoches|4928\s+north\s+st|75965'),('Lotus Pearland',r'(?:lotus\s*(?:health\s*)?)?pearland|pyramids\s+infusion|3569\s+business\s+center|77584'),('Lotus Plano',r'(?:lotus\s*(?:health\s*)?)?plano|4701\s+west\s+plano|75093'),('Lotus ATL',r'lotus\s*(?:health\s*)?\(?atl'),('Lotus San Antonio',r'lotus\s*(?:health\s*)?\(?san\s*antonio'),('PFWC',r'\bpfwc\b')]
    league_city = ('leaguecity' in normalized_address and '77573' in normalized_address) or ('1108gulffwys200leaguecitytx77573' in normalized_address) or ('1108gulffwysuite200leaguecitytx77573' in normalized_address)
    detected_site='LEAGUE CITY' if league_city else next((name for name,pattern in site_patterns if re.search(pattern,lower)), 'Lotus Health' if ('lotus' in lower or 'nature health' in lower) else '')
    site=detected_site if company=='NHW' else ('Infusion' if company=='SPX' else '')
    return dict(vendor=vendor, reference=ref, invoice_number=inv, document_type=document_type, amount=money(amount), invoice_date=inv_date, date_source=date_source, terms=terms, due_date=due, category=category, company=company, site=site, confidence=confidence, text=text[:5000])

class App:
    def __init__(self, root):
        self.root=root; root.title('Invoice Desk'); root.geometry('1250x760'); root.minsize(1000,600)
        style=ttk.Style(root)
        try: style.theme_use('clam')
        except tk.TclError: pass
        style.configure('TButton',font=('Segoe UI',10),padding=(8,5)); style.configure('Treeview',font=('Segoe UI',10),rowheight=30,background='#FFFFFF',fieldbackground='#FFFFFF',borderwidth=0); style.configure('Treeview.Heading',font=('Segoe UI',10,'bold'),background='#16324F',foreground='white',padding=7); style.map('Treeview',background=[('selected','#2563EB')],foreground=[('selected','white')]); style.configure('TLabel',font=('Segoe UI',10)); style.configure('Title.TLabel',font=('Segoe UI',24,'bold'),foreground='#16324F'); style.configure('Sidebar.TFrame',background='#16324F')
        self.db=sqlite3.connect(DB, check_same_thread=False); self.db.row_factory=sqlite3.Row; self.setup_db(); self.current_user='admin'; self.current_role='admin'; self.session_timeout=1800; self.http_server=None; self.root_path=tk.StringVar(value=self.setting('root',DEFAULT_ROOT)); self.build(); self.toggle_server(); self.refresh(); self._last_folder_signature=self.folder_signature(self.root_path.get()); self.watch()
    def setup_db(self):
        self.db.executescript('''CREATE TABLE IF NOT EXISTS invoices(id INTEGER PRIMARY KEY, vendor TEXT, reference TEXT, invoice_number TEXT, amount REAL, invoice_date TEXT, terms TEXT, due_date TEXT, category TEXT, path TEXT UNIQUE, filename TEXT, added TEXT, updated TEXT, status TEXT, confidence REAL, hash TEXT, duplicate INTEGER DEFAULT 0, ignored INTEGER DEFAULT 0, company TEXT DEFAULT 'Unknown', site TEXT DEFAULT ''); CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY, invoice_id INTEGER, changed TEXT, at TEXT, details TEXT, username TEXT DEFAULT 'system', action_type TEXT DEFAULT 'system', item TEXT, previous_value TEXT, new_value TEXT); CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT); CREATE TABLE IF NOT EXISTS excluded_vendors(vendor TEXT PRIMARY KEY, removed_at TEXT); CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE NOT NULL, email TEXT, password_hash TEXT NOT NULL, salt TEXT NOT NULL, active INTEGER DEFAULT 1, role TEXT DEFAULT 'user', created_at TEXT, updated_at TEXT); CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY, user_id INTEGER, created_at TEXT, last_seen TEXT, expires_at TEXT);''')
        for col, typ in (('username','TEXT DEFAULT \'system\''),('action_type','TEXT DEFAULT \'system\''),('item','TEXT'),('previous_value','TEXT'),('new_value','TEXT')):
            try: self.db.execute(f'ALTER TABLE audit ADD COLUMN {col} {typ}')
            except sqlite3.OperationalError: pass
        if not self.db.execute('SELECT 1 FROM users LIMIT 1').fetchone():
            salt=secrets.token_hex(16); pwd=self.setting('admin_password','admin'); digest=hashlib.pbkdf2_hmac('sha256',pwd.encode(),salt.encode(),240000).hex()
            self.db.execute("INSERT INTO users(username,email,password_hash,salt,active,role,created_at,updated_at) VALUES(?,?,?,?,1,'admin',?,?)",('admin','',digest,salt,datetime.now().isoformat(timespec='seconds'),datetime.now().isoformat(timespec='seconds')))
            self.db.commit()
        for col in ('duplicate','ignored','company','site','payment_status','paid_date','notes','document_type','recorded_at','date_source'):
            try: self.db.execute(f'ALTER TABLE invoices ADD COLUMN {col} INTEGER DEFAULT 0')
            except sqlite3.OperationalError: pass
        # Existing invoices are already part of the recorded history. New imports
        # are inserted with recorded=0 and appear in the New Invoices tab.
        try: self.db.execute('ALTER TABLE invoices ADD COLUMN recorded INTEGER DEFAULT 1')
        except sqlite3.OperationalError: pass
        # The recorded field was added after the first imports. Treat invoices
        # imported today as new once, so they can be reviewed and recorded.
        if not self.setting('recorded_migration_today'):
            today=date.today().isoformat()
            self.db.execute("UPDATE invoices SET recorded=0 WHERE substr(added,1,10)=?",(today,))
            self.save_setting('recorded_migration_today',today)
        self.db.execute("UPDATE invoices SET recorded_at=COALESCE(NULLIF(recorded_at,0),updated,added) WHERE COALESCE(recorded,0)=1")
        self.db.execute("UPDATE invoices SET payment_status='Unpaid' WHERE payment_status IS NULL OR payment_status=''" )
        self.normalize_existing()
        self.db.commit()
    def setting(self,k,d=''): 
        r=self.db.execute('SELECT value FROM settings WHERE key=?',(k,)).fetchone(); return r['value'] if r else d
    def save_setting(self,k,v): self.db.execute('INSERT OR REPLACE INTO settings VALUES(?,?)',(k,v)); self.db.commit()
    def audit_log(self, action, item='', previous='', new='', details='', username=None, invoice_id=None):
        now=datetime.now().isoformat(timespec='seconds'); self.db.execute('INSERT INTO audit(invoice_id,changed,at,details,username,action_type,item,previous_value,new_value) VALUES(?,?,?,?,?,?,?,?,?)',(invoice_id,action,now,details,username or self.current_user,action,item,previous,new)); self.db.commit()
    def password_ok(self, username, password):
        r=self.db.execute('SELECT * FROM users WHERE username=? AND active=1',(username,)).fetchone()
        ok=bool(r and hmac.compare_digest(hashlib.pbkdf2_hmac('sha256',password.encode(),r['salt'].encode(),240000).hex(),r['password_hash']))
        self.audit_log('login' if ok else 'failed login',username,'','', 'Network login attempt', username=username)
        return ok
    def local_ip(self):
        try:
            s=socket.socket(socket.AF_INET,socket.SOCK_DGRAM); s.connect(('8.8.8.8',80)); ip=s.getsockname()[0]; s.close(); return ip
        except Exception: return '127.0.0.1'
    def toggle_server(self):
        if self.http_server: self.http_server.shutdown(); self.http_server=None
        if self.setting('network_access','0')=='1':
            host=self.local_ip(); port=int(self.setting('network_port','8765'))
            app=self
            class Handler(BaseHTTPRequestHandler):
                def do_GET(self):
                    from urllib.parse import urlparse,parse_qs
                    import html
                    from http.cookies import SimpleCookie
                    cookie=SimpleCookie(self.headers.get('Cookie','')); token=cookie.get('invoice_session').value if cookie.get('invoice_session') else ''
                    session=app.db.execute('SELECT sessions.*,users.username,users.role FROM sessions JOIN users ON users.id=sessions.user_id WHERE sessions.token=? AND sessions.expires_at>? AND users.active=1',(token,datetime.now().isoformat(timespec='seconds'))).fetchone()
                    if not session:
                        return self.html('<h1>Invoice Desk</h1><form method="post"><label>Username <input name="username"></label><br><label>Password <input name="password" type="password"></label><br><button>Sign in</button></form>')
                    app.db.execute('UPDATE sessions SET last_seen=? WHERE token=?',(datetime.now().isoformat(timespec='seconds'),token)); app.db.commit()
                    path=urlparse(self.path).path; qs=parse_qs(urlparse(self.path).query)
                    if path in ('/settings','/users','/audit') and session['role']!='admin': return self.html('<h1>Access denied</h1><p>Administrator privileges are required.</p>',403)
                    if path in ('/','/dashboard'): return self.html(app.web_dashboard(session,qs))
                    if path=='/export':
                        import io
                        from openpyxl import Workbook
                        month=qs.get('month',[''])[0]; company=qs.get('company',[''])[0]; site=qs.get('site',[''])[0]; sql='SELECT reference,vendor,amount,company,invoice_date,terms,due_date FROM invoices WHERE COALESCE(duplicate,0)=0 AND COALESCE(ignored,0)=0 AND lower(trim(vendor)) NOT IN (SELECT lower(trim(vendor)) FROM excluded_vendors)'; args=[]
                        if month and month!='All months': sql+=' AND substr(invoice_date,1,7)=?'; args.append(month)
                        if company and company!='All companies': sql+=' AND company=?'; args.append(company)
                        if site and site!='All sites': sql+=' AND site=?'; args.append(site)
                        wb=Workbook(); ws=wb.active; ws.title='Invoices'; ws.append(['Reference','Vendor','Amount','Company','Invoice Date','Terms','Due Date'])
                        for r in app.db.execute(sql+' ORDER BY invoice_date DESC',args): ws.append([r['reference'],r['vendor'],r['amount'],r['company'],r['invoice_date'],r['terms'],r['due_date']])
                        out=io.BytesIO(); wb.save(out); self.send_response(200); self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); self.send_header('Content-Disposition','attachment; filename=Invoice_Report.xlsx'); self.end_headers(); self.wfile.write(out.getvalue()); return
                    if path=='/invoices': return self.html(app.web_invoices(session,qs))
                    if path=='/vendors':
                        vendor=qs.get('vendor',[''])[0]; df=qs.get('date_from',[''])[0]; dt=qs.get('date_to',[''])[0]; sql='SELECT vendor,COUNT(*) count,SUM(amount) total,MAX(invoice_date) latest FROM invoices WHERE 1=1'; args=[]
                        if vendor: sql+=' AND vendor LIKE ?'; args.append('%'+vendor+'%')
                        if df: sql+=' AND invoice_date>=?'; args.append(df)
                        if dt: sql+=' AND invoice_date<=?'; args.append(dt)
                        sql+=' GROUP BY vendor ORDER BY vendor'; rows=[[html.escape(str(r['vendor'] or '')),str(r['count']),f'${r["total"] or 0:,.2f}',str(r['latest'] or '')] for r in app.db.execute(sql,args)]
                        return self.html(app.web_simple(session,'Vendors','/vendors',rows,('Vendor','Invoice count','Total amount','Latest invoice date'),app.web_filters('/vendors',qs,[('vendor','Vendor name','text'),('date_from','Date from','date'),('date_to','Date to','date')])))
                    if path=='/review':
                        ref=qs.get('reference',[''])[0]; vendor=qs.get('vendor',[''])[0]; sql="SELECT reference,vendor,invoice_date,status FROM invoices WHERE status='Needs Review'"; args=[]
                        if ref: sql+=' AND reference LIKE ?'; args.append('%'+ref+'%')
                        if vendor: sql+=' AND vendor LIKE ?'; args.append('%'+vendor+'%')
                        rows=[[html.escape(str(r['reference'] or '')),html.escape(str(r['vendor'] or '')),str(r['invoice_date'] or ''),str(r['status'] or '')] for r in app.db.execute(sql+' ORDER BY id DESC',args)]
                        return self.html(app.web_simple(session,'Review Queue','/review',rows,('Reference','Vendor','Date','Status'),app.web_filters('/review',qs,[('reference','Reference','text'),('vendor','Vendor','text'),('date_from','Date from','date'),('date_to','Date to','date')])))
                    if path=='/duplicates':
                        ref=qs.get('reference',[''])[0]; vendor=qs.get('vendor',[''])[0]; sql='SELECT reference,vendor,amount,invoice_date FROM invoices WHERE duplicate=1 AND ignored=0'; args=[]
                        if ref: sql+=' AND reference LIKE ?'; args.append('%'+ref+'%')
                        if vendor: sql+=' AND vendor LIKE ?'; args.append('%'+vendor+'%')
                        rows=[[html.escape(str(r['reference'] or '')),html.escape(str(r['vendor'] or '')),str(r['amount'] or ''),str(r['invoice_date'] or '')] for r in app.db.execute(sql+' ORDER BY id DESC',args)]
                        return self.html(app.web_simple(session,'Duplicates','/duplicates',rows,('Reference','Vendor','Amount','Date'),app.web_filters('/duplicates',qs,[('reference','Reference','text'),('vendor','Vendor','text'),('date_from','Date from','date'),('date_to','Date to','date'),('amount','Amount','number')])))
                    if path=='/reconciliation': return self.html(app.web_simple(session,'Reconciliation','/reconciliation',[],('Status','Details')))
                    if path=='/downloads': return self.html(app.web_simple(session,'Download PDFs','/downloads',[],('Available action','Details')))
                    if path=='/users':
                        us=qs.get('username',[''])[0]; em=qs.get('email',[''])[0]; st=qs.get('status',['']); sql='SELECT username,email,active,role FROM users WHERE 1=1'; args=[]
                        if us: sql+=' AND username LIKE ?'; args.append('%'+us+'%')
                        if em: sql+=' AND email LIKE ?'; args.append('%'+em+'%')
                        if st and st[0] in ('active','inactive'): sql+=' AND active=?'; args.append(1 if st[0]=='active' else 0)
                        rows=[[html.escape(str(r['username'])),html.escape(str(r['email'] or '')),('Active' if r['active'] else 'Inactive'),str(r['role'])] for r in app.db.execute(sql+' ORDER BY username',args)]
                        return self.html(app.web_simple(session,'User Management','/users',rows,('Username','Email','Status','Role'),app.web_filters('/users',qs,[('username','Username','text'),('email','Email','text'),('status','Status (active/inactive)','text')])))
                    if path=='/audit':
                        sql='SELECT at,username,action_type,changed,item,previous_value,new_value FROM audit WHERE 1=1'; args=[]
                        for key,col in [('user','username'),('action','action_type')]:
                            if qs.get(key,[''])[0]: sql+=f' AND {col} LIKE ?'; args.append('%'+qs[key][0]+'%')
                        if qs.get('date_from',[''])[0]: sql+=' AND at>=?'; args.append(qs['date_from'][0])
                        if qs.get('date_to',[''])[0]: sql+=' AND at<=?'; args.append(qs['date_to'][0]+'T23:59:59')
                        rows=[[str(r['at']),html.escape(str(r['username'] or 'system')),str(r['action_type'] or r['changed']),html.escape(str(r['item'] or '')),html.escape(str(r['previous_value'] or '')),html.escape(str(r['new_value'] or ''))] for r in app.db.execute(sql+' ORDER BY id DESC LIMIT 500',args)]
                        return self.html(app.web_simple(session,'Audit Trail','/audit',rows,('Date/time','User','Action','Item','Previous value','New value'),app.web_filters('/audit',qs,[('user','User','text'),('date_from','Date from','date'),('date_to','Date to','date'),('action','Action','text'),('search','Search','text')])))
                    if path not in ('/new','/edit'):
                        return self.html('<h1>Page not found</h1><p>The requested page does not exist.</p><a href="/dashboard">Return to Dashboard</a>',404)
                    if path=='/edit':
                        r=app.db.execute('SELECT * FROM invoices WHERE id=?',(qs.get('id',[''])[0],)).fetchone()
                        body='<h2>Edit invoice</h2><form method="post" action="/edit"><input type="hidden" name="id" value="'+str(r['id'])+'">Vendor <input name="vendor" value="'+html.escape(r['vendor'] or '')+'"><br>Reference <input name="reference" value="'+html.escape(r['reference'] or '')+'"><br>Amount <input name="amount" value="'+str(r['amount'] or '')+'"><br>Date <input name="invoice_date" value="'+str(r['invoice_date'] or '')+'"><br><button>Save</button></form><p><a href="/">Back</a></p>' if r else '<p>Invoice not found</p>'
                    else:
                        rows=app.db.execute('SELECT id,reference,vendor,amount,invoice_date,status FROM invoices ORDER BY id DESC LIMIT 500').fetchall()
                        body='<h1>Invoice Desk</h1><p>Signed in as '+html.escape(session['username'])+'</p><nav><a href="/">Dashboard</a> | <a href="/invoices">Invoices</a> | <a href="/new">Add Invoice</a> | <a href="/vendors">Vendors</a>' + ( ' | <a href="/audit">Audit Trail</a> | <a href="/users">User Management</a> | <a href="/settings">Settings</a>' if session['role']=='admin' else '') + '</nav><hr><table border="1" cellpadding="6"><tr><th>Reference</th><th>Vendor</th><th>Amount</th><th>Date</th><th>Status</th><th>Actions</th></tr>' + ''.join(f'<tr><td>{html.escape(str(r["reference"] or ""))}</td><td>{html.escape(str(r["vendor"] or ""))}</td><td>{r["amount"] or ""}</td><td>{r["invoice_date"] or ""}</td><td>{r["status"] or ""}</td><td><a href="/edit?id={r["id"]}">Edit</a> <form style="display:inline" method="post" action="/delete"><input type="hidden" name="id" value="{r["id"]}"><button>Delete</button></form></td></tr>' for r in rows) + '</table>'
                        if path=='/new': body='<h2>Add invoice</h2><form method="post" action="/new">Vendor <input name="vendor" required><br>Reference <input name="reference"><br>Amount <input name="amount"><br>Date <input name="invoice_date"><br><button>Add invoice</button></form><p><a href="/">Back</a></p>'
                    return self.html(body)
                def do_POST(self):
                    from urllib.parse import parse_qs
                    length=int(self.headers.get('Content-Length','0')); data=parse_qs(self.rfile.read(length).decode()); u=data.get('username',[''])[0]; p=data.get('password',[''])[0]
                    from http.cookies import SimpleCookie
                    c=SimpleCookie(self.headers.get('Cookie','')); token=c.get('invoice_session').value if c.get('invoice_session') else ''; s=app.db.execute('SELECT sessions.*,users.username,users.role FROM sessions JOIN users ON users.id=sessions.user_id WHERE token=? AND expires_at>? AND users.active=1',(token,datetime.now().isoformat(timespec='seconds'))).fetchone()
                    if s and self.path in ('/new','/edit','/delete'):
                        if self.path=='/new':
                            now=datetime.now().isoformat(timespec='seconds'); cur=app.db.execute("INSERT INTO invoices(vendor,reference,amount,invoice_date,status,added,updated) VALUES(?,?,?,?,?,?,?)",(data.get('vendor',[''])[0],data.get('reference',['Not Detected'])[0],float(data.get('amount',['0'])[0] or 0),data.get('invoice_date',[''])[0],'Imported',now,now)); app.audit_log('invoice addition',data.get('reference',[''])[0],'','created',username=s['username'],invoice_id=cur.lastrowid)
                        elif self.path=='/edit':
                            iid=data.get('id',[''])[0]; old=app.db.execute('SELECT * FROM invoices WHERE id=?',(iid,)).fetchone(); app.db.execute('UPDATE invoices SET vendor=?,reference=?,amount=?,invoice_date=?,updated=? WHERE id=?',(data.get('vendor',[''])[0],data.get('reference',[''])[0],float(data.get('amount',['0'])[0] or 0),data.get('invoice_date',[''])[0],datetime.now().isoformat(timespec='seconds'),iid)); app.audit_log('invoice edit',iid,str(dict(old)) if old else '',str(dict(data)),username=s['username'],invoice_id=iid)
                        else:
                            iid=data.get('id',[''])[0]; old=app.db.execute('SELECT reference FROM invoices WHERE id=?',(iid,)).fetchone(); app.db.execute('DELETE FROM invoices WHERE id=?',(iid,)); app.audit_log('invoice deletion',iid,old['reference'] if old else '','','',username=s['username'],invoice_id=iid)
                        app.db.commit(); self.send_response(303); self.send_header('Location','/'); self.end_headers(); return
                    if not app.password_ok(u,p): return self.html('<h1>Sign in failed</h1><p>Invalid or inactive account.</p><a href="/">Try again</a>',401)
                    user=app.db.execute('SELECT * FROM users WHERE username=?',(u,)).fetchone(); token=secrets.token_urlsafe(32); now=datetime.now(); expires=now+timedelta(seconds=app.session_timeout); app.db.execute('INSERT INTO sessions VALUES(?,?,?,?,?)',(token,user['id'],now.isoformat(timespec='seconds'),now.isoformat(timespec='seconds'),expires.isoformat(timespec='seconds'))); app.db.commit()
                    self.send_response(303); self.send_header('Location','/'); self.send_header('Set-Cookie',f'invoice_session={token}; HttpOnly; SameSite=Strict'); self.end_headers()
                def html(self,body,status=200):
                    css='''*{box-sizing:border-box}html,body{margin:0;min-height:100%;background:#f5f7fb;color:#172b4d;font:14px Segoe UI,Arial,sans-serif}.app-shell{min-height:100vh}.sidebar{position:fixed;left:0;top:0;width:240px;height:100vh;padding:28px 16px;background:#102a43;color:white;display:flex;flex-direction:column;gap:5px;box-shadow:3px 0 12px #ccd3dd;z-index:10}.sidebar:before{content:"INVOICE\\A DESK";white-space:pre;color:white;font-size:22px;font-weight:700;line-height:1.25;margin:0 5px 28px}.sidebar a{display:block;padding:11px 14px;border-radius:7px;color:#d9e2ec;text-decoration:none;font-size:14px;font-weight:500}.sidebar a:hover,.sidebar a:first-of-type{background:#1f4e79;color:white}.main-content{margin-left:240px;width:calc(100% - 240px);min-width:0;padding:32px;overflow-x:hidden}.main-content h1{margin:0 0 6px;font-size:30px;color:#102a43}.main-content>p{margin:0 0 22px;color:#667085}.main-content>hr{border:0;border-top:1px solid #e2e8f0}.table-card{width:100%;max-width:1240px;background:white;border:1px solid #e2e8f0;border-radius:10px;box-shadow:0 2px 8px #dce2ea;overflow:hidden}.table-responsive{width:100%;overflow-x:auto}table{min-width:720px;width:100%;border:0;border-collapse:separate;border-spacing:0;background:white}th{background:#edf2f7;color:#486581;text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}td,th{padding:13px 15px;border-bottom:1px solid #edf2f7}td:nth-child(3){text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}tr:last-child td{border-bottom:0}td form{display:inline;padding:0;box-shadow:none;background:transparent}td a{display:inline-block;border:1px solid #0e7490;border-radius:5px;padding:6px 10px;text-decoration:none}td button{background:#dc3545;padding:6px 10px}input{padding:10px;border:1px solid #cbd5e1;border-radius:6px;margin:8px 4px}button{background:#0ea5a8;color:white;border:0;border-radius:6px;padding:10px 16px;font-weight:600;cursor:pointer}button:hover{background:#0f8c90}form{background:white;padding:22px;border-radius:10px;box-shadow:0 1px 4px #dce2ea;max-width:560px}a{color:#0e7490}@media(max-width:900px){.sidebar{width:200px}.main-content{margin-left:200px;width:calc(100% - 200px);padding:24px}.sidebar a{font-size:13px;padding:10px}}'''
                    if '<div class="app-shell">' in body:
                        pass
                    elif '<nav>' in body and '</nav>' in body:
                        before, navpart=body.split('<nav>',1); navinner, after=navpart.split('</nav>',1)
                        body='<div class="app-shell"><aside class="sidebar">'+navinner+'</aside><main class="main-content">'+before+after+'</main></div>'
                    else:
                        body='<main class="main-content">'+body+'</main>'
                    if '<table' in body:
                        body=body.replace('<table','<div class="table-card"><div class="table-responsive"><table',1).replace('</table>','</table></div></div>',1)
                    self.send_response(status); self.send_header('Content-Type','text/html; charset=utf-8'); self.end_headers(); self.wfile.write(('<!doctype html><meta name="viewport" content="width=device-width"><style>'+css+'</style>'+body).encode())
                def log_message(self,*args): pass
            self.http_server=ThreadingHTTPServer((host,port),Handler); threading.Thread(target=self.http_server.serve_forever,daemon=True).start()
    def admin_only(self):
        if self.current_role!='admin': messagebox.showerror('Administrator only','This section is available only to the administrator.'); return False
        return True
    def web_page(self, title, content, user, active):
        links=[('Dashboard','/dashboard'),('New Invoices','/new'),('Invoices','/invoices'),('Download PDFs','/downloads'),('Reconciliation','/reconciliation'),('Duplicates','/duplicates'),('Review Queue','/review'),('Vendors','/vendors')]
        if user['role']=='admin': links += [('Settings','/settings'),('User Management','/users'),('Audit Trail','/audit')]
        nav=''.join(f'<a class="{"active" if path==active else ""}" href="{path}">{label}</a>' for label,path in links)
        return f'<div class="app-shell"><aside class="sidebar">{nav}</aside><main class="main-content"><h1>{title}</h1><p>Invoice Desk &bull; Signed in as {user["username"]}</p>{content}</main></div>'
    def web_table(self, headers, rows):
        return '<div class="table-card"><div class="table-responsive"><table><thead><tr>'+''.join(f'<th>{h}</th>' for h in headers)+'</tr></thead><tbody>'+''.join('<tr>'+''.join(f'<td>{v}</td>' for v in row)+'</tr>' for row in rows)+'</tbody></table></div></div>'
    def web_dashboard(self,user,q=None):
        q=q or {}; month=q.get('month',['All months'])[0]; company=q.get('company',['All companies'])[0]; site=q.get('site',['All sites'])[0]; where=' WHERE COALESCE(ignored,0)=0 AND lower(trim(vendor)) NOT IN (SELECT lower(trim(vendor)) FROM excluded_vendors)'; args=[]
        if month!='All months': where+=' AND substr(invoice_date,1,7)=?'; args.append(month)
        if company!='All companies': where+=' AND company=?'; args.append(company)
        if site!='All sites': where+=' AND site=?'; args.append(site)
        total=self.db.execute('SELECT COUNT(*) n,COALESCE(SUM(amount),0) total FROM invoices'+where,args).fetchone(); month_n=self.db.execute("SELECT COUNT(*) n FROM invoices WHERE substr(invoice_date,1,7)=?",(datetime.now().strftime('%Y-%m'),)).fetchone()['n']; review=self.db.execute("SELECT COUNT(*) n FROM invoices WHERE status='Needs Review'"+where.replace(' WHERE',' AND'),args).fetchone()['n']; vendors=self.db.execute('SELECT COUNT(DISTINCT vendor) n FROM invoices'+where,args).fetchone()['n']; dup=self.db.execute('SELECT COUNT(*) n FROM invoices WHERE duplicate=1'+where.replace(' WHERE',' AND'),args).fetchone()['n']; cards=''.join(f'<div class="stat"><span>{label}</span><strong>{value}</strong></div>' for label,value in [('Total invoices',total['n']),('Total amount',f'${total["total"]:,.2f}'),('This month',month_n),('Needs Review',review),('Vendors',vendors),('Duplicates',dup)])
        rows=[[str(r['reference'] or ''),str(r['vendor'] or ''),f'${r["amount"]:,.2f}' if r['amount'] is not None else '',str(r['invoice_date'] or ''),str(r['status'] or '')] for r in self.db.execute('SELECT reference,vendor,amount,invoice_date,status FROM invoices ORDER BY id DESC LIMIT 8')]
        months=['All months']+[r['month'] for r in self.db.execute("SELECT DISTINCT substr(invoice_date,1,7) month FROM invoices WHERE invoice_date IS NOT NULL ORDER BY month DESC")]; companies=['All companies']+[r['company'] for r in self.db.execute("SELECT DISTINCT company FROM invoices WHERE company IS NOT NULL AND company<>'' ORDER BY company")]; sites=['All sites']+[r['site'] for r in self.db.execute("SELECT DISTINCT site FROM invoices WHERE site IS NOT NULL AND site<>'' ORDER BY site")]; f=f'<form class="filter-card" method="get" action="/dashboard"><label>Month<select name="month">'+''.join(f'<option selected>{html.escape(x)}</option>' if x==month else f'<option>{html.escape(x)}</option>' for x in months)+f'</select></label><label>Company<select name="company">'+''.join(f'<option selected>{html.escape(x)}</option>' if x==company else f'<option>{html.escape(x)}</option>' for x in companies)+f'</select></label><label>Site<select name="site">'+''.join(f'<option selected>{html.escape(x)}</option>' if x==site else f'<option>{html.escape(x)}</option>' for x in sites)+f'</select></label><button>Apply Filters</button> <a class="clear" href="/dashboard">Clear Filters</a> <a class="export" href="/export?month={html.escape(month)}&company={html.escape(company)}&site={html.escape(site)}">Export Excel</a></form>'
        return self.web_page('Dashboard',f+f'<div class="stats">{cards}</div><h2>Recent invoices</h2>{self.web_table(("Reference","Vendor","Amount","Date","Status"),rows)}',user,'/dashboard')
    def web_filters(self, action, q, fields):
        parts=[]
        for name,label,typ in fields:
            value=html.escape(q.get(name,[''])[0],quote=True); parts.append(f'<label>{label}<input name="{name}" type="{typ}" value="{value}"></label>')
        return f'<form class="filter-card" method="get" action="{action}">'+''.join(parts)+'<button>Apply Filters</button> <a class="clear" href="{action}">Clear Filters</a></form>'
    def web_invoices(self,user,q=None,active='/invoices'):
        q=q or {}; sql='SELECT id,reference,vendor,amount,invoice_date,status FROM invoices WHERE 1=1'; args=[]
        for key,clause in [('reference','reference LIKE ?'),('vendor','vendor=?'),('status','status=?')]:
            value=q.get(key,[''])[0]
            if value: sql+=' AND '+clause; args.append('%'+value+'%' if key=='reference' else value)
        if q.get('date_from',[''])[0]: sql+=' AND invoice_date>=?'; args.append(q['date_from'][0])
        if q.get('date_to',[''])[0]: sql+=' AND invoice_date<=?'; args.append(q['date_to'][0])
        for key,op in [('min_amount','>='),('max_amount','<=')]:
            try:
                if q.get(key,[''])[0]: sql+=f' AND amount{op}?'; args.append(float(q[key][0]))
            except ValueError: pass
        rows=self.db.execute(sql+' ORDER BY id DESC LIMIT 500',args).fetchall(); trs=[]
        for r in rows:
            status=r['status'] or ''; badge=f'<span class="badge">{status}</span>'; actions=f'<a href="/edit?id={r["id"]}">Edit</a> <form class="inline" method="post" action="/delete"><input type="hidden" name="id" value="{r["id"]}"><button class="danger">Delete</button></form>'
            trs.append((html.escape(str(r['reference'] or '')),html.escape(str(r['vendor'] or '')),f'${r["amount"]:,.2f}' if r['amount'] is not None else '',r['invoice_date'] or '',badge,actions))
        vendors=[r['vendor'] for r in self.db.execute('SELECT DISTINCT vendor FROM invoices WHERE vendor IS NOT NULL ORDER BY vendor')]; statuses=[r['status'] for r in self.db.execute('SELECT DISTINCT status FROM invoices WHERE status IS NOT NULL ORDER BY status')]
        filter_html=self.web_filters('/invoices',q,[('reference','Reference','text'),('vendor','Vendor','text'),('date_from','Date from','date'),('date_to','Date to','date'),('status','Status','text'),('min_amount','Min amount','number'),('max_amount','Max amount','number')])
        return self.web_page('Invoices',filter_html+self.web_table(('Reference','Vendor','Amount','Date','Status','Actions'),trs),user,active)
    def web_simple(self,user,title,active,query,headers,filters=''):
        return self.web_page(title,filters+self.web_table(headers,query),user,active)
    def normalize_existing(self):
        self.db.execute("UPDATE invoices SET category='Nursing Services' WHERE lower(path) LIKE '%prodigy%' OR lower(path) LIKE '%nursing notes%' OR category='Infusion Services'")
        self.db.execute("UPDATE invoices SET company='SPX' WHERE lower(path) LIKE '%prodigy%' OR lower(path) LIKE '%nursing notes%'")
        self.db.execute("UPDATE invoices SET company='NHW' WHERE lower(path) LIKE '%amerisource%'")
        self.db.execute("UPDATE invoices SET reference='Not Detected' WHERE reference IS NULL OR trim(reference)=''")
        self.db.execute("UPDATE invoices SET status='Needs Review' WHERE invoice_date IS NULL AND COALESCE(status,'') NOT IN ('Resolved','Paid')")
        self.db.commit()
    def build(self):
        style=ttk.Style(self.root); self.root.configure(bg='#F3F6FA');
        style.configure('TFrame',background='#F3F6FA'); style.configure('TLabel',background='#F3F6FA',foreground='#243B53',font=('Segoe UI',10)); style.configure('TButton',background='#E6EEF5',foreground='#16324F',font=('Segoe UI',10,'bold'),padding=(10,7)); style.map('TButton',background=[('active','#BDEFEA')]); style.configure('TEntry',padding=6); style.configure('TCombobox',padding=5); style.configure('TLabelframe',background='#F3F6FA',foreground='#16324F'); style.configure('TLabelframe.Label',background='#F3F6FA',foreground='#16324F',font=('Segoe UI',10,'bold')); style.configure('Card.TLabelframe',background='white',bordercolor='#E2E8F0',relief='solid',borderwidth=1); style.configure('Card.TLabelframe.Label',background='white',foreground='#16324F',font=('Segoe UI',10,'bold')); style.configure('Title.TLabel',background='#F3F6FA',foreground='#16324F',font=('Segoe UI',24,'bold'))
        side=tk.Frame(self.root,bg='#102A43',width=235); side.pack(side='left',fill='y'); side.pack_propagate(False); tk.Label(side,text='INVOICE',bg='#102A43',fg='#7DD3FC',font=('Segoe UI',11,'bold')).pack(anchor='w',padx=22,pady=(26,0)); tk.Label(side,text='DESK',bg='#102A43',fg='white',font=('Segoe UI',24,'bold')).pack(anchor='w',padx=20); tk.Label(side,text='Finance operations',bg='#102A43',fg='#9FB3C8',font=('Segoe UI',9)).pack(anchor='w',padx=22,pady=(0,28)); self.buttons={}
        labels=('Dashboard','New Invoices','Unrecorded Invoices','Invoices','Download PDFs','Reconciliation','Duplicates','Review Queue','Vendors','Settings','User Management','Audit Trail')
        for label in labels:
            self.buttons[label]=tk.Button(side,text=label,anchor='w',padx=22,relief='flat',bd=0,bg='#102A43',fg='#D9E2EC',activebackground='#1F4E79',activeforeground='white',font=('Segoe UI',10),command=lambda x=label:self.page(x)); self.buttons[label].pack(fill='x',pady=1,ipady=8)
        tk.Frame(side,bg='#486581',height=1).pack(fill='x',padx=20,pady=20); tk.Button(side,text='＋  Scan Now',anchor='w',padx=18,relief='flat',bd=0,bg='#0EA5A8',fg='white',activebackground='#14B8A6',font=('Segoe UI',10,'bold'),command=self.scan).pack(fill='x',padx=14,pady=4,ipady=8); tk.Button(side,text='↻  Rescan All',anchor='w',padx=18,relief='flat',bd=0,bg='#1F4E79',fg='white',activebackground='#2563EB',font=('Segoe UI',10),command=lambda:self.scan(True)).pack(fill='x',padx=14,pady=4,ipady=8); tk.Button(side,text='⇩  Export Excel',anchor='w',padx=18,relief='flat',bd=0,bg='#1F4E79',fg='white',activebackground='#2563EB',font=('Segoe UI',10),command=self.export).pack(fill='x',padx=14,pady=4,ipady=8)
        self.main=tk.Frame(self.root,bg='#F3F6FA',padx=26,pady=22); self.main.pack(side='left',fill='both',expand=True)
        self.scan_status=tk.StringVar(value='Ready'); self.scan_progress=ttk.Progressbar(self.root,mode='determinate',length=320); self.scan_progress.pack(side='bottom',fill='x',padx=18,pady=(0,5)); ttk.Label(self.root,textvariable=self.scan_status,anchor='w').pack(side='bottom',fill='x',padx=18)
        self.page('Dashboard')
    def clear(self):
        for w in self.main.winfo_children(): w.destroy()
    def page(self,name):
        if name in ('Settings','User Management','Audit Trail') and not self.admin_only(): return
        for label,button in getattr(self,'buttons',{}).items(): button.configure(bg='#1F4E79' if label==name else '#102A43',fg='white' if label==name else '#D9E2EC')
        self.clear(); ttk.Label(self.main,text=name,style='Title.TLabel').pack(anchor='w',pady=(0,3)); ttk.Label(self.main,text='Local invoice intelligence • original files stay in place',foreground='#667085').pack(anchor='w',pady=(0,15)); getattr(self,'page_'+name.lower().replace(' ','_'))()
    def table(self, rows, cols):
        ttk.Label(self.main,text=f'{len(rows):,} item(s) displayed',foreground='#475569').pack(anchor='w',pady=(0,5))
        wrap=ttk.Frame(self.main); wrap.pack(fill='both',expand=True); tv=ttk.Treeview(wrap,columns=cols,show='headings',selectmode='extended'); vs=ttk.Scrollbar(wrap,orient='vertical',command=tv.yview); hs=ttk.Scrollbar(wrap,orient='horizontal',command=tv.xview); tv.configure(yscrollcommand=vs.set,xscrollcommand=hs.set); tv.grid(row=0,column=0,sticky='nsew'); vs.grid(row=0,column=1,sticky='ns'); hs.grid(row=1,column=0,sticky='ew'); wrap.rowconfigure(0,weight=1); wrap.columnconfigure(0,weight=1)
        for c in cols: tv.heading(c,text=c.replace('_',' ').title()); tv.column(c,width=145)
        for i,r in enumerate(rows):
            display=[]
            for c in cols:
                value=r[c] if r[c] is not None else ''
                if c == 'amount' and value != '':
                    try:
                        number=float(value)
                        try: doc_type=r['document_type']
                        except (KeyError,IndexError): doc_type=''
                        if str(doc_type or '').casefold() == 'credit memo': number=-abs(number)
                        value=f'{number:,.2f}'
                    except (TypeError,ValueError): pass
                display.append(value)
            tv.insert('', 'end', values=display,tags=('odd' if i%2 else 'even',))
        tv.tag_configure('odd',background='#F7FAFC'); tv.tag_configure('even',background='#FFFFFF')
        tv.bind('<Double-1>',lambda e:self.open_row(tv,cols))
        return tv
    def open_row(self,tv,cols):
        if 'path' not in cols: return
        item=tv.focus(); vals=tv.item(item,'values'); path=vals[cols.index('path')] if item else ''
        if path and Path(path).exists(): os.startfile(path)
    def page_dashboard(self):
        toolbar=ttk.Frame(self.main); toolbar.pack(fill='x',pady=(0,10)); ttk.Label(toolbar,text='Reporting month:').pack(side='left'); ttk.Button(toolbar,text='Refresh Data',command=self.refresh_dashboard).pack(side='right')
        months=['All months']+[r['month'] for r in self.db.execute("SELECT DISTINCT substr(invoice_date,1,7) month FROM invoices WHERE invoice_date IS NOT NULL ORDER BY month DESC").fetchall()]
        self.month_filter=tk.StringVar(value=getattr(self,'dashboard_month','All months')); combo=ttk.Combobox(toolbar,textvariable=self.month_filter,values=months,state='readonly',width=18); combo.pack(side='left',padx=8); combo.bind('<<ComboboxSelected>>',lambda e:(setattr(self,'dashboard_month',self.month_filter.get()),self.page('Dashboard')))
        companies=['All companies']+[r['company'] for r in self.db.execute("SELECT DISTINCT company FROM invoices WHERE company IS NOT NULL AND company<>'' ORDER BY company").fetchall()]; sites=['All sites']+[r['site'] for r in self.db.execute("SELECT DISTINCT site FROM invoices WHERE site IS NOT NULL AND site<>'' ORDER BY site").fetchall()]
        self.dashboard_company=tk.StringVar(value=getattr(self,'dashboard_company_value','All companies')); self.dashboard_site=tk.StringVar(value=getattr(self,'dashboard_site_value','All sites'))
        ttk.Label(toolbar,text='Company:').pack(side='left',padx=(12,3)); company_combo=ttk.Combobox(toolbar,textvariable=self.dashboard_company,values=companies,state='readonly',width=16); company_combo.pack(side='left',padx=4); company_combo.bind('<<ComboboxSelected>>',lambda e:(setattr(self,'dashboard_company_value',self.dashboard_company.get()),self.page('Dashboard')))
        ttk.Label(toolbar,text='Site:').pack(side='left',padx=(8,3)); site_combo=ttk.Combobox(toolbar,textvariable=self.dashboard_site,values=sites,state='readonly',width=22); site_combo.pack(side='left',padx=4); site_combo.bind('<<ComboboxSelected>>',lambda e:(setattr(self,'dashboard_site_value',self.dashboard_site.get()),self.page('Dashboard')))
        rows=self.dashboard_rows(); total=sum((r['amount'] or 0) for r in rows); review=sum(1 for r in rows if (r['confidence'] or 0)<.7); nhw=sum((r['amount'] or 0) for r in rows if r['company']=='NHW'); spx=sum((r['amount'] or 0) for r in rows if r['company']=='SPX'); unpaid=sum((r['amount'] or 0) for r in rows if str(r['payment_status'] or 'Unpaid').casefold()!='paid'); added_today=sum(1 for r in rows if str(r['added'] or '')[:10]==date.today().isoformat()); dup_count=self.db.execute("SELECT COUNT(*) n FROM invoices WHERE duplicate=1 AND ignored=0 AND lower(trim(vendor)) NOT IN (SELECT lower(trim(vendor)) FROM excluded_vendors)").fetchone()['n']
        cards=ttk.Frame(self.main); cards.pack(fill='x',pady=10)
        ttk.Label(self.main,text=f'Filtered data: {len(rows):,} invoices  |  Unpaid: ${unpaid:,.2f}  |  Added today: {added_today}',foreground='#475569').pack(anchor='w')
        counts={'NHW':sum(1 for r in rows if r['company']=='NHW'),'SPX':sum(1 for r in rows if r['company']=='SPX')}
        for title,val,cmd in [('Total Spending',f'${total:,.2f}',None),('NHW',f'${nhw:,.2f}  •  {counts["NHW"]} invoices',lambda:self.company_view('NHW')),('SPX',f'${spx:,.2f}  •  {counts["SPX"]} invoices',lambda:self.company_view('SPX')),('Invoices',len(rows),None),('Needs Review',review,None)]:
            f=ttk.LabelFrame(cards,text=title,padding=14,style='Card.TLabelframe'); f.pack(side='left',fill='x',expand=True,padx=5)
            w=ttk.Button(f,text=str(val),command=cmd) if cmd else ttk.Label(f,text=str(val))
            w.pack(fill='x');
        charts=ttk.Frame(self.main); charts.pack(fill='both',expand=True,pady=(20,0)); self.chart_monthly(charts,rows); self.chart_companies(charts,rows)
        ttk.Label(self.main,text=f'Invoices  |  Probable duplicates flagged: {dup_count:,}',font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(25,8)); ttk.Button(self.main,text='Export to Excel',command=lambda:self.export(rows)).pack(anchor='e'); self.table(rows[::-1],('vendor','reference','amount','invoice_date','due_date','status'))
    def refresh_dashboard(self):
        self.db.execute('PRAGMA optimize'); self.normalize_existing(); self.db.commit(); self.page('Dashboard')
    def dashboard_rows(self):
        where='WHERE COALESCE(duplicate,0)=0 AND COALESCE(ignored,0)=0 AND lower(trim(vendor)) NOT IN (SELECT lower(trim(vendor)) FROM excluded_vendors)'; args=[]
        if self.month_filter.get()!='All months': where+=' AND substr(invoice_date,1,7)=?'; args.append(self.month_filter.get())
        if self.dashboard_company.get()!='All companies': where+=' AND company=?'; args.append(self.dashboard_company.get())
        if self.dashboard_site.get()!='All sites': where+=' AND site=?'; args.append(self.dashboard_site.get())
        return self.db.execute('SELECT * FROM invoices '+where,args).fetchall()
    def page_new_invoices(self):
        """Show recently imported invoices that have not been recorded yet."""
        toolbar=ttk.Frame(self.main); toolbar.pack(fill='x',pady=(0,10))
        ttk.Label(toolbar,text='Newly imported invoices waiting to be recorded').pack(side='left')
        ttk.Button(toolbar,text='Scan New Invoices',command=self.scan).pack(side='right',padx=4)
        ttk.Button(toolbar,text='Rescan Existing',command=self.rescan_existing_invoices).pack(side='right',padx=4)
        ttk.Button(toolbar,text='Refresh',command=lambda:self.page('New Invoices')).pack(side='right',padx=4)
        ttk.Button(toolbar,text='Mark All Recorded',command=self.mark_all_recorded).pack(side='right',padx=4)
        ttk.Button(toolbar,text='Mark Selected Recorded',command=self.mark_selected_recorded).pack(side='right',padx=4)
        ttk.Label(toolbar,text='Company:').pack(side='left',padx=(18,4))
        self.new_company_filter=tk.StringVar(value=getattr(self,'new_company_value','All'))
        company_combo=ttk.Combobox(toolbar,textvariable=self.new_company_filter,values=['All','SPX','NHW'],state='readonly',width=9)
        company_combo.pack(side='left'); company_combo.bind('<<ComboboxSelected>>',lambda e:(setattr(self,'new_company_value',self.new_company_filter.get()),self.page('New Invoices')))
        sql="SELECT id,path,vendor,company,COALESCE(document_type,'Bill') AS document_type,reference,amount,invoice_date,'Not Recorded' AS status,added FROM invoices WHERE COALESCE(recorded,0)=0 AND substr(added,1,10)=?"
        args=[date.today().isoformat()]
        if self.new_company_filter.get()!='All': sql+=' AND company=?'; args.append(self.new_company_filter.get())
        rows=self.db.execute(sql+' ORDER BY id DESC',args).fetchall()
        ttk.Label(self.main,text=f'{len(rows):,} new invoice(s) waiting to be recorded',foreground='#475569').pack(anchor='w',pady=(0,5))
        self.new_invoice_tree=self.table(rows,('id','path','vendor','company','document_type','reference','amount','invoice_date','status','added'))
        self.new_invoice_tree.bind('<Double-1>',self.open_new_invoice)

    def page_unrecorded_invoices(self):
        """Show all imported invoices that have not been recorded yet."""
        toolbar=tk.Frame(self.main,bg='#F3F6FA'); toolbar.pack(fill='x',pady=(0,10))
        ttk.Label(toolbar,text='All invoices waiting to be recorded').pack(side='left')
        ttk.Button(toolbar,text='Refresh',command=lambda:self.page('Unrecorded Invoices')).pack(side='right',padx=4)
        ttk.Button(toolbar,text='Mark All Recorded',command=self.mark_all_unrecorded).pack(side='right',padx=4)
        ttk.Button(toolbar,text='Mark Selected Recorded',command=self.mark_selected_unrecorded).pack(side='right',padx=4)
        filters=ttk.LabelFrame(self.main,text='Unrecorded invoice filters',padding=10); filters.pack(fill='x',pady=(0,10))
        self.unrecorded_from=tk.StringVar(value=getattr(self,'unrecorded_from_value',''))
        self.unrecorded_to=tk.StringVar(value=getattr(self,'unrecorded_to_value',''))
        self.unrecorded_type=tk.StringVar(value=getattr(self,'unrecorded_type_value','All types'))
        ttk.Label(filters,text='Date From').grid(row=0,column=0,padx=(0,4),pady=3,sticky='w')
        ttk.Entry(filters,textvariable=self.unrecorded_from,width=14).grid(row=0,column=1,padx=(0,12),pady=3)
        ttk.Label(filters,text='Date To').grid(row=0,column=2,padx=(0,4),pady=3,sticky='w')
        ttk.Entry(filters,textvariable=self.unrecorded_to,width=14).grid(row=0,column=3,padx=(0,12),pady=3)
        ttk.Label(filters,text='Document Type').grid(row=0,column=4,padx=(0,4),pady=3,sticky='w')
        ttk.Combobox(filters,textvariable=self.unrecorded_type,values=['All types','Bill','Debit Memo','Credit Memo'],state='readonly',width=16).grid(row=0,column=5,padx=(0,12),pady=3)
        ttk.Button(filters,text='Apply Filters',command=self.apply_unrecorded_filters).grid(row=0,column=6,padx=3)
        ttk.Button(filters,text='Clear',command=self.clear_unrecorded_filters).grid(row=0,column=7,padx=3)
        sql="SELECT id,path,vendor,company,COALESCE(document_type,'Bill') AS document_type,reference,amount,invoice_date,'Not Recorded' AS status,added FROM invoices WHERE COALESCE(recorded,0)=0"
        args=[]
        if self.unrecorded_from.get(): sql+=' AND invoice_date>=?'; args.append(self.unrecorded_from.get())
        if self.unrecorded_to.get(): sql+=' AND invoice_date<=?'; args.append(self.unrecorded_to.get())
        if self.unrecorded_type.get()!='All types': sql+=' AND COALESCE(document_type,\'Bill\')=?'; args.append(self.unrecorded_type.get())
        rows=self.db.execute(sql+' ORDER BY id DESC',args).fetchall()
        ttk.Label(self.main,text=f'{len(rows):,} unrecorded invoice(s)',foreground='#475569').pack(anchor='w',pady=(0,5))
        self.new_invoice_tree=self.table(rows,('id','path','vendor','company','document_type','reference','amount','invoice_date','status','added'))
        self.new_invoice_tree._invoice_columns=('id','path','vendor','company','document_type','reference','amount','invoice_date','status','added')
        self.new_invoice_tree.bind('<ButtonRelease-1>',self.copy_unrecorded_reference,add='+')
        self.new_invoice_tree.bind('<Double-1>',self.open_new_invoice)

    def copy_unrecorded_reference(self,event):
        item=self.new_invoice_tree.identify_row(event.y)
        if not item: return
        self.new_invoice_tree.selection_set(item); self.new_invoice_tree.focus(item)
        values=self.new_invoice_tree.item(item,'values')
        reference=values[5] if len(values)>5 else ''
        if reference:
            self.root.clipboard_clear(); self.root.clipboard_append(str(reference))
            self.scan_status.set(f'Reference copied: {reference}')

    def apply_unrecorded_filters(self):
        self.unrecorded_from_value=self.unrecorded_from.get().strip()
        self.unrecorded_to_value=self.unrecorded_to.get().strip()
        self.unrecorded_type_value=self.unrecorded_type.get()
        self.page('Unrecorded Invoices')

    def clear_unrecorded_filters(self):
        self.unrecorded_from_value=''; self.unrecorded_to_value=''; self.unrecorded_type_value='All types'
        self.page('Unrecorded Invoices')

    def mark_selected_unrecorded(self):
        items=self.new_invoice_tree.selection()
        if not items: return messagebox.showwarning('Select invoices','Select one or more invoices first.')
        ids=[self.new_invoice_tree.item(item,'values')[0] for item in items]
        stamp=datetime.now().isoformat(timespec='seconds')
        self.db.executemany('UPDATE invoices SET recorded=1,recorded_at=?,updated=? WHERE id=?',[(stamp,stamp,i) for i in ids]); self.db.commit(); self.page('Unrecorded Invoices')

    def mark_all_unrecorded(self):
        count=self.db.execute('SELECT COUNT(*) FROM invoices WHERE COALESCE(recorded,0)=0').fetchone()[0]
        if not count: return messagebox.showinfo('Unrecorded Invoices','There are no unrecorded invoices.')
        stamp=datetime.now().isoformat(timespec='seconds')
        self.db.execute('UPDATE invoices SET recorded=1,recorded_at=?,updated=? WHERE COALESCE(recorded,0)=0',(stamp,stamp)); self.db.commit(); self.page('Unrecorded Invoices')

    def page_download_pdfs(self):
        """Preview PDFs downloaded today and route them into the Nursing Notes tree."""
        bar=ttk.Frame(self.main); bar.pack(fill='x',pady=(0,10))
        ttk.Label(bar,text=r'Source: D:\Downloads  •  PDFs modified today only').pack(side='left')
        ttk.Button(bar,text='Scan Downloads',command=self.scan_downloads).pack(side='right',padx=4)
        ttk.Button(bar,text='Edit Location',command=self.edit_download_location).pack(side='right',padx=4)
        ttk.Button(bar,text='Remove Selected',command=self.remove_downloads).pack(side='right',padx=4)
        ttk.Button(bar,text='Move Selected',command=self.move_downloads).pack(side='right',padx=4)
        ttk.Button(bar,text='Move All',command=lambda:self.move_downloads(all_rows=True)).pack(side='right',padx=4)
        ttk.Label(self.main,text='Double-click a destination to edit it before moving.').pack(anchor='w',pady=(0,5))
        self.download_progress_text=tk.StringVar(value='Ready')
        ttk.Label(self.main,textvariable=self.download_progress_text).pack(anchor='w')
        self.download_progress=ttk.Progressbar(self.main,mode='determinate',length=500)
        self.download_progress.pack(fill='x',pady=(2,8))
        self.download_tree=self.table([],('filename','vendor','destination','status'))
        self.download_tree._download_columns=('filename','vendor','destination','status')
        self.download_tree.bind('<Double-1>',lambda e:self.edit_download_location())
        self.scan_downloads()

    def remove_downloads(self):
        """Remove items from this preview only; do not delete or move the PDFs."""
        items=self.download_tree.selection()
        if not items: return messagebox.showwarning('Select PDFs','Select one or more PDFs to remove from the preview.')
        names={self.download_tree.item(item,'values')[0] for item in items}
        self.download_rows=[row for row in self.download_rows if row['filename'] not in names]
        for item in items:
            self.download_tree.delete(item)

    def download_destination(self, vendor, when):
        vendor=re.sub(r'[<>:"/\\|?*]','',clean(vendor) or 'Unknown Vendor') or 'Unknown Vendor'
        month=calendar.month_abbr[when.month]
        # These vendors use their own top-level invoice folders. All other
        # vendors remain under the Nursing notes structure.
        direct_vendors={'amerisource':'Amerisource','biocare':'Biocare','curascript':'Curascript','prodigy':'Prodigy'}
        folder_name=direct_vendors.get(vendor.casefold(), vendor)
        base=Path(self.root_path.get())/folder_name if vendor.casefold() in direct_vendors else Path(self.root_path.get())/'Nursing notes'/folder_name
        # Keep the existing 2026 layout. For later years, add the year level so
        # annual folders cannot collide with the original Done/Month folders.
        if when.year == date.today().year:
            return base/'Done'/month/f'{when.month}-{when.day}'
        return base/str(when.year)/month/f'{when.month}-{when.day}'

    def scan_downloads(self):
        if getattr(self, '_download_scanning', False):
            return
        self._download_scanning=True
        source=Path(r'D:\Downloads'); today=date.today(); root_path=self.root_path.get()
        if hasattr(self,'download_progress'):
            self.download_progress['value']=0; self.download_progress['maximum']=1
            self.download_progress_text.set('Preparing download scan…')
        threading.Thread(target=self._scan_downloads_worker,args=(source,today,root_path),daemon=True).start()

    def _scan_downloads_worker(self, source, today, root_path):
        rows=[]
        allowed_dates={today, today-timedelta(days=1)}
        files=sorted((p for p in source.glob('*.pdf')
                      if datetime.fromtimestamp(p.stat().st_mtime).date() in allowed_dates),
                     key=lambda x:x.stat().st_mtime) if source.exists() else []
        self.root.after(0, lambda n=len(files): self._download_progress_start(n))
        if source.exists():
            for index,p in enumerate(files,1):
                # Update before extraction: OCR can take a while on a scanned
                # PDF, so the user can see which file is currently being read.
                self.root.after(0, lambda i=index,n=len(files),name=p.name: self._download_progress_update(i-1,n,name+' (opening)'))
                try:
                    # Read only lightweight embedded text from page 1. Do not
                    # run OCR here, since scanned/malformed PDFs can hang.
                    info={}
                    pdf_text=''
                    if fitz:
                        try:
                            doc=fitz.open(str(p))
                            if doc.page_count:
                                pdf_text=doc.load_page(0).get_text().casefold()
                            doc.close()
                        except Exception:
                            pdf_text=''
                    known_vendors=(('amerisource','Amerisource'),('biocare','Biocare'),('curascript','Curascript'),('prodigy','Prodigy'))
                    vendor=next((label for needle,label in known_vendors if needle in pdf_text or needle in p.stem.casefold()), None)
                    if not vendor:
                        nursing_root=Path(root_path)/'Nursing notes'
                        existing_vendors=[d.name for d in nursing_root.iterdir() if d.is_dir()] if nursing_root.exists() else []
                        haystack=(pdf_text+' '+p.stem.casefold())
                        vendor=next((name for name in sorted(existing_vendors,key=len,reverse=True) if name.casefold() in haystack), None)
                    if not vendor and info.get('vendor') not in (None,'','Unknown Vendor','Unknown'):
                        vendor=info.get('vendor')
                    if not vendor:
                        # Do not mistake a filename such as Invoice 90165329
                        # for a vendor; leave it editable in the preview.
                        vendor='Unknown Vendor'
                    rows.append({'filename':p.name,'vendor':vendor,'destination':str(self.download_destination_for_root(vendor,today,root_path)),'status':'Ready','source':str(p)})
                except Exception as e:
                    rows.append({'filename':p.name,'vendor':'Unknown Vendor','destination':str(self.download_destination_for_root('Unknown Vendor',today,root_path)),'status':f'Could not read: {e}','source':str(p)})
                self.root.after(0, lambda i=index,n=len(files),name=p.name: self._download_progress_update(i,n,name))
        self.root.after(0, lambda: self._finish_download_scan(rows))

    def _download_progress_start(self, total):
        if hasattr(self,'download_progress'):
            self.download_progress['maximum']=max(total,1); self.download_progress['value']=0
            self.download_progress_text.set(f'Scanning 0 of {total} PDFs…')

    def _download_progress_update(self, done, total, name):
        if hasattr(self,'download_progress'):
            self.download_progress['maximum']=max(total,1); self.download_progress['value']=done
            self.download_progress_text.set(f'Scanning {done} of {total}: {name}')

    def download_destination_for_root(self, vendor, when, root_path):
        vendor=re.sub(r'[<>:"/\\|?*]','',clean(vendor) or 'Unknown Vendor') or 'Unknown Vendor'
        month=calendar.month_abbr[when.month]
        direct_vendors={'amerisource':'Amerisource','biocare':'Biocare','curascript':'Curascript','prodigy':'Prodigy'}
        folder_name=direct_vendors.get(vendor.casefold(), vendor)
        base=Path(root_path)/folder_name if vendor.casefold() in direct_vendors else Path(root_path)/'Nursing notes'/folder_name
        return base/'Done'/month/f'{when.month}-{when.day}' if when.year == date.today().year else base/str(when.year)/month/f'{when.month}-{when.day}'

    def _finish_download_scan(self, rows):
        self._download_scanning=False
        if hasattr(self,'download_progress'):
            self.download_progress['value']=self.download_progress['maximum']
            self.download_progress_text.set(f'Scan complete — {len(rows)} PDF(s) found')
        self.download_rows=rows
        if hasattr(self,'download_tree') and self.download_tree.winfo_exists():
            # Destroy the complete table wrapper, including its scrollbars.
            # Destroying only the Treeview leaves a scrollbar callback pointing
            # at a widget that no longer exists (Tkinter TclError on xview).
            self.download_tree.master.destroy()
        self.download_tree=self.table(rows,('filename','vendor','destination','status'))
        self.download_tree._download_columns=('filename','vendor','destination','status')
        self.download_tree.bind('<Double-1>',lambda e:self.edit_download_location())

    def edit_download_location(self):
        items=self.download_tree.selection()
        if len(items)!=1: return messagebox.showwarning('Select one PDF','Select one PDF to edit its destination.')
        item=items[0]; values=list(self.download_tree.item(item,'values')); current=values[2]
        new=filedialog.askdirectory(title='Choose destination folder',initialdir=current if Path(current).exists() else str(Path(self.root_path.get())))
        if new:
            values[2]=new; self.download_tree.item(item,values=values)
            for row in self.download_rows:
                if row['filename']==values[0]: row['destination']=new; break

    def move_downloads(self, all_rows=False):
        items=list(self.download_tree.get_children()) if all_rows else list(self.download_tree.selection())
        if not items: return messagebox.showwarning('Select PDFs','Select one or more PDFs first.')
        moved=0; errors=[]
        for item in items:
            values=self.download_tree.item(item,'values'); source=Path(next((r['source'] for r in self.download_rows if r['filename']==values[0]),'')); dest=Path(values[2])
            try:
                dest.mkdir(parents=True,exist_ok=True); target=dest/source.name
                if target.exists():
                    stem,suffix=source.stem,source.suffix; n=2
                    while (dest/f'{stem} ({n}){suffix}').exists(): n+=1
                    target=dest/f'{stem} ({n}){suffix}'
                shutil.move(str(source),str(target)); moved+=1
                self.download_tree.set(item,'status',f'Moved to {target}')
            except Exception as e: errors.append(f'{source.name}: {e}')
        messagebox.showinfo('Download routing complete',f'Moved {moved} PDF(s).'+(('\n\n'+'\n'.join(errors)) if errors else ''))
        self.scan_downloads()
    def open_new_invoice(self,event):
        # Use the row under the pointer; focus() can still refer to the
        # previously selected row when the user double-clicks a new one.
        item=self.new_invoice_tree.identify_row(event.y)
        if not item: return
        self.new_invoice_tree.selection_set(item)
        self.new_invoice_tree.focus(item)
        values=self.new_invoice_tree.item(item,'values')
        # New Invoices columns are: id, path, vendor, ... .  The old
        # index pointed at vendor, so double-clicking newer invoices tried
        # to open a non-existent path and silently did nothing.
        invoice_id=values[0]; path=values[1]
        if path and Path(path).exists():
            os.startfile(path)
        else:
            messagebox.showwarning('PDF not found',f'The original PDF could not be found:\n{path}')
        stamp=datetime.now().isoformat(timespec='seconds'); self.db.execute('UPDATE invoices SET recorded=1,recorded_at=?,updated=? WHERE id=?',(stamp,stamp,invoice_id)); self.db.commit()
        self.root.after(150,self.page_new_invoices)
    def mark_selected_recorded(self):
        items=self.new_invoice_tree.selection()
        if not items: return messagebox.showwarning('Select invoices','Select one or more invoices first.')
        ids=[self.new_invoice_tree.item(item,'values')[0] for item in items]
        stamp=datetime.now().isoformat(timespec='seconds'); self.db.executemany('UPDATE invoices SET recorded=1,recorded_at=?,updated=? WHERE id=?',[(stamp,stamp,i) for i in ids]); self.db.commit(); self.page('New Invoices')
    def mark_all_recorded(self):
        count=self.db.execute("SELECT COUNT(*) FROM invoices WHERE COALESCE(recorded,0)=0 AND substr(added,1,10)=?",(date.today().isoformat(),)).fetchone()[0]
        if not count: return messagebox.showinfo('New Invoices','There are no unrecorded invoices.')
        stamp=datetime.now().isoformat(timespec='seconds'); self.db.execute("UPDATE invoices SET recorded=1,recorded_at=?,updated=? WHERE COALESCE(recorded,0)=0 AND substr(added,1,10)=?",(stamp,stamp,date.today().isoformat())); self.db.commit(); self.page('New Invoices')
    def company_view(self,company):
        self.last_query=company; self.page('Invoices')
    def chart_monthly(self,parent,rows):
        f=ttk.LabelFrame(parent,text='Monthly spending',padding=8); f.pack(side='left',fill='both',expand=True,padx=(0,8)); c=tk.Canvas(f,height=190,background='white',highlightthickness=0); c.pack(fill='both',expand=True)
        data={}
        for r in rows:
            label=r['invoice_date'][:7] if r['invoice_date'] and re.fullmatch(r'\d{4}-\d{2}-\d{2}',str(r['invoice_date'])) else 'Unknown'
            data[label]=data.get(label,0)+(r['amount'] or 0)
        valid=sorted(((k,v) for k,v in data.items() if k!='Unknown'),key=lambda x:x[0])[-7:]
        vals=valid+([( 'Unknown',data['Unknown'])] if 'Unknown' in data else []); mx=max([v for _,v in vals] or [1]); w=70
        for i,(label,val) in enumerate(vals):
            x=25+i*w; h=145*val/mx; c.create_rectangle(x,165-h,x+42,165,fill='#D97706' if label=='Unknown' else '#2B7A78',outline=''); c.create_text(x+21,178,text=label,font=('Segoe UI',8)); c.create_text(x+21,155-h,text=f'${val/1000:.0f}k',font=('Segoe UI',8),fill='#16324F')
            if label=='Unknown': c.tag_bind(c.create_rectangle(x,0,x+42,190,fill='',outline=''),'<Button-1>',lambda e:self.show_unknown_dates())
    def show_unknown_dates(self):
        self.f_unknown_value='Unknown Date'; self.page('Invoices')
    def pick_invoice_filter_date(self, var):
        pop=tk.Toplevel(self.root); pop.title('Choose date'); pop.transient(self.root); pop.grab_set()
        try: selected=date.fromisoformat(var.get())
        except ValueError: selected=date.today()
        state={'year':selected.year,'month':selected.month}
        ttk.Label(pop,text='Select a date').pack(padx=15,pady=10); cal=ttk.Frame(pop); cal.pack(padx=12,pady=8)
        def draw(year,month):
            state.update(year=year,month=month)
            for w in cal.winfo_children(): w.destroy()
            nav=tk.Frame(cal); nav.grid(row=0,column=0,columnspan=7,pady=5)
            ttk.Button(nav,text='<',width=3,command=lambda: change_month(-1)).pack(side='left')
            ttk.Label(nav,text=f'{calendar.month_name[month]} {year}',font=('Segoe UI',11,'bold'),width=18,anchor='center').pack(side='left')
            ttk.Button(nav,text='>',width=3,command=lambda: change_month(1)).pack(side='left')
            for col,name in enumerate(('Mo','Tu','We','Th','Fr','Sa','Su')): ttk.Label(cal,text=name,width=4).grid(row=1,column=col)
            for row,week in enumerate(calendar.monthcalendar(year,month),2):
                for col,daynum in enumerate(week):
                    if daynum: ttk.Button(cal,text=str(daynum),width=4,command=lambda d=daynum: (var.set(f'{year:04d}-{month:02d}-{d:02d}'),pop.destroy())).grid(row=row,column=col,padx=1,pady=1)
        def change_month(delta):
            month=state['month']+delta; year=state['year']
            if month<1: month,year=12,year-1
            if month>12: month,year=1,year+1
            draw(year,month)
        draw(state['year'],state['month'])
    def chart_companies(self,parent,rows):
        f=ttk.LabelFrame(parent,text='Company spending',padding=8); f.pack(side='left',fill='both',expand=True); c=tk.Canvas(f,height=190,background='white',highlightthickness=0); c.pack(fill='both',expand=True); data=[('NHW',sum((r['amount'] or 0) for r in rows if r['company']=='NHW'), '#D97706'),('SPX',sum((r['amount'] or 0) for r in rows if r['company']=='SPX'), '#2563EB')]; mx=max([v for _,v,_ in data] or [1])
        for i,(label,val,color) in enumerate(data):
            y=35+i*65; c.create_text(12,y+15,text=label,anchor='w',font=('Segoe UI',10,'bold')); c.create_rectangle(65,y,265,y+30,fill='#E5E7EB',outline=''); c.create_rectangle(65,y,65+200*val/mx,y+30,fill=color,outline=''); c.create_text(275,y+15,text=f'${val:,.0f}',anchor='w',font=('Segoe UI',9))
    def page_invoices(self):
        filters=ttk.LabelFrame(self.main,text='Invoice filters',padding=10); filters.pack(fill='x',pady=(0,10)); filters.columnconfigure(1,weight=1); filters.columnconfigure(3,weight=1); self.q=tk.StringVar(value=getattr(self,'last_query','')); self.f_vendor=tk.StringVar(value=getattr(self,'f_vendor_value','All vendors')); self.f_from=tk.StringVar(value=getattr(self,'f_from_value','')); self.f_to=tk.StringVar(value=getattr(self,'f_to_value','')); self.f_due=tk.StringVar(value=getattr(self,'f_due_value','')); self.f_terms=tk.StringVar(value=getattr(self,'f_terms_value','')); self.f_company=tk.StringVar(value=getattr(self,'f_company_value','All companies')); self.f_unknown=tk.StringVar(value=getattr(self,'f_unknown_value','All dates')); self.f_document_type=tk.StringVar(value=getattr(self,'f_document_type_value','All types'))
        vendors=['All vendors']+[r['vendor'] for r in self.db.execute('SELECT DISTINCT vendor FROM invoices WHERE vendor IS NOT NULL ORDER BY vendor').fetchall()]; companies=['All companies','NHW','SPX','Unknown']; document_types=['All types','Bill','Debit Memo','Credit Memo']
        self.f_added=tk.StringVar(value='' if getattr(self,'f_added_value','') == 'All added dates' else getattr(self,'f_added_value',''))
        fields=[('Vendor',self.f_vendor,vendors,20),('Company',self.f_company,companies,16)]
        for i,(label,var,values,width) in enumerate(fields): ttk.Label(filters,text=label).grid(row=0,column=i*2,sticky='w',padx=(0,4),pady=3); ttk.Combobox(filters,textvariable=var,values=values,state='readonly',width=width).grid(row=0,column=i*2+1,sticky='ew',padx=(0,10),pady=3)
        ttk.Label(filters,text='Added At').grid(row=0,column=4,sticky='w',padx=(0,4),pady=3)
        added_row=tk.Frame(filters); added_row.grid(row=0,column=5,sticky='w',padx=(0,10),pady=3); ttk.Entry(added_row,textvariable=self.f_added,width=12).pack(side='left'); ttk.Button(added_row,text='Calendar',command=lambda:self.pick_invoice_filter_date(self.f_added)).pack(side='left',padx=3); ttk.Button(added_row,text='Clear',command=lambda:self.f_added.set('')).pack(side='left')
        ttk.Label(filters,text='Unknown Date').grid(row=0,column=6,sticky='w',padx=(0,4),pady=3); ttk.Combobox(filters,textvariable=self.f_unknown,values=['All dates','Unknown Date','Has Date'],state='readonly',width=14).grid(row=0,column=7,sticky='ew',padx=(0,10),pady=3)
        ttk.Label(filters,text='Document Type').grid(row=1,column=0,sticky='w',padx=(0,4),pady=3); ttk.Combobox(filters,textvariable=self.f_document_type,values=document_types,state='readonly',width=16).grid(row=1,column=1,sticky='ew',padx=(0,10),pady=3)
        for i,(label,var) in enumerate([('Date From',self.f_from),('Date To',self.f_to),('Due Date',self.f_due),('Terms',self.f_terms)]): ttk.Label(filters,text=label).grid(row=2,column=i*2,sticky='w',padx=(0,4),pady=3); ttk.Entry(filters,textvariable=var,width=12).grid(row=2,column=i*2+1,sticky='ew',padx=(0,10),pady=3)
        ttk.Label(filters,text='Search').grid(row=3,column=0,sticky='w',padx=(0,4),pady=3); search_entry=ttk.Entry(filters,textvariable=self.q); search_entry.grid(row=3,column=1,columnspan=3,sticky='ew',padx=(0,10),pady=3); search_entry.bind('<Return>',lambda e:self.apply_filters()); actions=ttk.Frame(filters); actions.grid(row=3,column=4,columnspan=3,sticky='e',pady=3); ttk.Button(actions,text='Apply Filters',command=self.apply_filters).pack(side='left',padx=3); ttk.Button(actions,text='Clear',command=self.clear_filters).pack(side='left',padx=3); ttk.Button(actions,text='Refresh Data',command=self.apply_filters).pack(side='left',padx=3); ttk.Button(actions,text='Edit Selected',command=self.edit_invoice_selected).pack(side='left',padx=3)
        ttk.Label(self.main,text='Dates use YYYY-MM-DD. Left double-click opens the PDF; right double-click opens its folder.',foreground='#667085').pack(anchor='w'); self.invoice_tree=self.table(self.fetch(),('company','site','vendor','reference','document_type','amount','invoice_date','due_date','category','recorded_at','path')); self.invoice_tree._invoice_columns=('company','site','vendor','reference','document_type','amount','invoice_date','due_date','category','recorded_at','path'); self.invoice_tree.bind('<ButtonRelease-1>',self.copy_reference_selected,add='+'); self.invoice_tree.bind('<Double-3>',self.open_invoice_folder,add='+')
    def copy_reference_selected(self,event):
        item=self.invoice_tree.identify_row(event.y)
        if not item: return
        self.invoice_tree.selection_set(item); self.invoice_tree.focus(item); values=self.invoice_tree.item(item,'values'); reference=values[3] if len(values)>3 else ''
        if reference: self.root.clipboard_clear(); self.root.clipboard_append(str(reference)); self.scan_status.set(f'Reference copied: {reference}')
    def open_invoice_folder(self,event):
        item=self.invoice_tree.identify_row(event.y)
        if not item: return
        self.invoice_tree.selection_set(item); self.invoice_tree.focus(item)
        values=self.invoice_tree.item(item,'values'); path=values[self.invoice_tree._invoice_columns.index('path')] if values else ''
        if path and Path(path).exists():
            os.startfile(str(Path(path).parent))
        elif path:
            messagebox.showwarning('Folder not found',f'The invoice folder could not be found:\n{Path(path).parent}')
    def apply_filters(self):
        for key,var in [('f_vendor_value',self.f_vendor),('f_from_value',self.f_from),('f_to_value',self.f_to),('f_due_value',self.f_due),('f_terms_value',self.f_terms),('f_company_value',self.f_company),('f_document_type_value',self.f_document_type),('f_added_value',self.f_added),('f_unknown_value',self.f_unknown)]: setattr(self,key,var.get())
        self.last_query=self.q.get(); self.page('Invoices')
    def clear_filters(self):
        self.last_query=''; self.f_vendor_value='All vendors'; self.f_from_value=''; self.f_to_value=''; self.f_due_value=''; self.f_terms_value=''; self.f_company_value='All companies'; self.f_document_type_value='All types'; self.f_added_value=''; self.f_unknown_value='All dates'; self.page('Invoices')
    def fetch(self):
        q=getattr(self,'last_query','').strip(); sql='SELECT * FROM invoices WHERE 1=1'; args=[]
        filters=[('vendor',getattr(self,'f_vendor_value','All vendors'),'All vendors','='),('company',getattr(self,'f_company_value','All companies'),'All companies','='),('document_type',getattr(self,'f_document_type_value','All types'),'All types','='),('invoice_date',getattr(self,'f_from_value',''),'','>='),('invoice_date',getattr(self,'f_to_value',''),'','<='),('due_date',getattr(self,'f_due_value',''),'','=')]
        for col,val,skip,op in filters:
            if val and val!=skip: sql+=f' AND {col} {op} ?'; args.append(val)
        added=getattr(self,'f_added_value','')
        if added: sql+=' AND substr(added,1,10)=?'; args.append(added)
        unknown=getattr(self,'f_unknown_value','All dates')
        if unknown=='Unknown Date': sql+=' AND (invoice_date IS NULL OR date_source IN (\'file timestamp\',\'\'))'
        elif unknown=='Has Date': sql+=' AND invoice_date IS NOT NULL AND COALESCE(date_source,\'\') NOT IN (\'file timestamp\',\'\')'
        terms=getattr(self,'f_terms_value','').strip()
        if terms: sql+=' AND terms LIKE ?'; args.append(f'%{terms}%')
        if q: sql+=' AND (vendor LIKE ? OR reference LIKE ? OR invoice_number LIKE ? OR filename LIKE ? OR category LIKE ? OR company LIKE ? OR site LIKE ? OR payment_status LIKE ?)'; args.extend([f'%{q}%']*8)
        return self.db.execute(sql+' ORDER BY invoice_date DESC, id DESC',args).fetchall()
    def page_review_queue(self):
        bar=ttk.Frame(self.main); bar.pack(fill='x',pady=(0,8)); ttk.Label(bar,text='Review uncertain extractions before reporting').pack(side='left'); ttk.Button(bar,text='Edit Selected',command=self.edit_review).pack(side='right',padx=4); ttk.Button(bar,text='Set Company',command=self.set_company).pack(side='right',padx=4); ttk.Button(bar,text='Enhanced Image OCR',command=lambda:self.scan_review(True,True)).pack(side='right',padx=4); ttk.Button(bar,text='Rescan with OCR',command=lambda:self.scan_review(True)).pack(side='right',padx=4); ttk.Button(bar,text='Rescan Review Queue',command=self.scan_review).pack(side='right',padx=4); ttk.Button(bar,text='Select All',command=lambda:self.select_all(self.review_tree)).pack(side='right',padx=4); ttk.Button(bar,text='Resolve Selected',command=lambda:self.review_decision('Resolved')).pack(side='right',padx=4); ttk.Button(bar,text='Needs Attention',command=lambda:self.review_decision('Needs Attention')).pack(side='right')
        ttk.Button(bar,text='Resolve & Record',command=self.resolve_and_record).pack(side='right',padx=4)
        self.review_tree=self.table(self.db.execute('SELECT id,* FROM invoices WHERE confidence<.7 OR status IN ("Review","Needs Attention") ORDER BY id DESC').fetchall(),('id','vendor','reference','amount','invoice_date','terms','due_date','status','recorded','confidence','path'))
    def review_decision(self,decision):
        items=self.review_tree.selection()
        if not items: return messagebox.showwarning('Select invoice','Select one or more review items first.')
        stamp=datetime.now().isoformat(timespec='seconds'); ids=[self.review_tree.item(i,'values')[0] for i in items]; self.db.executemany('UPDATE invoices SET status=?,updated=? WHERE id=?',[(decision,stamp,i) for i in ids]); self.db.executemany('INSERT INTO audit(invoice_id,changed,at,details) VALUES(?,?,?,?)',[(i,'Review status',stamp,decision) for i in ids]); self.db.commit(); self.page('Review Queue')
    def resolve_and_record(self):
        items=self.review_tree.selection()
        if not items: return messagebox.showwarning('Select invoice','Select one or more review items first.')
        stamp=datetime.now().isoformat(timespec='seconds'); ids=[self.review_tree.item(i,'values')[0] for i in items]
        self.db.executemany("UPDATE invoices SET status='Resolved',recorded=1,recorded_at=?,updated=? WHERE id=?",[(stamp,stamp,i) for i in ids]); self.db.executemany('INSERT INTO audit(invoice_id,changed,at,details) VALUES(?,?,?,?)',[(i,'Review resolved and recorded',stamp,'Marked Resolved and Recorded') for i in ids]); self.db.commit(); self.page('Review Queue')
    def select_all(self,tree): tree.selection_set(tree.get_children())
    def set_company(self):
        items=self.review_tree.selection()
        if not items: return messagebox.showwarning('Select invoice','Select one or more invoices first.')
        company=simpledialog.askstring('Set company','Enter company (NHW, SPX, PMS, or another name):',parent=self.root)
        if not company or not company.strip(): return
        company=company.strip().upper(); stamp=datetime.now().isoformat(timespec='seconds'); ids=[self.review_tree.item(i,'values')[0] for i in items]
        self.db.executemany('UPDATE invoices SET company=?,updated=? WHERE id=?',[(company,stamp,i) for i in ids]); self.db.executemany('INSERT INTO audit(invoice_id,changed,at,details) VALUES(?,?,?,?)',[(i,'Company',stamp,'Set to '+company) for i in ids]); self.db.commit(); self.page('Review Queue')
    def edit_invoice_selected(self): self.edit_review(getattr(self,'invoice_tree',None),'Invoices')
    def edit_review(self, tree=None, return_page='Review Queue'):
        tree=tree or self.review_tree; items=tree.selection()
        if len(items)!=1: return messagebox.showwarning('Select one invoice','Select exactly one review item to edit.')
        values=tree.item(items[0],'values'); invoice_id=values[0]; r=self.db.execute('SELECT * FROM invoices WHERE id=?',(invoice_id,)).fetchone()
        if r is None and 'path' in getattr(tree,'_invoice_columns',[]): r=self.db.execute('SELECT * FROM invoices WHERE path=?',(values[tree._invoice_columns.index('path')],)).fetchone(); invoice_id=r['id'] if r else None
        if r is None: return messagebox.showerror('Invoice not found','Refresh the invoice list and try again.')
        win=tk.Toplevel(self.root); win.title('Edit invoice'); win.geometry('520x720'); win.transient(self.root); win.grab_set(); fields=[('Vendor','vendor'),('Reference','reference'),('Invoice number','invoice_number'),('Amount','amount'),('Invoice date','invoice_date'),('Terms','terms'),('Due date','due_date'),('Category','category'),('Company','company'),('Site','site'),('Document type','document_type'),('Payment status','payment_status'),('Notes','notes')]; vars={}
        vendor_values=[x['vendor'] for x in self.db.execute("SELECT DISTINCT vendor FROM invoices WHERE vendor IS NOT NULL AND trim(vendor)<>'' ORDER BY vendor COLLATE NOCASE")]
        choices={'vendor':vendor_values,'terms':['Net 0','Net 15','Net 30','Net 45','Net 60','Due on receipt'],'company':['NHW','SPX','PMS','Unknown'],'site':['','LEAGUE CITY','Lotus El Paso','Lotus Allen','Lotus Nacogdoches','Lotus Pearland','Lotus Plano','Lotus ATL','Lotus San Antonio','PFWC','Infusion'],'document_type':['Bill','Credit Memo','Debit Memo'],'payment_status':['Unpaid','Paid']}
        def pick_date(var):
            pop=tk.Toplevel(win); pop.title('Choose date'); pop.transient(win); pop.grab_set(); today=date.today(); state={'year':today.year,'month':today.month}
            ttk.Label(pop,text='Select a date').pack(padx=15,pady=10); cal=ttk.Frame(pop); cal.pack(padx=12,pady=8)
            def draw(year,month):
                state.update(year=year,month=month)
                for w in cal.winfo_children(): w.destroy()
                nav=tk.Frame(cal); nav.grid(row=0,column=0,columnspan=7,pady=5); ttk.Button(nav,text='‹',width=3,command=lambda: change_month(-1)).pack(side='left'); ttk.Label(nav,text=f'{calendar.month_name[month]} {year}',font=('Segoe UI',11,'bold'),width=18,anchor='center').pack(side='left'); ttk.Button(nav,text='›',width=3,command=lambda: change_month(1)).pack(side='left')
                for col,name in enumerate(('Mo','Tu','We','Th','Fr','Sa','Su')): ttk.Label(cal,text=name,width=4).grid(row=1,column=col)
                for row,week in enumerate(calendar.monthcalendar(year,month),2):
                    for col,daynum in enumerate(week):
                        if daynum: ttk.Button(cal,text=str(daynum),width=4,command=lambda d=daynum: (var.set(f'{year:04d}-{month:02d}-{d:02d}'),pop.destroy())).grid(row=row,column=col,padx=1,pady=1)
            def change_month(delta):
                month=state['month']+delta; year=state['year']
                if month<1: month,year=12,year-1
                if month>12: month,year=1,year+1
                draw(year,month)
            draw(state['year'],state['month'])
        def recalc_due(*_):
            try:
                base=date.fromisoformat(vars['invoice_date'].get()); text=vars['terms'].get().lower(); match=re.search(r'\d+',text); days=int(match.group()) if match else 0; vars['due_date'].set((base+timedelta(days=days)).isoformat())
            except Exception: pass
        for label,key in fields:
            ttk.Label(win,text=label).pack(anchor='w',padx=18,pady=(7,1)); v=tk.StringVar(value='' if r[key] is None else str(r[key])); vars[key]=v
            if key in choices:
                widget=ttk.Combobox(win,textvariable=v,values=choices[key],width=39,state='readonly'); widget.pack(padx=18); widget.bind('<<ComboboxSelected>>',recalc_due)
            elif key in ('invoice_date','due_date'):
                row=tk.Frame(win); row.pack(anchor='w',padx=18); ttk.Entry(row,textvariable=v,width=32).pack(side='left'); ttk.Button(row,text='Calendar',command=lambda value=v:pick_date(value)).pack(side='left',padx=5)
                if key=='invoice_date': v.trace_add('write',recalc_due)
            else: ttk.Entry(win,textvariable=v,width=42).pack(padx=18)
        vars['due_date'].trace_add('write',lambda *_: None)
        def save():
            try: amount=float(vars['amount'].get().replace(',','')) if vars['amount'].get().strip() else None
            except ValueError: return messagebox.showerror('Invalid amount','Enter a valid number.')
            stamp=datetime.now().isoformat(timespec='seconds'); keys=[k for _,k in fields]; values=[vars[k].get().strip() or None for k in keys]; values[keys.index('amount')]=amount; assignments=','.join(f'{k}=?' for k in keys); self.db.execute(f'UPDATE invoices SET {assignments},status=\'Review\',confidence=1.0,updated=? WHERE id=?',(*values,stamp,invoice_id)); self.audit_log('invoice edit',str(invoice_id),'','Manual correction','Review fields edited',invoice_id=invoice_id); self.db.commit(); win.destroy(); self.page(return_page)
        ttk.Button(win,text='Save Corrections',command=save).pack(pady=18)
    def scan_review(self, force_ocr=False, enhanced_ocr=False):
        if getattr(self,'_scanning',False): return
        paths=[r['path'] for r in self.db.execute('SELECT path FROM invoices WHERE confidence<.7 OR status IN ("Review","Needs Attention")').fetchall() if r['path']]
        if not paths: return messagebox.showinfo('Review Queue','There are no review items to rescan.')
        self._scanning=True; self._silent_scan=False; self._review_scan=True; self.scan_status.set('Running enhanced image OCR…' if enhanced_ocr else ('Rescanning Review Queue with OCR…' if force_ocr else 'Rescanning Review Queue…')); threading.Thread(target=self._scan_worker,args=(self.root_path.get(),set(),True,paths,force_ocr,enhanced_ocr),daemon=True).start()
    def page_duplicates(self): self.table(self.db.execute('SELECT * FROM invoices WHERE duplicate=1 AND ignored=0 ORDER BY id DESC').fetchall(),('vendor','reference','invoice_number','amount','invoice_date','path','filename'))
    def page_reconciliation(self):
        bar=ttk.Frame(self.main); bar.pack(fill='x',pady=(0,8)); ttk.Label(bar,text='Compare recorded invoices with QuickBooks CSV reports').pack(side='left')
        ttk.Label(bar,text='Month').pack(side='left',padx=(18,4)); self.recon_month=tk.StringVar(value=getattr(self,'recon_month_value',date.today().strftime('%Y-%m')))
        months=[r['month'] for r in self.db.execute("SELECT DISTINCT substr(invoice_date,1,7) month FROM invoices WHERE invoice_date IS NOT NULL ORDER BY month DESC").fetchall()]
        ttk.Combobox(bar,textvariable=self.recon_month,values=months,state='readonly',width=10).pack(side='left')
        ttk.Label(bar,text='Company').pack(side='left',padx=(10,4)); self.recon_company=tk.StringVar(value=getattr(self,'recon_company_value','All'))
        ttk.Combobox(bar,textvariable=self.recon_company,values=['All','NHW','SPX'],state='readonly',width=8).pack(side='left')
        self.recon_file_label=tk.StringVar(value='No QuickBooks CSV selected'); ttk.Button(bar,text='Choose QuickBooks CSV',command=self.choose_reconciliation_file).pack(side='left',padx=6); ttk.Label(bar,textvariable=self.recon_file_label,foreground='#475569').pack(side='left',padx=6); ttk.Button(bar,text='Run Reconciliation',command=self.refresh_reconciliation).pack(side='left',padx=6); ttk.Button(bar,text='Choose Another File',command=self.choose_reconciliation_file).pack(side='left',padx=6)
        ttk.Label(bar,text='Match Status').pack(side='left',padx=(14,3)); self.recon_status=tk.StringVar(value='All Results'); ttk.Combobox(bar,textvariable=self.recon_status,values=['All Results','Matched','Unmatched/Mismatched','Missing in QuickBooks','Missing in App','Amount Mismatch','Date Mismatch','Vendor Mismatch'],state='readonly',width=20).pack(side='left'); self.recon_status.trace_add('write',lambda *_: self.apply_reconciliation_filter()); ttk.Button(bar,text='Export Reconciliation',command=self.export_reconciliation).pack(side='left',padx=6)
        ttk.Button(bar,text='Select All',command=lambda:self.select_all(self.recon_tree)).pack(side='right',padx=4); ttk.Button(bar,text='Mark Selected as Recorded',command=self.mark_recorded).pack(side='right',padx=4); ttk.Button(bar,text='Mark All Matched as Recorded',command=self.mark_all_recorded_reconciliation).pack(side='right')
        self.recon_tree=self.table([],('id','company','App Vendor','App Reference','App Amount','App Invoice Date','Match Status','QB Vendor','QB Number','QB Amount','QB Date','QB Memo','Difference','Review Notes'))
        self.recon_count=self.main.winfo_children()[-2]
        # Comparison is intentionally not run until the user selects a file.
    def quickbooks_csv(self, company):
        desktop=Path.home()/'Desktop'; wanted=['BillsNHW.CSV'] if company=='NHW' else ['BillSPX.CSV'] if company=='SPX' else ['BillsNHW.CSV','BillSPX.CSV']
        files=[]
        for name in wanted:
            exact=desktop/name
            if exact.exists(): files.append(exact); continue
            files += [p for p in desktop.iterdir() if p.is_file() and p.name.lower()==name.lower()]
        return files
    def choose_reconciliation_file(self):
        from tkinter import filedialog
        paths=filedialog.askopenfilenames(parent=self.root,title='Select QuickBooks CSV',initialdir=getattr(self,'recon_last_folder',str(Path.home())),filetypes=[('CSV files','*.csv'),('All files','*.*')])
        if not paths: return
        self.recon_paths=[Path(p) for p in paths]; self.recon_last_folder=str(self.recon_paths[0].parent); self.recon_file_label.set(' + '.join(p.name for p in self.recon_paths))
    def read_quickbooks(self, company, selected_paths=None):
        rows=[]; self.qb_rejected=[]; self.qb_ignored=[]
        aliases={
            'curascript':['Curascript SD Specialty','Curascript Specialty Distribution'],
            'prodigy':['Prodigy Health Supplier Corp'],'biocare':['BioCareSD','BioCareSD, Inc.'],
            'amerisource':['Amerisource Bergen'],'aps':['APS Health and Infusion Nursing Services'],
            'mkare':['MKare Management, Inc.'],'helga':['HELGA Healthcare, PLLC'],
            'autola':['Atulo Health'],'brightstar':['BrightStar Care - San Antonio','Brightstar Care of Friendswood'],
            'dinovo':['Dinovo Pharmacy & Packaging Provisions'],
        }
        def norm(v):
            v=re.sub(r'[^a-z0-9 ]','',str(v or '').casefold()); v=re.sub(r'\b(llc|inc|corp|corporation|pllc)\b','',v); return re.sub(r'\s+',' ',v).strip()
        app_vendors={norm(r['vendor']):r['vendor'] for r in self.db.execute("SELECT DISTINCT vendor FROM invoices WHERE vendor IS NOT NULL AND trim(vendor)<>''").fetchall()}
        lookup={norm(k):k for k in app_vendors}
        for app, names in aliases.items():
            if app in app_vendors: lookup.update({norm(name):app_vendors[app] for name in names})
        paths=selected_paths or []
        for path in paths:
            filename=path.name.casefold()
            file_company='NHW' if 'nhw' in filename else ('SPX' if 'spx' in filename else company)
            if company!='All' and file_company not in (company,''):
                continue
            current_vendor=''
            try:
                with path.open('r',encoding='utf-8-sig',newline='') as f:
                    for values in csv.reader(f):
                        values=(values+['']*10)[:10]
                        if values[0].strip() and not values[1].strip(): current_vendor=values[0].strip(); continue
                        if not values[1].strip(): continue
                        if values[1].strip().lower()!='bill': self.qb_rejected.append((path.name,'unsupported type')); continue
                        try: amount=Decimal((values[9] or values[8]).replace(',','').replace('$',''))
                        except (InvalidOperation,AttributeError): self.qb_rejected.append((path.name,'invalid amount')); continue
                        try: qdate=datetime.strptime(values[2].strip(),'%m/%d/%Y').date().isoformat()
                        except ValueError: qdate=''
                        mapped=lookup.get(norm(current_vendor))
                        if not mapped:
                            self.qb_ignored.append({'vendor':current_vendor,'number':values[3].strip(),'amount':amount.quantize(Decimal('0.01')),'date':qdate,'company':company,'memo':values[4].strip(),'source':path.name}); continue
                        rows.append({'vendor':current_vendor,'app_vendor':mapped,'number':values[3].strip(),'amount':amount.quantize(Decimal('0.01')),'date':qdate,'company':file_company,'memo':values[4].strip(),'source':path.name})
            except OSError: pass
        return rows
    def refresh_reconciliation(self):
        if not getattr(self,'recon_paths',[]): return messagebox.showwarning('Select QuickBooks CSV','Choose a QuickBooks CSV before running reconciliation.')
        self.recon_month_value=self.recon_month.get(); self.recon_company_value=self.recon_company.get()
        args=[self.recon_month_value]; sql="SELECT id,company,vendor,reference,amount,invoice_date FROM invoices WHERE COALESCE(duplicate,0)=0 AND COALESCE(ignored,0)=0 AND substr(invoice_date,1,7)=?"
        if self.recon_company_value!='All': sql+=' AND company=?'; args.append(self.recon_company_value)
        invoices=self.db.execute(sql+' ORDER BY invoice_date,id',args).fetchall(); qb=self.read_quickbooks(self.recon_company_value,self.recon_paths)
        qb=[r for r in qb if not self.recon_month_value or r.get('date','').startswith(self.recon_month_value)]
        display=[]; used=set(); aliases={'curascript sd specialty':'curascript','curascript specialty distribution':'curascript','curascript':'curascript','amerisource bergen':'amerisource','amerisource':'amerisource'}
        norm=lambda s: re.sub(r'[^a-z0-9]','',str(s or '').casefold())
        vnorm=lambda s: aliases.get(str(s or '').casefold().strip(),norm(s))
        for inv in invoices:
            match=None; reference_candidate=None
            for i,row in enumerate(qb):
                if i in used: continue
                same_amount=inv['amount'] is not None and abs(Decimal(str(inv['amount']))-row['amount'])<=Decimal('0.01')
                same_date=not inv['invoice_date'] or not row['date'] or inv['invoice_date']==row['date']; same_number=bool(inv['reference'] and inv['reference']!='Not Detected' and row['number'] and norm(inv['reference'])==norm(row['number']))
                same_vendor=vnorm(inv['vendor'])==vnorm(row.get('app_vendor') or row['vendor'])
                if same_number and same_vendor:
                    reference_candidate=row
                    if same_amount: match=row; used.add(i); break
            if not match and reference_candidate: match=reference_candidate; used.add(qb.index(reference_candidate))
            if match:
                amount_matches=inv['amount'] is not None and abs(Decimal(str(inv['amount']))-match['amount'])<=Decimal('0.01')
                if not amount_matches: status='Amount Mismatch'
                else: status='Matched - Date Difference' if inv['invoice_date'] and match.get('date') and inv['invoice_date']!=match['date'] else 'Exact Match'
                diff=f'{(Decimal(str(inv["amount"] or 0))-match["amount"]):.2f}'
            else: status='Missing in QuickBooks'; diff=''
            display.append({'id':inv['id'],'company':inv['company'],'App Vendor':inv['vendor'],'App Reference':inv['reference'],'App Amount':inv['amount'],'App Invoice Date':inv['invoice_date'],'Match Status':status,'QB Vendor':match['vendor'] if match else '','QB Number':match['number'] if match else '','QB Amount':match['amount'] if match else '','QB Date':match['date'] if match else '','QB Memo':match['memo'] if match else '','Difference':diff,'Review Notes':''})
        for i,row in enumerate(qb):
            if i not in used: display.append({'id':'','company':row['company'],'App Vendor':'','App Reference':'','App Amount':'','App Invoice Date':'','Match Status':'Missing in App','QB Vendor':row['vendor'],'QB Number':row['number'],'QB Amount':row['amount'],'QB Date':row['date'],'QB Memo':row['memo'],'Difference':'','Review Notes':''})
        self.recon_display=display; self.apply_reconciliation_filter()
        self.scan_status.set(f'QuickBooks rows included: {len(qb):,} | ignored vendor rows: {len(getattr(self,"qb_ignored",[])):,}')
    def apply_reconciliation_filter(self):
        if not hasattr(self,'recon_tree'): return
        wanted=getattr(self,'recon_status',tk.StringVar(value='All Results')).get(); rows=getattr(self,'recon_display',[])
        if wanted=='All Results': filtered=rows
        elif wanted=='Matched': filtered=[r for r in rows if str(r.get('Match Status','')).startswith(('Exact Match','Matched'))]
        elif wanted=='Unmatched/Mismatched': filtered=[r for r in rows if not str(r.get('Match Status','')).startswith(('Exact Match','Matched'))]
        else: filtered=[r for r in rows if r.get('Match Status')==wanted or (wanted=='Date Mismatch' and r.get('Match Status')=='Matched - Date Difference')]
        for item in self.recon_tree.get_children(): self.recon_tree.delete(item)
        cols=('id','company','App Vendor','App Reference','App Amount','App Invoice Date','Match Status','QB Vendor','QB Number','QB Amount','QB Date','QB Memo','Difference','Review Notes')
        for i,r in enumerate(filtered):
            values=[r.get(c,'') if r.get(c,'') is not None else '' for c in cols]; self.recon_tree.insert('', 'end',values=values,tags=('odd' if i%2 else 'even',))
        self.recon_count.configure(text=f'{len(filtered):,} item(s) displayed') if hasattr(self,'recon_count') else None
        self.recon_visible=filtered
    def export_reconciliation(self):
        from tkinter import filedialog
        rows=getattr(self,'recon_visible',[])
        if not rows: return messagebox.showinfo('Export Reconciliation','There are no visible reconciliation results to export.')
        from openpyxl import Workbook
        from openpyxl.styles import Font,PatternFill
        from openpyxl.worksheet.table import Table,TableStyleInfo
        from openpyxl.utils import get_column_letter
        company=self.recon_company_value if self.recon_company_value!='All' else 'All'; status=self.recon_status.get().replace('/','_'); out=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile=f'{company}_Reconciliation_{self.recon_month_value}_{status}.xlsx',filetypes=[('Excel workbook','*.xlsx')])
        if not out: return
        headers=['Company','App Vendor','App Reference','App Amount','App Invoice Date','Match Status','QB Vendor','QB Number','QB Amount','QB Date','Difference']; wb=Workbook(); ws=wb.active; ws.title='Reconciliation'; ws.append(headers)
        for r in rows:
            def dt(v): return datetime.fromisoformat(v).date() if v else None
            ws.append([r.get('company',''),r.get('App Vendor',''),str(r.get('App Reference','')),r.get('App Amount'),dt(r.get('App Invoice Date')),r.get('Match Status',''),r.get('QB Vendor',''),str(r.get('QB Number','')),r.get('QB Amount'),dt(r.get('QB Date')),float(r.get('Difference')) if r.get('Difference') not in ('',None) else None])
        for cell in ws[1]: cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='1F4E78')
        for row in ws.iter_rows(min_row=2):
            for i in (3,8,10): row[i].number_format='$#,##0.00'
            for i in (4,9):
                if row[i].value: row[i].number_format='dddd, mmmm d, yyyy'
        ws.freeze_panes='A2'; tab=Table(displayName='ReconciliationResults',ref=ws.dimensions); tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True,showColumnStripes=False); ws.add_table(tab)
        for col in ws.columns: ws.column_dimensions[get_column_letter(col[0].column)].width=min(max(max(len(str(x.value or '')) for x in col)+2,12),40)
        wb.save(out); messagebox.showinfo('Exported',out)
    def mark_recorded(self):
        items=self.recon_tree.selection()
        if not items: return messagebox.showwarning('Select invoice','Select one or more invoices first.')
        now=date.today().isoformat(); stamp=datetime.now().isoformat(timespec='seconds'); ids=[self.recon_tree.item(i,'values')[0] for i in items]
        ids=[i for i in ids if str(i).isdigit()]
        self.db.executemany("UPDATE invoices SET recorded=1,recorded_at=?,updated=? WHERE id=?",[(stamp,stamp,i) for i in ids]); self.db.commit(); self.page('Reconciliation')
    def mark_all_recorded_reconciliation(self):
        if not messagebox.askyesno('Mark matched as recorded','Mark all currently matched app invoices as recorded?'): return
        ids=[]
        for item in self.recon_tree.get_children():
            values=self.recon_tree.item(item,'values')
            if values and str(values[6])=='Exact Match' and str(values[0]).isdigit(): ids.append(values[0])
        stamp=datetime.now().isoformat(timespec='seconds'); self.db.executemany('UPDATE invoices SET recorded=1,recorded_at=?,updated=? WHERE id=?',[(stamp,stamp,i) for i in ids]); self.db.commit(); self.page('Reconciliation')
    def page_vendors(self):
        bar=ttk.Frame(self.main); bar.pack(fill='x',pady=(0,10)); vendors=['All vendors']+[r['vendor'] for r in self.db.execute('SELECT DISTINCT vendor FROM invoices WHERE vendor IS NOT NULL ORDER BY vendor').fetchall()]; months=['All months']+[r['month'] for r in self.db.execute("SELECT DISTINCT substr(invoice_date,1,7) month FROM invoices WHERE invoice_date IS NOT NULL ORDER BY month DESC").fetchall()]
        self.vendor_filter=tk.StringVar(value=getattr(self,'vendor_filter_value','All vendors')); self.vendor_month=tk.StringVar(value=getattr(self,'vendor_month_value','All months'))
        ttk.Label(bar,text='Vendor').pack(side='left'); ttk.Combobox(bar,textvariable=self.vendor_filter,values=vendors,state='readonly',width=20).pack(side='left',padx=5); ttk.Label(bar,text='Month').pack(side='left'); ttk.Combobox(bar,textvariable=self.vendor_month,values=months,state='readonly',width=14).pack(side='left',padx=5); ttk.Button(bar,text='Apply',command=self.apply_vendor_filters).pack(side='left'); ttk.Button(bar,text='Clear',command=self.clear_vendor_filters).pack(side='left',padx=5)
        sql="SELECT vendor, COUNT(*) count, SUM(amount) spending, AVG(amount) average FROM invoices WHERE COALESCE(duplicate,0)=0 AND COALESCE(ignored,0)=0"; args=[]
        if self.vendor_filter.get()!='All vendors': sql+=' AND vendor=?'; args.append(self.vendor_filter.get())
        if self.vendor_month.get()!='All months': sql+=' AND substr(invoice_date,1,7)=?'; args.append(self.vendor_month.get())
        sql+=' AND lower(trim(vendor)) NOT IN (SELECT lower(trim(vendor)) FROM excluded_vendors)'; rows=self.db.execute(sql+' GROUP BY vendor ORDER BY spending DESC',args).fetchall(); self.vendor_tree=self.table(rows,('vendor','count','spending','average')); ttk.Button(self.main,text='Remove Selected Vendor',command=self.remove_selected_vendor).pack(anchor='e',pady=6); ttk.Label(self.main,text='Removed vendors are hidden from dashboards but invoice records are preserved.',foreground='#667085').pack(anchor='e')
    def remove_selected_vendor(self):
        items=self.vendor_tree.selection()
        if len(items)!=1: return messagebox.showwarning('Select vendor','Select exactly one vendor to remove.')
        vendor=self.vendor_tree.item(items[0],'values')[0]
        if not messagebox.askyesno('Remove vendor',f'Hide {vendor} from Vendors and Dashboard? Invoice records will be preserved.'): return
        self.db.execute('INSERT OR REPLACE INTO excluded_vendors(vendor,removed_at) VALUES(?,?)',(vendor.strip(),datetime.now().isoformat(timespec='seconds'))); self.db.commit(); self.refresh_dashboard(); self.page('Vendors')
    def apply_vendor_filters(self): self.vendor_filter_value=self.vendor_filter.get(); self.vendor_month_value=self.vendor_month.get(); self.page('Vendors')
    def clear_vendor_filters(self): self.vendor_filter_value='All vendors'; self.vendor_month_value='All months'; self.page('Vendors')
    def page_settings(self):
        ttk.Label(self.main,text='Invoice scan folder').pack(anchor='w'); e=ttk.Entry(self.main,textvariable=self.root_path,width=80); e.pack(anchor='w',pady=6); ttk.Button(self.main,text='Browse',command=lambda:self.root_path.set(filedialog.askdirectory() or self.root_path.get())).pack(anchor='w'); ttk.Button(self.main,text='Save Settings',command=lambda:(self.save_setting('root',self.root_path.get()),self.audit_log('settings changed','Scan folder','',self.root_path.get()),messagebox.showinfo('Saved','Settings saved.'))).pack(anchor='w',pady=10)
        box=ttk.LabelFrame(self.main,text='Local Network Access',padding=12); box.pack(fill='x',pady=18)
        enabled=tk.BooleanVar(value=self.setting('network_access','0')=='1'); ttk.Checkbutton(box,text='Local Network Access: ON/OFF',variable=enabled,command=lambda:(self.save_setting('network_access','1' if enabled.get() else '0'),self.toggle_server(),self.audit_log('settings changed','Local Network Access','','ON' if enabled.get() else 'OFF'),self.page('Settings'))).pack(anchor='w')
        host=self.local_ip(); port=self.setting('network_port','8765'); running='Running' if self.http_server else 'Stopped'; url=f'http://{host}:{port}' if enabled.get() else f'http://127.0.0.1:{port}'
        for label,value in [('Current server status',running),('Local IP address',host),('Port',port),('Network URL',url)]: ttk.Label(box,text=f'{label}: {value}').pack(anchor='w',pady=2)
        ttk.Button(box,text='Start server',command=lambda:(self.save_setting('network_access','1'),self.toggle_server(),self.page('Settings'))).pack(side='left',pady=8); ttk.Button(box,text='Stop server',command=lambda:(self.save_setting('network_access','0'),self.toggle_server(),self.page('Settings'))).pack(side='left',padx=8,pady=8)

    def page_user_management(self):
        if not self.admin_only(): return
        ttk.Button(self.main,text='Create user',command=self.create_user).pack(anchor='e',pady=5)
        self.user_tree=self.table(self.db.execute('SELECT username,email,active,role,created_at FROM users ORDER BY username').fetchall(),('username','email','active','role','created_at'))
        ttk.Button(self.main,text='Set/reset password for selected user',command=self.reset_user_password).pack(anchor='e',pady=6)
        ttk.Button(self.main,text='Activate/deactivate selected user',command=self.toggle_user).pack(anchor='e')
    def create_user(self):
        u=simpledialog.askstring('Create user','Username:',parent=self.root); email=simpledialog.askstring('Create user','Email address:',parent=self.root); p=simpledialog.askstring('Create user','Password:',show='*',parent=self.root)
        if not u or not p: return
        salt=secrets.token_hex(16); now=datetime.now().isoformat(timespec='seconds'); digest=hashlib.pbkdf2_hmac('sha256',p.encode(),salt.encode(),240000).hex()
        try: self.db.execute('INSERT INTO users(username,email,password_hash,salt,created_at,updated_at) VALUES(?,?,?,?,?,?)',(u,email or '',digest,salt,now,now)); self.audit_log('user creation',u,'','created'); self.db.commit(); self.page('User Management')
        except sqlite3.IntegrityError: messagebox.showerror('Create user','That username already exists.')
    def reset_user_password(self):
        sel=self.user_tree.selection();
        if not sel: return
        u=self.user_tree.item(sel[0],'values')[0]; p=simpledialog.askstring('Password',f'New password for {u}:',show='*',parent=self.root)
        if p:
            salt=secrets.token_hex(16); digest=hashlib.pbkdf2_hmac('sha256',p.encode(),salt.encode(),240000).hex(); self.db.execute('UPDATE users SET password_hash=?,salt=?,updated_at=? WHERE username=?',(digest,salt,datetime.now().isoformat(timespec='seconds'),u)); self.audit_log('password change',u,'','changed'); self.page('User Management')
    def toggle_user(self):
        sel=self.user_tree.selection();
        if not sel: return
        u=self.user_tree.item(sel[0],'values')[0]
        if u=='admin': return messagebox.showwarning('Protected account','The administrator cannot be deactivated.')
        old=self.db.execute('SELECT active FROM users WHERE username=?',(u,)).fetchone()['active']; self.db.execute('UPDATE users SET active=? WHERE username=?',(0 if old else 1,u)); self.audit_log('user deactivation' if old else 'user activation',u,str(old),str(0 if old else 1)); self.page('User Management')
    def page_audit_trail(self):
        if not self.admin_only(): return
        rows=self.db.execute('SELECT at,username,action_type,item,previous_value,new_value,details FROM audit ORDER BY id DESC').fetchall(); self.table(rows,('at','username','action_type','item','previous_value','new_value','details'))
    def refresh(self):
        if hasattr(self,'main'):
            self.normalize_existing(); self.db.commit(); self.page('Invoices')
    def scan(self, rescan_all=False):
        if getattr(self, '_scanning', False): return
        self._scanning=True; self._silent_scan=False; self.scan_status.set('Preparing scan…'); self.scan_progress['value']=0
        root=self.root_path.get(); self._last_folder_signature=self.folder_signature(root); skip=set() if rescan_all else {r['hash'] for r in self.db.execute('SELECT hash FROM invoices')}
        self._review_scan=False; threading.Thread(target=self._scan_worker,args=(root,skip,rescan_all),daemon=True).start()

    def rescan_existing_invoices(self):
        if getattr(self, '_scanning', False): return
        # Rescan only the invoices currently shown in New Invoices.
        paths=[r['path'] for r in self.db.execute("SELECT path FROM invoices WHERE COALESCE(recorded,0)=0 AND substr(added,1,10)=? AND path IS NOT NULL",(date.today().isoformat(),)).fetchall() if r['path'] and Path(r['path']).exists()]
        if not paths: return messagebox.showinfo('Rescan Existing','There are no existing PDF invoices to rescan.')
        self._scanning=True; self._silent_scan=False; self._review_scan=False
        self.scan_status.set('Rescanning existing invoices…'); self.scan_progress['value']=0
        threading.Thread(target=self._scan_worker,args=(self.root_path.get(),set(),True,paths),daemon=True).start()
    def _scan_worker(self, root, skip=set(), rescan_all=False, selected_paths=None, force_ocr=False, enhanced_ocr=False):
        files=[Path(x) for x in selected_paths] if selected_paths is not None else (list(Path(root).rglob('*.pdf')) if Path(root).exists() else []); results=[]
        self.root.after(0,lambda:self._scan_started(len(files)))
        for index,p in enumerate(files,1):
            try:
                self.root.after(0,lambda i=index,n=len(files),name=p.name:self._scan_progress(i-1,n,name+' (opening)'))
                h=hashlib.sha256(p.read_bytes()).hexdigest()
                if h in skip: continue
                self.root.after(0,lambda i=index,n=len(files),name=p.name:self._scan_progress(i-1,n,name+' (extracting)'))
                # SQLite stays on the UI thread; only do PDF/file work here.
                results.append((p,h,extract(str(p),root,force_ocr,enhanced_ocr)))
            except Exception as e:
                # Keep the scan moving and make the problematic file visible
                # instead of leaving the progress bar apparently frozen.
                self.root.after(0,lambda i=index,n=len(files),name=p.name,err=str(e):self._scan_progress(i,n,f'{name} (skipped: {err[:80]})'))
            else:
                self.root.after(0,lambda i=index,n=len(files),name=p.name:self._scan_progress(i,n,name))
        self.root.after(0,lambda:self._finish_scan(files,results,rescan_all))
    def _scan_started(self,total):
        self.scan_progress['maximum']=max(total,1); self.scan_progress['value']=0
        self.scan_status.set(f'Scanning 0 of {total} PDFs…')
    def _scan_progress(self,done,total,name):
        self.scan_progress['maximum']=max(total,1); self.scan_progress['value']=done
        self.scan_status.set(f'Scanning {done} of {total}: {name}')
    def _finish_scan(self, files, results, rescan_all=False):
        added=0; improved=0
        for p,h,x in results:
            try:
                existing=self.db.execute('SELECT id FROM invoices WHERE path=?',(str(p),)).fetchone()
                if existing and not rescan_all: continue
                if existing and x.get('date_source') == 'file timestamp':
                    saved=self.db.execute('SELECT invoice_date,date_source FROM invoices WHERE id=?',(existing['id'],)).fetchone()
                    if saved and saved['invoice_date']:
                        x['invoice_date']=saved['invoice_date']; x['date_source']=saved['date_source'] or 'existing'
                now=datetime.now().isoformat(timespec='seconds'); status='Review' if x['confidence']<.7 else 'Imported'
                if x.get('date_source') == 'file timestamp' or not x.get('invoice_date'): status='Needs Review'
                dup=self.db.execute('SELECT id FROM invoices WHERE vendor=? AND (reference=? OR invoice_number=?) AND amount IS NOT NULL AND amount=? AND (invoice_date=? OR invoice_date IS NULL OR ? IS NULL) AND path<>? AND COALESCE(duplicate,0)=0',(x['vendor'],x['reference'],x['invoice_number'],x['amount'],x['invoice_date'],x['invoice_date'],str(p))).fetchone()
                if existing:
                    self.db.execute('''UPDATE invoices SET vendor=?,reference=?,invoice_number=?,document_type=?,amount=?,invoice_date=?,date_source=?,terms=?,due_date=?,category=?,filename=?,updated=?,status=?,confidence=?,hash=?,duplicate=?,company=?,site=? WHERE id=?''',(x['vendor'],x['reference'],x['invoice_number'],x.get('document_type','Bill'),x['amount'],x['invoice_date'],x.get('date_source',''),x['terms'],x['due_date'],x['category'],p.name,now,status,x['confidence'],h,1 if dup else 0,x.get('company','Unknown'),x.get('site',''),existing['id'])); invoice_id=existing['id']; action='Rescanned'
                else:
                    cur=self.db.execute('''INSERT INTO invoices(vendor,reference,invoice_number,document_type,amount,invoice_date,date_source,terms,due_date,category,path,filename,added,updated,status,confidence,hash,duplicate,company,site,recorded) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(x['vendor'],x['reference'],x['invoice_number'],x.get('document_type','Bill'),x['amount'],x['invoice_date'],x.get('date_source',''),x['terms'],x['due_date'],x['category'],str(p),p.name,now,now,status,x['confidence'],h,1 if dup else 0,x.get('company','Unknown'),x.get('site',''),0)); invoice_id=cur.lastrowid; action='Imported'
                self.db.execute('INSERT INTO audit(invoice_id,changed,at,details) VALUES(?,?,?,?)',(invoice_id,action,now,'Automatic extraction')); added+=1
                if x['confidence'] >= .7 and x.get('amount') is not None and x.get('reference'): improved+=1
            except Exception: pass
        self.db.commit(); self._scanning=False; self.scan_progress['value']=self.scan_progress['maximum']; self.scan_status.set(f'Scan complete — {added} new invoice(s) imported.'); self.page('Dashboard')
        if not getattr(self, '_silent_scan', False): messagebox.showinfo('Scan complete',f'Rescanned: {len(files)}\nImproved/resolved: {improved}\nStill needing attention: {len(files)-improved}')
    def watch(self): self.root.after(30000,lambda:(self.scan_silent(),self.watch()))
    def folder_signature(self,root):
        try:
            fs=[p for p in Path(root).rglob('*.pdf')]; return (len(fs),max((p.stat().st_mtime_ns for p in fs),default=0))
        except Exception: return (0,0)
    def scan_silent(self):
        root=self.root_path.get()
        if not getattr(self, '_scanning', False) and self.folder_signature(root)!=getattr(self,'_last_folder_signature',None):
            self._scanning=True; self._silent_scan=True; self.scan_status.set('Background scan running…')
            self._last_folder_signature=self.folder_signature(root)
            skip={r['hash'] for r in self.db.execute('SELECT hash FROM invoices')}
            threading.Thread(target=self._scan_worker,args=(root,skip),daemon=True).start()
    def export(self, rows=None):
        global pd
        if pd is None:
            try:
                import pandas as pd
            except ImportError:
                return messagebox.showerror('Missing dependency','Install pandas and openpyxl first.')
        rows = rows if rows is not None else self.db.execute('SELECT * FROM invoices WHERE COALESCE(duplicate,0)=0 AND COALESCE(ignored,0)=0').fetchall()
        month=getattr(self,'dashboard_month','All months'); suffix=datetime.strptime(month,'%Y-%m').strftime('%B_%Y') if month!='All months' else datetime.now().strftime('%B_%Y')
        out=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile=f'Invoice_Report_{suffix}.xlsx',filetypes=[('Excel workbook','*.xlsx')]);
        if not out: return
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.worksheet.table import Table, TableStyleInfo
        from openpyxl.utils import get_column_letter
        wb=Workbook(); ws=wb.active; ws.title='Invoices'; headers=['Reference No.','Vendor','Amount','Company','Date','Terms','Due Date']; ws.append(headers)
        for r in rows:
            inv_date=datetime.fromisoformat(r['invoice_date']).date() if r['invoice_date'] else None
            due_date=datetime.fromisoformat(r['due_date']).date() if r['due_date'] else None
            ws.append([str(r['reference'] or 'Not Detected'),r['vendor'] or '',r['amount'],r['company'] or '',inv_date,r['terms'] or '',due_date])
        for cell in ws[1]: cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='1F4E78')
        for row in ws.iter_rows(min_row=2):
            row[2].number_format='$#,##0.00';
            for idx in (4,6):
                if row[idx].value: row[idx].number_format='dddd, mmmm d, yyyy'
        ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
        if ws.max_row>=2:
            tab=Table(displayName='InvoiceReport',ref=ws.dimensions); tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showFirstColumn=False,showLastColumn=False,showRowStripes=True,showColumnStripes=False); ws.add_table(tab)
        for col in ws.columns: ws.column_dimensions[get_column_letter(col[0].column)].width=min(max(max(len(str(c.value or '')) for c in col)+2,12),40)
        wb.save(out); messagebox.showinfo('Exported',out)

if __name__=='__main__': App(tk.Tk()).root.mainloop()
