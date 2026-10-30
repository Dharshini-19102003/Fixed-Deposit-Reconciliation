from flask import Flask, request, jsonify, render_template_string, send_file
import sqlite3, os, csv, io, json, re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

APP_DIR=os.path.dirname(os.path.abspath(__file__))
DB=os.environ.get('FD_DATABASE',os.path.join(APP_DIR,'fd_reconciliation.sqlite3'))
app=Flask(__name__)

def db():
 c=sqlite3.connect(DB); c.row_factory=sqlite3.Row
 c.execute('''CREATE TABLE IF NOT EXISTS deposits(id INTEGER PRIMARY KEY, source TEXT, bank TEXT, deposit_no TEXT, principal REAL, currency TEXT, rate REAL, start_date TEXT, maturity TEXT, bank_balance REAL, book_balance REAL, interest_book REAL, interest_calc REAL, restricted TEXT, notes TEXT, status TEXT, reason TEXT, created TEXT)''')
 cols={r['name'] for r in c.execute('PRAGMA table_info(deposits)')}
 for name,sql in [('bank_imported','INTEGER NOT NULL DEFAULT 0'),('book_imported','INTEGER NOT NULL DEFAULT 0'),('bank_file','TEXT'),('book_file','TEXT')]:
  if name not in cols:c.execute(f'ALTER TABLE deposits ADD COLUMN {name} {sql}')
 c.execute('''CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY, at TEXT, actor TEXT, action TEXT, detail TEXT)''')
 c.execute('''CREATE TABLE IF NOT EXISTS review(id INTEGER PRIMARY KEY, preparer TEXT, prepared TEXT, reviewer TEXT, reviewed TEXT, note TEXT)'''); c.commit(); return c
def audit(c,actor,action,detail): c.execute('INSERT INTO audit(at,actor,action,detail) VALUES(?,?,?,?)',(datetime.now().isoformat(timespec='seconds'),actor or 'User',action,detail))
def parse_date(v):
 if not v:return ''
 s=str(v).strip()
 for f in ('%Y-%m-%d','%d/%m/%Y','%d-%m-%Y','%m/%d/%Y','%d-%b-%Y'):
  try:return datetime.strptime(s[:10],f).date().isoformat()
  except: pass
 return s[:10]
def parse_num(v):
 try:return float(str(v).replace(',','').replace('₹','').strip())
 except:return 0.0
def read_file(f):
 name=f.filename or 'upload'; raw=f.read(); ext=os.path.splitext(name)[1].lower(); rows=[]
 if ext in ('.csv','.tsv'):
  try:
   text=raw.decode('utf-8-sig',errors='replace'); rows=list(csv.DictReader(io.StringIO(text),delimiter='\t' if ext=='.tsv' else ','))
  except: pass
 elif ext in ('.xlsx','.xlsm'):
  try:
   import openpyxl
   wb=openpyxl.load_workbook(io.BytesIO(raw),data_only=True,read_only=True)
   for ws in wb.worksheets:
    it=ws.iter_rows(values_only=True); heads=[str(x or '').strip() for x in next(it,[])];
    for vals in it:
     d={heads[i]:vals[i] for i in range(min(len(heads),len(vals))) if heads[i]}
     if any(v is not None for v in vals):rows.append(d)
  except Exception as e:return [],f'Could not read workbook: {e}. Install openpyxl.'
 elif ext=='.xml':
  try:
   import xml.etree.ElementTree as ET
   root=ET.fromstring(raw); rows=[{re.sub(r'^.*}','',ch.tag):ch.text for ch in el} for el in root.iter() if list(el) and all(not list(ch) for ch in el)]
  except Exception as e:return [],f'Could not read XML: {e}'
 elif ext=='.docx':
  try:
   from docx import Document
   d=Document(io.BytesIO(raw));
   for t in d.tables:
    if t.rows:
     h=[c.text.strip() for c in t.rows[0].cells]
     rows += [dict(zip(h,[c.text.strip() for c in r.cells])) for r in t.rows[1:]]
   if not rows:return [],'No tabular data found. Enter extracted values in the review form.'
  except Exception as e:return [],f'Could not read Word file: {e}. Install python-docx.'
 elif ext=='.pdf':
  try:
   import fitz
   text='\n'.join(p.get_text() for p in fitz.open(stream=raw,filetype='pdf'))
   if not text.strip():return [],'This PDF appears scanned. For local OCR, install Tesseract and pytesseract; inspect extracted fields before matching.'
   return [],'PDF text was read locally. This layout needs field mapping; use the entry form to confirm values.'
  except:return [],'Install PyMuPDF for local PDF text extraction. Scanned PDFs need local Tesseract OCR.'
 elif ext in ('.png','.jpg','.jpeg','.tif','.tiff','.bmp'):
  return [],'Image received. Local OCR requires Tesseract and pytesseract; please manually confirm fields before reconciliation.'
 else:return [],'Unsupported file type. Supported: XLSX, CSV, TSV, XML, DOCX, PDF, and common image formats.'
 return rows,''
def pick(r,*terms):
 norm={re.sub(r'[^a-z0-9]','',str(k).lower()):v for k,v in r.items()}
 for t in terms:
  for k,v in norm.items():
   if t in k and v is not None:return v
 return ''
def norm(v): return re.sub(r'[^a-z0-9]','',str(v or '').lower())
def date_compatible(a,b): return not (a and b and a!=b)
def import_rows(c,rows,kind,filename,default_bank,tolerance):
 imported=matched=mismatched=unpaired=0; ambiguous=0
 for r in rows:
  bank=str(pick(r,'bankname','bank') or default_bank or '').strip()
  no=str(pick(r,'depositnumber','depositno','fdnumber','fdno','receiptno','instrumentnumber') or '').strip()
  principal=parse_num(pick(r,'principal','depositamount','originalamount'))
  bal=parse_num(pick(r,'closingbalance','bankbalance','balance','amount','depositamount','principal'))
  ledger=parse_num(pick(r,'ledgerbalance','bookbalance','closingbalance','balance','amount','debit'))
  if not (bank or no or principal or bal or ledger): continue
  sd=parse_date(pick(r,'startdate','opendate','acquisitiondate','depositdate'))
  md=parse_date(pick(r,'maturitydate','maturity'))
  opposite='book_imported' if kind=='bank' else 'bank_imported'; same='bank_imported' if kind=='bank' else 'book_imported'
  candidates=[]
  existing=c.execute('SELECT * FROM deposits WHERE COALESCE('+opposite+',0)=0 AND COALESCE('+same+',0)=1').fetchall()
  if no:
   candidates=[x for x in existing if norm(x['deposit_no'])==norm(no)]
  if not candidates and bank:
   amount=principal or (bal if kind=='bank' else ledger)
   for x in existing:
    if norm(x['bank'])!=norm(bank):continue
    other_amount=parse_num(x['principal']) or (parse_num(x['bank_balance']) if kind=='books' else parse_num(x['book_balance']))
    if not amount or abs(amount-other_amount)>tolerance:continue
    if not (sd or md or x['start_date'] or x['maturity']):continue
    if not date_compatible(sd,x['start_date']) or not date_compatible(md,x['maturity']):continue
    candidates.append(x)
  target=candidates[0] if len(candidates)==1 else None
  if len(candidates)>1:ambiguous+=1
  if target:
   if kind=='bank':
    newbank=bank or target['bank']; newprincipal=principal or target['principal']; newbal=bal or principal; newrate=parse_num(pick(r,'interestrate','rate')) or target['rate']; newstart=sd or target['start_date']; newmat=md or target['maturity']; newcurrency=str(pick(r,'currency') or target['currency'] or 'INR'); newrestrict=str(pick(r,'restricted','lien','collateral') or target['restricted'] or 'No')
    c.execute('UPDATE deposits SET bank=?,principal=?,bank_balance=?,rate=?,start_date=?,maturity=?,currency=?,restricted=?,bank_imported=1,bank_file=?,source=?,notes=COALESCE(notes,?) WHERE id=?',(newbank,newprincipal,newbal,newrate,newstart,newmat,newcurrency,newrestrict,filename,filename,filename,target['id']))
   else:
    newbank=bank or target['bank']; newprincipal=principal or target['principal']; newbook=ledger or principal or bal
    c.execute('UPDATE deposits SET bank=?,principal=?,book_balance=?,interest_book=?,book_imported=1,book_file=?,source=?,notes=COALESCE(notes,?) WHERE id=?',(newbank,newprincipal,newbook,parse_num(pick(r,'interestbook','accruedinterest')),filename,filename,filename,target['id']))
   rec=c.execute('SELECT * FROM deposits WHERE id=?',(target['id'],)).fetchone()
  else:
   if kind=='bank':
    c.execute('INSERT INTO deposits(source,bank,deposit_no,principal,currency,rate,start_date,maturity,bank_balance,book_balance,interest_book,interest_calc,restricted,notes,status,reason,created,bank_imported,book_imported,bank_file) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(filename,bank,no,principal,str(pick(r,'currency') or 'INR'),parse_num(pick(r,'interestrate','rate')),sd,md,bal or principal,0,0,0,str(pick(r,'restricted','lien','collateral') or 'No'),filename,'Awaiting book','Waiting for matching Tally/book evidence',datetime.now().isoformat(timespec='seconds'),1,0,filename))
   else:
    c.execute('INSERT INTO deposits(source,bank,deposit_no,principal,currency,rate,start_date,maturity,bank_balance,book_balance,interest_book,interest_calc,restricted,notes,status,reason,created,bank_imported,book_imported,book_file) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(filename,bank,no,principal,str(pick(r,'currency') or 'INR'),parse_num(pick(r,'interestrate','rate')),sd,md,0,ledger or principal or bal,parse_num(pick(r,'interestbook','accruedinterest')),0,'No',filename,'Awaiting bank','Waiting for matching bank evidence',datetime.now().isoformat(timespec='seconds'),0,1,filename))
   rec=c.execute('SELECT * FROM deposits ORDER BY id DESC LIMIT 1').fetchone()
  imported+=1
  if rec['bank_imported'] and rec['book_imported']:
   diff=parse_num(rec['bank_balance'])-parse_num(rec['book_balance']); status='Matched' if abs(diff)<=tolerance else 'Mismatch'; reason='' if status=='Matched' else f'Bank and book balances differ by {diff:.2f}; investigate interest, timing, TDS, renewal, or ledger classification.'
   c.execute('UPDATE deposits SET status=?,reason=? WHERE id=?',(status,reason,rec['id']))
   if status=='Matched':matched+=1
   else:mismatched+=1
  else:unpaired+=1
 return imported,matched,mismatched,unpaired,ambiguous
@app.get('/')
def home():return render_template_string(HTML)
@app.get('/api/data')
def data():
 c=db(); rows=[dict(x) for x in c.execute('SELECT * FROM deposits ORDER BY id DESC')]; auditrows=[dict(x) for x in c.execute('SELECT * FROM audit ORDER BY id DESC LIMIT 100')]; reviews=[dict(x) for x in c.execute('SELECT * FROM review ORDER BY id DESC LIMIT 1')]; c.close(); return jsonify(rows=rows,audit=auditrows,review=reviews[0] if reviews else {})
@app.post('/api/upload')
def upload():
 kind=request.form.get('kind','bank'); kind='books' if kind=='books' else 'bank'; default_bank=request.form.get('default_bank','').strip(); tolerance=max(0,parse_num(request.form.get('tolerance') or 1)); files=request.files.getlist('files'); c=db(); imported=matched=mismatched=unpaired=ambiguous=0; msgs=[]
 for f in files:
  rows,msg=read_file(f)
  if msg:msgs.append(f.filename+': '+msg); continue
  a,b,d,e,g=import_rows(c,rows,kind,f.filename,default_bank,tolerance); imported+=a; matched+=b; mismatched+=d; unpaired+=e; ambiguous+=g
  audit(c,'User','Upload '+('bank statement' if kind=='bank' else 'Tally/book data'),f'{f.filename}: {a} row(s), {b} matched, {d} mismatch, {e} awaiting counterpart')
 c.commit(); c.close(); return jsonify(imported=imported,matched=matched,mismatched=mismatched,unpaired=unpaired,ambiguous=ambiguous,messages=msgs)
@app.post('/api/deposit')
def deposit():
 d=request.json; c=db(); tol=float(d.get('tolerance') or 1); principal=parse_num(d.get('principal')); bankbal=parse_num(d.get('bank_balance')); book=parse_num(d.get('book_balance')); diff=bankbal-book
 status='Matched' if abs(diff)<=tol else 'Mismatch'; reason=d.get('reason','') or ('Unexplained difference — investigate supporting bank advice and ledger' if status=='Mismatch' else '')
 rate=parse_num(d.get('rate')); calc=principal*rate/100*max(0,(date.fromisoformat(d['audit_date'])-date.fromisoformat(d['start_date'])).days)/365 if d.get('audit_date') and d.get('start_date') else 0
 fields=('source','bank','deposit_no','principal','currency','rate','start_date','maturity','bank_balance','book_balance','interest_book','interest_calc','restricted','notes','status','reason','created')
 vals=(d.get('source','Manual entry'),d.get('bank',''),d.get('deposit_no',''),principal,d.get('currency','INR'),rate,parse_date(d.get('start_date')),parse_date(d.get('maturity')),bankbal,book,parse_num(d.get('interest_book')),calc,d.get('restricted','No'),d.get('notes',''),status,reason,datetime.now().isoformat(timespec='seconds'))
 c.execute('INSERT INTO deposits('+','.join(fields)+',bank_imported,book_imported) VALUES('+','.join('?' for _ in fields)+',1,1)',vals); audit(c,d.get('actor','User'),'Add FD',f"{d.get('bank','')} / {d.get('deposit_no','')} — {status}; variance {diff:.2f}"); c.commit(); c.close(); return jsonify(ok=True)
@app.post('/api/review')
def review():
 d=request.json
 if d.get('reviewer','').strip().lower()==d.get('preparer','').strip().lower():return jsonify(error='Preparer and reviewer must be different people.'),400
 c=db(); c.execute('INSERT INTO review(preparer,prepared,reviewer,reviewed,note) VALUES(?,?,?,?,?)',(d.get('preparer'),d.get('prepared'),d.get('reviewer'),d.get('reviewed'),d.get('note'))); audit(c,d.get('reviewer'),'Review sign-off','Review sign-off saved'); c.commit(); c.close(); return jsonify(ok=True)
@app.post('/api/delete/<int:i>')
def delete(i):
 c=db(); c.execute('DELETE FROM deposits WHERE id=?',(i,)); audit(c,'User','Delete FD',f'Record {i} removed'); c.commit(); c.close(); return jsonify(ok=True)
@app.get('/api/export.csv')
def export():
 c=db(); rows=c.execute('SELECT * FROM deposits').fetchall(); c.close(); out=io.StringIO(); w=csv.writer(out); w.writerow(rows[0].keys() if rows else ['No records']); [w.writerow(tuple(r)) for r in rows]; return app.response_class(out.getvalue(),mimetype='text/csv',headers={'Content-Disposition':'attachment; filename=fd_reconciliation.csv'})

HTML=r'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>FD Reconciliation</title><style>
:root{--ink:#17251f;--muted:#6d7973;--line:#e4e9e5;--bg:#f6f8f5;--green:#176c51;--mint:#e6f3ec;--amber:#a26017;--red:#b7483f;--white:#fff}*{box-sizing:border-box}body{margin:0;background:var(--bg);font:14px/1.45 Inter,Segoe UI,Arial;color:var(--ink)}button,input,select,textarea{font:inherit}button{cursor:pointer;border:0}.layout{display:grid;grid-template-columns:228px 1fr;min-height:100vh}.side{background:#122b24;color:#eaf2ed;padding:25px 16px}.brand{font-size:18px;font-weight:700;display:flex;gap:10px;align-items:center;padding:0 10px 30px}.logo{height:34px;width:34px;background:#67bb91;border-radius:10px;display:grid;place-items:center;color:#12352a;font-weight:800}.label{font-size:10px;text-transform:uppercase;letter-spacing:1.2px;color:#92aa9f;padding:0 10px;margin:18px 0 8px}.nav{padding:10px;border-radius:8px;color:#d5e2db;margin:3px 0;cursor:pointer}.nav.active,.nav:hover{background:#26473b;color:white}.foot{position:fixed;bottom:22px;color:#aac0b4;font-size:12px}.main{padding:30px 36px;max-width:1600px;width:100%;margin:auto}.top{display:flex;justify-content:space-between;align-items:center;margin-bottom:24px}.eyebrow{color:var(--green);font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:1.4px}.top h1{font-size:26px;margin:3px 0}.sub{color:var(--muted)}.btn{padding:10px 14px;border-radius:8px;background:var(--green);color:white;font-weight:600}.btn.light{background:white;color:var(--ink);border:1px solid var(--line)}.btn.sm{padding:7px 10px;font-size:12px}.datebox{display:flex;align-items:center;gap:8px;background:white;padding:8px 12px;border:1px solid var(--line);border-radius:8px}.datebox input{border:0;width:130px}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.card,.panel{background:white;border:1px solid var(--line);border-radius:12px;padding:17px}.card .cap{color:var(--muted);font-size:12px}.card strong{display:block;font-size:24px;margin:5px 0}.card small{color:var(--muted)}.card .ico{float:right;color:var(--green);background:var(--mint);padding:8px;border-radius:8px}.grid{display:grid;grid-template-columns:1.6fr 1fr;gap:14px;margin-top:16px}.panel h2{font-size:15px;margin:0 0 12px}.tools{display:flex;justify-content:space-between;gap:10px;align-items:center;margin:22px 0 12px}.tools h2{margin:0;font-size:17px}.toolsRight{display:flex;gap:8px}.search,.filter{background:white;border:1px solid var(--line);padding:9px 11px;border-radius:8px}.search{width:220px}table{width:100%;border-collapse:collapse;background:white}th{text-align:left;font-size:10px;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);padding:11px 10px;border-bottom:1px solid var(--line)}td{padding:11px 10px;border-bottom:1px solid #eff2ef;font-size:12px}tr:hover td{background:#fbfcfb}.pill{display:inline-block;border-radius:20px;padding:4px 9px;font-size:11px;font-weight:700}.Matched{background:#e7f5ec;color:#21724d}.Mismatch{background:#fff0e5;color:#a85f17}.Needs{background:#eef0f1;color:#667}.Restricted{background:#fce8e8;color:#a43737}.tablewrap{overflow:auto}.barline{display:flex;align-items:center;gap:10px;margin:13px 0;font-size:12px}.barline b{width:95px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.track{height:8px;background:#edf1ed;flex:1;border-radius:10px;overflow:hidden}.fill{height:100%;background:#56a980;border-radius:10px}.barline span{width:86px;text-align:right;color:var(--muted)}.chart{height:118px;display:flex;align-items:end;gap:14px;border-bottom:1px solid var(--line);padding:0 8px}.col{flex:1;background:#70b28d;border-radius:5px 5px 0 0;min-height:6px;position:relative}.col em{position:absolute;top:100%;font-style:normal;font-size:10px;color:var(--muted);white-space:nowrap;left:50%;transform:translateX(-50%);padding-top:6px}.barsTitle{margin-top:27px;color:var(--muted);font-size:11px}.modal{position:fixed;inset:0;background:#10251fcc;display:none;align-items:center;justify-content:center;padding:20px;z-index:5}.modal.show{display:flex}.dialog{background:white;border-radius:14px;width:min(800px,100%);max-height:94vh;overflow:auto;padding:22px}.dialogHead{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}.dialogHead h2{margin:0}.x{background:transparent;font-size:22px;color:var(--muted)}.form{display:grid;grid-template-columns:repeat(3,1fr);gap:11px}.field label{display:block;font-size:11px;color:var(--muted);margin:0 0 4px}.field input,.field select,.field textarea{width:100%;border:1px solid var(--line);border-radius:7px;padding:9px;background:white}.field.full{grid-column:1/-1}.actions{display:flex;justify-content:end;gap:8px;margin-top:18px}.hint{font-size:11px;color:var(--muted)}.auditrow{padding:9px 0;border-bottom:1px solid var(--line);font-size:12px}.auditrow small{display:block;color:var(--muted)}.toast{position:fixed;right:20px;bottom:20px;background:#163a2e;color:white;padding:13px 18px;border-radius:9px;display:none;z-index:9}.empty{padding:32px;text-align:center;color:var(--muted)}@media(max-width:850px){.layout{grid-template-columns:1fr}.side{display:none}.main{padding:20px}.cards{grid-template-columns:repeat(2,1fr)}.grid{grid-template-columns:1fr}.form{grid-template-columns:repeat(2,1fr)}}
</style></head><body><div class="layout"><aside class="side"><div class="brand"><div class="logo">FD</div>LedgerLens</div><div class="label">Workspace</div><div class="nav active">▦ &nbsp; Overview</div><div class="nav" onclick="document.querySelector('#records').scrollIntoView()">≋ &nbsp; Deposit schedule</div><div class="nav" onclick="openModal('reviewModal')">✓ &nbsp; Review sign-off</div><div class="label">Tools</div><div class="nav" onclick="openModal('uploadModal')">⇧ &nbsp; Import files</div><div class="nav" onclick="document.querySelector('#audit').scrollIntoView()">◷ &nbsp; Audit trail</div><div class="foot">Local workspace · Data stays on this device</div></aside><main class="main"><header class="top"><div><div class="eyebrow">Audit workspace</div><h1>Fixed deposit reconciliation</h1><div class="sub">Reconcile bank evidence against books and review exceptions.</div></div><div style="display:flex;gap:9px;align-items:center"><label class="datebox">Audit date <input id="auditDate" type="date"></label><button class="btn" onclick="openModal('entryModal')">＋ Add deposit</button></div></header><section class="cards"><div class="card"><span class="ico">▤</span><div class="cap">Bank deposits</div><strong id="total">₹0</strong><small id="bankCount">Across 0 banks</small></div><div class="card"><span class="ico">✓</span><div class="cap">Matched</div><strong id="matched">0</strong><small>Within your tolerance</small></div><div class="card"><span class="ico">!</span><div class="cap">Needs attention</div><strong id="exceptions">0</strong><small>Mismatch or needs review</small></div><div class="card"><span class="ico">◷</span><div class="cap">Restricted deposits</div><strong id="restricted">0</strong><small>Lien / collateral flagged</small></div></section><section class="grid"><div class="panel"><h2>Deposits by bank</h2><div id="banks"></div></div><div class="panel"><h2>Maturity ladder</h2><div class="chart" id="chart"></div><div class="barsTitle">Maturing in the next 12 months</div></div></section><section id="records"><div class="tools"><h2>Deposit schedule <span class="hint" id="count"></span></h2><div class="toolsRight"><input class="search" id="q" placeholder="Search bank or FD number…" oninput="draw()"><select class="filter" id="statusFilter" onchange="draw()"><option value="">All statuses</option><option>Matched</option><option>Mismatch</option><option>Needs review</option><option>Awaiting bank</option><option>Awaiting book</option></select><a class="btn light sm" href="/api/export.csv">Export CSV</a></div></div><div class="panel" style="padding:0"><div class="tablewrap"><table><thead><tr><th>Bank / FD number</th><th>Principal</th><th>Bank balance</th><th>Book balance</th><th>Variance</th><th>Rate</th><th>Maturity</th><th>Status</th><th></th></tr></thead><tbody id="rows"></tbody></table><div class="empty" id="empty">No deposits yet. Import a file or add an FD to begin.</div></div></div></section><section class="panel" id="audit" style="margin-top:16px"><h2>Recent activity</h2><div id="auditRows"></div></section></main></div>
<div class="modal" id="uploadModal"><div class="dialog"><div class="dialogHead"><div><h2>Import and reconcile source data</h2><div class="hint">Upload bank evidence and book/Tally data in separate panels. Matching runs after both are imported.</div></div><button class="x" onclick="closeModal('uploadModal')">×</button></div><div class="field" style="max-width:220px;margin-bottom:14px"><label>Matching tolerance</label><input id="matchTolerance" type="number" min="0" step="0.01" value="1"></div><div class="grid" style="grid-template-columns:1fr 1fr;gap:14px"><div class="panel"><h2>1 · Bank statements / advice</h2><p class="hint">Choose bank statement, deposit advice, or confirmation files.</p><div class="field"><label>Default bank name (if not in file)</label><input id="bankDefault" placeholder="e.g. SBI"></div><input id="bankFiles" style="margin-top:12px" type="file" multiple accept=".xlsx,.xlsm,.csv,.tsv,.xml,.docx,.pdf,.png,.jpg,.jpeg,.tif,.tiff"><div id="bankUploadMsg" class="hint" style="margin-top:8px"></div><button class="btn" style="margin-top:12px" onclick="uploadFiles('bank')">Import bank files</button></div><div class="panel"><h2>2 · Tally / book data</h2><p class="hint">Choose Tally exports or other ledger/trial-balance files.</p><div class="field"><label>Default bank name (optional)</label><input id="bookDefault" placeholder="Use if export has no bank column"></div><input id="bookFiles" style="margin-top:12px" type="file" multiple accept=".xlsx,.xlsm,.csv,.tsv,.xml,.docx,.pdf,.png,.jpg,.jpeg,.tif,.tiff"><div id="bookUploadMsg" class="hint" style="margin-top:8px"></div><button class="btn" style="margin-top:12px" onclick="uploadFiles('books')">Import Tally / book files</button></div></div><p class="hint">Matching uses FD/deposit number first. Fallback requires bank, amount within tolerance, and compatible start or maturity dates. Ambiguous or incomplete records remain open for review. Supported extraction varies by format; scanned images need local OCR setup and field confirmation.</p><div class="actions"><button class="btn light" onclick="closeModal('uploadModal')">Done</button></div></div></div>
<div class="modal" id="entryModal"><div class="dialog"><div class="dialogHead"><div><h2>Add / reconcile deposit</h2><div class="hint">FD number is the primary key. Amount, dates and bank support fallback matching.</div></div><button class="x" onclick="closeModal('entryModal')">×</button></div><form id="entryForm" class="form"><div class="field"><label>Bank *</label><input name="bank" required></div><div class="field"><label>Deposit / receipt number</label><input name="deposit_no"></div><div class="field"><label>Currency</label><select name="currency"><option>INR</option><option>USD</option><option>EUR</option><option>Other</option></select></div><div class="field"><label>Principal</label><input name="principal" type="number" step="0.01"></div><div class="field"><label>Bank balance at audit date</label><input name="bank_balance" type="number" step="0.01"></div><div class="field"><label>Book / ledger balance</label><input name="book_balance" type="number" step="0.01"></div><div class="field"><label>Interest rate (% p.a.)</label><input name="rate" type="number" step="0.001"></div><div class="field"><label>Start / acquisition date</label><input name="start_date" type="date"></div><div class="field"><label>Maturity date</label><input name="maturity" type="date"></div><div class="field"><label>Book interest accrued</label><input name="interest_book" type="number" step="0.01"></div><div class="field"><label>Restricted / lien marked?</label><select name="restricted"><option>No</option><option>Yes — lien</option><option>Yes — collateral</option><option>Yes — margin money</option><option>Review required</option></select></div><div class="field"><label>Difference tolerance</label><input name="tolerance" value="1" type="number" min="0" step="0.01"></div><div class="field full"><label>Difference reason / supporting reference</label><textarea name="reason" rows="2" placeholder="Explain variance, evidence reference, owner and target resolution date"></textarea></div><div class="field"><label>Preparer</label><input name="actor" placeholder="Name"></div><div class="field"><label>Notes / source document reference</label><input name="notes"></div><div class="field"><label>Classification note</label><input placeholder="Current/non-current; cash equivalent basis"></div></form><div class="actions"><button class="btn light" onclick="closeModal('entryModal')">Cancel</button><button class="btn" onclick="saveDeposit()">Save reconciliation</button></div></div></div>
<div class="modal" id="reviewModal"><div class="dialog" style="max-width:540px"><div class="dialogHead"><h2>Preparer & reviewer sign-off</h2><button class="x" onclick="closeModal('reviewModal')">×</button></div><div class="form" style="grid-template-columns:1fr 1fr"><div class="field"><label>Preparer name</label><input id="preparer"></div><div class="field"><label>Date completed</label><input id="prepared" type="date"></div><div class="field"><label>Reviewer name</label><input id="reviewer"></div><div class="field"><label>Review date</label><input id="reviewed" type="date"></div><div class="field full"><label>Review note</label><textarea id="reviewNote"></textarea></div></div><div class="actions"><button class="btn" onclick="saveReview()">Save sign-off</button></div></div></div><div class="toast" id="toast"></div>
<script>
let data={rows:[],audit:[]};const money=n=>new Intl.NumberFormat('en-IN',{style:'currency',currency:'INR',maximumFractionDigits:0}).format(+n||0);function openModal(x){document.getElementById(x).classList.add('show')}function closeModal(x){document.getElementById(x).classList.remove('show')}function toast(s){let t=document.getElementById('toast');t.textContent=s;t.style.display='block';setTimeout(()=>t.style.display='none',3200)}
async function refresh(){data=await(await fetch('/api/data')).json();draw();document.getElementById('auditRows').innerHTML=data.audit.map(a=>`<div class="auditrow">${esc(a.action)} · ${esc(a.detail)}<small>${esc(a.actor)} · ${new Date(a.at).toLocaleString()}</small></div>`).join('')||'<div class="hint">No activity recorded yet.</div>';if(data.review.preparer){preparer.value=data.review.preparer;prepared.value=data.review.prepared;reviewer.value=data.review.reviewer;reviewed.value=data.review.reviewed;reviewNote.value=data.review.note||''}}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}function draw(){let rs=data.rows,q=document.getElementById('q').value.toLowerCase(),sf=document.getElementById('statusFilter').value;let total=rs.reduce((a,r)=>a+(+r.bank_balance||0),0),banks=[...new Set(rs.map(r=>r.bank).filter(Boolean))];document.getElementById('total').textContent=money(total);document.getElementById('bankCount').textContent=`Across ${banks.length} bank${banks.length===1?'':'s'}`;document.getElementById('matched').textContent=rs.filter(r=>r.status==='Matched').length;document.getElementById('exceptions').textContent=rs.filter(r=>r.status!=='Matched').length;document.getElementById('restricted').textContent=rs.filter(r=>r.restricted&&r.restricted!=='No').length;document.getElementById('count').textContent=`(${rs.length})`;let bankTotals={};rs.forEach(r=>bankTotals[r.bank||'Unspecified']=(bankTotals[r.bank||'Unspecified']||0)+(+r.bank_balance||0));let max=Math.max(1,...Object.values(bankTotals));document.getElementById('banks').innerHTML=Object.entries(bankTotals).map(([b,v])=>`<div class="barline"><b title="${esc(b)}">${esc(b)}</b><div class="track"><div class="fill" style="width:${v/max*100}%"></div></div><span>${money(v)}</span></div>`).join('')||'<div class="empty">Add deposits to see bank totals.</div>';let today=new Date(),months=Array.from({length:12},(_,i)=>new Date(today.getFullYear(),today.getMonth()+i,1)),vals=months.map((m,i)=>rs.filter(r=>{if(!r.maturity)return false;let d=new Date(r.maturity);return d.getFullYear()===m.getFullYear()&&d.getMonth()===m.getMonth()}).reduce((a,r)=>a+(+r.bank_balance||0),0)),mx=Math.max(1,...vals);document.getElementById('chart').innerHTML=months.map((m,i)=>`<div class="col" title="${money(vals[i])}" style="height:${Math.max(5,vals[i]/mx*100)}%"><em>${m.toLocaleString('en',{month:'short'})}</em></div>`).join('');let view=rs.filter(r=>(!q||`${r.bank} ${r.deposit_no}`.toLowerCase().includes(q))&&(!sf||r.status===sf));document.getElementById('empty').style.display=view.length?'none':'block';document.getElementById('rows').innerHTML=view.map(r=>`<tr><td><b>${esc(r.bank||'—')}</b><br><span class="hint">${esc(r.deposit_no||'No FD no.')}</span>${r.restricted&&r.restricted!=='No'?'<br><span class="pill Restricted">Restricted</span>':''}</td><td>${money(r.principal)} ${esc(r.currency)}</td><td>${money(r.bank_balance)}</td><td>${money(r.book_balance)}</td><td style="color:${Math.abs(r.bank_balance-r.book_balance)>(+toleranceDefault())?'#aa621a':'inherit'}">${money(r.bank_balance-r.book_balance)}</td><td>${(+r.rate||0).toFixed(2)}%</td><td>${esc(r.maturity||'—')}</td><td><span class="pill ${r.status==='Needs review'?'Needs':r.status}">${esc(r.status)}</span><br><span class="hint">${esc(r.reason||'')}</span></td><td><button class="btn light sm" onclick="removeRow(${r.id})">×</button></td></tr>`).join('')}
 function toleranceDefault(){return 1}async function saveDeposit(){let f=document.getElementById('entryForm'),d=Object.fromEntries(new FormData(f));d.audit_date=document.getElementById('auditDate').value;let x=await fetch('/api/deposit',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});if(x.ok){f.reset();closeModal('entryModal');await refresh();toast('Deposit saved. Review the variance and evidence.')}}async function removeRow(id){if(!confirm('Remove this deposit record?'))return;await fetch('/api/delete/'+id,{method:'POST'});refresh()}
async function uploadFiles(){let fs=document.getElementById('files').files;if(!fs.length)return toast('Choose one or more files first.');let fd=new FormData();for(let f of fs)fd.append('files',f);let r=await(await fetch('/api/upload',{method:'POST',body:fd})).json();document.getElementById('uploadMsg').innerHTML=`Imported ${r.imported} row(s).<br>${r.messages.map(esc).join('<br>')}`;await refresh();toast(`${r.imported} row(s) imported; confirm values before relying on matching.`)}async function saveReview(){let d={preparer:preparer.value,prepared:prepared.value,reviewer:reviewer.value,reviewed:reviewed.value,note:reviewNote.value};let r=await fetch('/api/review',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)}),j=await r.json();if(!r.ok)return toast(j.error);closeModal('reviewModal');refresh();toast('Review sign-off saved.')}
async function uploadFiles(kind){let isBank=kind==='bank',input=document.getElementById(isBank?'bankFiles':'bookFiles'),fs=input.files;if(!fs.length)return toast('Choose one or more files for this source first.');let fd=new FormData();fd.append('kind',kind);fd.append('tolerance',document.getElementById('matchTolerance').value||'1');fd.append('default_bank',document.getElementById(isBank?'bankDefault':'bookDefault').value);for(let f of fs)fd.append('files',f);let r=await(await fetch('/api/upload',{method:'POST',body:fd})).json();let msg=document.getElementById(isBank?'bankUploadMsg':'bookUploadMsg');msg.textContent=`Imported ${r.imported}; matched ${r.matched}; mismatches ${r.mismatched}; awaiting counterpart ${r.unpaired}; ambiguous ${r.ambiguous}. ${r.messages.join(' ')}`;input.value='';await refresh();toast(`${r.imported} ${isBank?'bank':'book'} row(s) processed. Review unmatched and ambiguous records.`)}
document.getElementById('auditDate').value=new Date().toISOString().slice(0,10);refresh();
</script></body></html>'''
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT','5000')),debug=False)



