from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from pathlib import Path
from openpyxl import load_workbook
from pypdf import PdfReader
from rapidfuzz import fuzz
import io, csv, re, json

app = FastAPI(title='Stock Statement Compilation Portal')
BASE_DIR = Path(__file__).resolve().parent.parent

@app.get('/', include_in_schema=False)
def home():
    return FileResponse(BASE_DIR / 'index.html')

@app.get('/health')
def health():
    return {'ok': True, 'service': 'stock-statement-compiler'}

def num(x):
    if x is None: return 0.0
    s=str(x).strip().replace(',','')
    if s in ('','-','—','–'): return 0.0
    try: return float(s)
    except: return 0.0

def norm_identity(s):
    s=str(s or '').upper().replace('’', "'").replace('&',' AND ')
    s=re.sub(r'[^A-Z0-9]+',' ',s)
    return re.sub(r'\s+',' ',s).strip()

def norm_customer_identity(s):
    s=norm_identity(s)
    # Keep distinctive business words; remove only location / legal-noise words.
    s=re.sub(r'\b(M S|MS|MESSRS|AKOLA|NAGPUR|AMRAVATI|WASHIM|BULDHANA|YAVATMAL|WARDHA|JALGAON|NASHIK|PUNE|AURANGABAD|CHHATRAPATI SAMBHAJINAGAR|HEADQUARTER|HQ|THE)\b',' ',s)
    return re.sub(r'\s+',' ',s).strip()

def identity_score(a,b,customer=False):
    na=norm_customer_identity(a) if customer else norm_identity(a)
    nb=norm_customer_identity(b) if customer else norm_identity(b)
    if not na or not nb: return 0
    if na==nb: return 100
    return round(max(fuzz.ratio(na,nb), fuzz.WRatio(na,nb)))

def clean_candidate(s):
    s=' '.join(str(s or '').replace('\u00a0',' ').split()).strip(' .:-;,|')
    return s

def extract_text_pdf(data):
    reader=PdfReader(io.BytesIO(data))
    chunks=[]
    for page in reader.pages[:5]:
        try: chunks.append(page.extract_text() or '')
        except Exception: pass
    return '\n'.join(chunks)

def extract_pdf_identity(data):
    text=extract_text_pdf(data)
    lines=[clean_candidate(x) for x in text.splitlines() if clean_candidate(x)]
    joined='\n'.join(lines)

    # 1) Explicit labels are the most reliable form.
    stock_patterns=[
        r'(?:STOCKIST|STOCKIST NAME|CUSTOMER|CUSTOMER NAME|PARTY|PARTY NAME|DEALER|DEALER NAME)\s*[:\-]\s*(.+)',
        r'(?:NAME OF STOCKIST|NAME OF CUSTOMER)\s*[:\-]\s*(.+)'
    ]
    hq_patterns=[r'(?:HQ|H\.Q\.|HEADQUARTER|HEAD QUARTER|LOCATION)\s*[:\-]\s*([A-Z][A-Z .&()\-/]+)']
    stock=''; hq=''; raw=''
    for line in lines[:80]:
        for pat in stock_patterns:
            m=re.search(pat,line,re.I)
            if m and len(clean_candidate(m.group(1)))>=3:
                stock=clean_candidate(m.group(1)); raw=line; break
        if stock: break
    for line in lines[:80]:
        for pat in hq_patterns:
            m=re.search(pat,line,re.I)
            if m:
                hq=clean_candidate(m.group(1)); break
        if hq: break

    # 2) Common Alkem stock statement header: CUSTOMER, HQ before report title.
    report_idx=next((i for i,x in enumerate(lines) if re.search(r'Stock\s*&\s*Sales\s*Report',x,re.I)),None)
    candidates=lines[:report_idx] if report_idx is not None else lines[:25]
    if not stock or not hq:
        for line in candidates:
            if re.search(r'Product\s+Name|ALKEM|Stock\s*&\s*Sales',line,re.I): continue
            m=re.match(r'^(.+?)\s*,\s*([A-Za-z][A-Za-z0-9 .&()\-/]*)$',line)
            if m:
                c=clean_candidate(m.group(1)); q=clean_candidate(m.group(2))
                if len(c)>=3 and len(q)>=2:
                    stock=stock or c; hq=hq or q; raw=raw or line
                    break

    # 3) Some exports put stockist and HQ on adjacent lines.
    if not stock and candidates:
        for i,line in enumerate(candidates[:12]):
            if re.search(r'(STOCK|CUSTOMER|PARTY|DEALER)',line,re.I) and not re.search(r'ALKEM|REPORT',line,re.I):
                nxt=candidates[i+1] if i+1<len(candidates) else ''
                if len(line)>=3:
                    stock=clean_candidate(re.sub(r'^(STOCKIST|CUSTOMER|PARTY|DEALER)\s*(NAME)?\s*[:\-]?\s*','',line,flags=re.I))
                    if not hq and re.fullmatch(r'[A-Za-z .&\-/]{2,30}',nxt): hq=clean_candidate(nxt)
                    raw=line
                    break

    return {'customer':stock,'hq':hq,'source':'PDF text/header' if stock or hq else 'PDF text not detected','raw_header':raw,'raw_text_preview':joined[:2000]}

def excel_values(data, ext):
    if ext in ('.xlsx','.xlsm'):
        wb=load_workbook(io.BytesIO(data),data_only=True,read_only=True)
        for ws in wb.worksheets[:3]:
            vals=[]
            for row in ws.iter_rows(min_row=1,max_row=30,max_col=15,values_only=True):
                vals.append([str(x).strip() if x is not None else '' for x in row])
            yield ws.title, vals
    elif ext=='.xls':
        try:
            import xlrd
        except Exception as e:
            raise ValueError('Legacy .XLS support requires xlrd; please redeploy with the updated requirements.txt') from e
        book=xlrd.open_workbook(file_contents=data,on_demand=True)
        for sh in book.sheets()[:3]:
            vals=[]
            for r in range(min(sh.nrows,40)):
                vals.append([str(sh.cell_value(r,c)).strip() for c in range(min(sh.ncols,20))])
            yield sh.name, vals

def find_labeled(values):
    stock=''; hq=''; raw=''
    flat=[]
    for row in values:
        flat.extend(row)
        line=' '.join(x for x in row if x)
        if not line: continue
        # cell labels and label:value in same row
        for pat in [r'(?:STOCKIST|STOCKIST NAME|CUSTOMER|CUSTOMER NAME|PARTY|PARTY NAME|DEALER|DEALER NAME)\s*[:\-]?\s*(.+)$']:
            m=re.search(pat,line,re.I)
            if m and not stock: stock=clean_candidate(m.group(1)); raw=line
        for pat in [r'(?:HQ|H\.Q\.|HEADQUARTER|HEAD QUARTER|LOCATION)\s*[:\-]?\s*(.+)$']:
            m=re.search(pat,line,re.I)
            if m and not hq: hq=clean_candidate(m.group(1))
    # Look at adjacent cells: [HQ, AKOLA], [CUSTOMER NAME, BALAJI...] etc.
    for r,row in enumerate(values):
        for c,cell in enumerate(row):
            label=norm_identity(cell)
            nxt=row[c+1] if c+1<len(row) else ''
            if not hq and label in ('HQ','H Q','HEADQUARTER','HEAD QUARTER','LOCATION') and nxt: hq=clean_candidate(nxt)
            if not stock and label in ('STOCKIST','STOCKIST NAME','CUSTOMER','CUSTOMER NAME','PARTY','PARTY NAME','DEALER','DEALER NAME') and nxt:
                stock=clean_candidate(nxt); raw=' | '.join(row)
    return stock,hq,raw

def extract_excel_identity(data,ext):
    all_rows=[]
    for sheet,values in excel_values(data,ext):
        stock,hq,raw=find_labeled(values)
        if stock or hq: return {'customer':stock,'hq':hq,'source':'Excel labels','raw_header':raw,'sheet':sheet}
        all_rows.extend(values)
    # Fallback: search first 20 rows for a comma-separated customer,HQ style header.
    for row in all_rows[:60]:
        line=' '.join(x for x in row if x)
        m=re.search(r'^(.+?)\s*,\s*([A-Za-z][A-Za-z .&()\-/]+)$',line)
        if m: return {'customer':clean_candidate(m.group(1)),'hq':clean_candidate(m.group(2)),'source':'Excel header','raw_header':line}
    return {'customer':'','hq':'','source':'Excel identity not detected','raw_header':''}

def identify_against_master(identity, master_rows):
    raw_customer=identity.get('customer',''); raw_hq=identity.get('hq','')
    if not raw_customer and not raw_hq:
        return {**identity,'customer':'','hq':'','confidence':0,'match_status':'NOT DETECTED'}
    best=None
    for row in master_rows:
        mc=str(row.get('CUSTOMER NAME','')); mh=str(row.get('HQ',''))
        cs=identity_score(raw_customer,mc,True) if raw_customer else 0
        hs=identity_score(raw_hq,mh,False) if raw_hq else 0
        # If only one side was found, don't invent the other from weak evidence.
        if raw_customer and raw_hq: score=round(cs*0.78+hs*0.22)
        elif raw_customer: score=cs
        else: score=hs
        if best is None or score>best['score']: best={'score':score,'customer':mc,'hq':mh,'cscore':cs,'hscore':hs}
    if not best or best['score']<70:
        return {**identity,'customer':'','hq':'','confidence':best['score'] if best else 0,'match_status':'NOT MATCHED'}
    status='EXACT' if best['score']>=98 and (not raw_hq or best['hscore']>=95) else ('HIGH CONFIDENCE' if best['score']>=90 else 'REVIEW')
    return {**identity,'customer':best['customer'],'hq':best['hq'],'confidence':round(best['score']),'customer_confidence':round(best['cscore']),'hq_confidence':round(best['hscore']),'match_status':status}

# Existing statement parsing retained.
def sku_norm(s):
    s=str(s or '').upper().replace('’', "'")
    s=s.replace('CILNIREM','CILNIKEM')
    s=re.sub(r'(?<=\d)SMG\b','5MG',s)
    s=s.replace('–','-').replace('—','-')
    s=re.sub(r'\b(?:TABLETS?|TABS?|CAPSULES?|CAPS?|ORAL|ER)\b',' ',s)
    s=re.sub(r"\b\d+(?:X1|X|TAB|TABS|,S|'S|S)\b",' ',s)
    return re.sub(r'[^A-Z0-9./+]+','',s)

def parse_pdf(data):
    reader=PdfReader(io.BytesIO(data)); rows=[]
    for page in reader.pages:
        text=page.extract_text() or ''
        for raw in text.splitlines():
            line=' '.join(str(raw).replace('\u00a0',' ').split())
            if not line or re.match(r'^[-_=]{5,}',line): continue
            if re.search(r'Product Name\s+Pack|Stock & Sales Report|Last Month Sales|Closing Value|Stk\.Value',line,re.I): continue
            dm=re.search(r'\b\d{2}/\d{2}/\d{2}\b',line)
            if not dm: continue
            pre=line[:dm.start()].strip(); toks=pre.split(); vals=[]; i=len(toks)-1
            while i>=0 and len(vals)<7:
                t=toks[i]
                try: v=0.0 if t in ('-','—','–') else float(t.replace(',',''))
                except: break
                vals.append(v); i-=1
            if len(vals)!=7: continue
            vals.reverse(); prefix=toks[:i+1]
            pack_idx=None; pack_re=r"\d+(?:\.\d+)?(?:X1|X|TAB|TABS|,TAB|,S|'S|S|CAP|CAPS)?"
            for j in range(len(prefix)-1,0,-1):
                if re.fullmatch(pack_re,prefix[j],re.I): pack_idx=j; break
            if pack_idx is None: continue
            sku=' '.join(prefix[:pack_idx]).strip()
            if len(sku)<3: continue
            lastsl,open_v,recd,sales,close,order,pend=vals
            rows.append({'source_sku':sku,'sec':sales,'close':close})
    return rows

def parse_csv(data):
    text=data.decode('utf-8-sig','replace'); out=[]; rows=csv.DictReader(io.StringIO(text))
    for r in rows:
        keys={re.sub(r'[^A-Z]','',str(k).upper()):k for k in r}
        def g(*names):
            for n in names:
                if n in keys:return r.get(keys[n])
            return None
        sku=g('PRODUCTNAME','SKU','PRODUCT','ITEM','PRODUCTDESCRIPTION') or ''
        if sku: out.append({'source_sku':str(sku),'sec':num(g('SALES','SECONDARYUNITS','SECONDARY')),'close':num(g('CLOSE','CLOSING','CLOSINGUNITS'))})
    return out

def parse_xlsx(data,ext='.xlsx'):
    for sheet,rows in excel_values(data,ext):
        if not rows: continue
        # Find header row, not necessarily row 1.
        hi=None; headers=None
        for r,row in enumerate(rows[:20]):
            h=[re.sub(r'[^A-Z]','',str(x or '').upper()) for x in row]
            if any(x in h for x in ('PRODUCTNAME','SKU','PRODUCT','ITEM')):
                hi=r; headers=h; break
        if hi is None: continue
        def idx(*names):
            for n in names:
                if n in headers:return headers.index(n)
            return None
        si=idx('PRODUCTNAME','SKU','PRODUCT','ITEM','PRODUCTDESCRIPTION'); sec_i=idx('SALES','SECONDARYUNITS','SECONDARY'); clo_i=idx('CLOSE','CLOSING','CLOSINGUNITS')
        if si is None: continue
        out=[]
        for rr in rows[hi+1:]:
            sku=rr[si] if si<len(rr) else ''
            if sku: out.append({'source_sku':str(sku),'sec':num(rr[sec_i]) if sec_i is not None and sec_i<len(rr) else 0,'close':num(rr[clo_i]) if clo_i is not None and clo_i<len(rr) else 0})
        if out:return out
    raise ValueError('Could not identify SKU/Product column in Excel')

def parse_statement(data,filename):
    ext=Path(filename).suffix.lower()
    if ext=='.pdf': return parse_pdf(data)
    if ext in ('.xlsx','.xlsm','.xls'): return parse_xlsx(data,ext)
    if ext=='.csv': return parse_csv(data)
    raise ValueError(f'Unsupported file type: {ext}')

def match_rows(rows,pool,aliases,threshold):
    out=[];reviews=[]
    for r in rows:
        key=sku_norm(r['source_sku']); chosen=aliases.get(key)
        if chosen:
            m=next((p for p in pool if str(p.get('SKU_NAME',''))==chosen),None)
            if m:
                out.append({**r,'master_sku':chosen,'suggestion':chosen,'pts':num(m.get('PTS')),'confidence':100,'status':'VALIDATED'});continue
        best=None
        for p in pool:
            name=str(p.get('SKU_NAME',''));score=max(fuzz.ratio(key,sku_norm(name)),fuzz.WRatio(key,sku_norm(name)))
            if best is None or score>best[0]:best=(score,p)
        if best:
            score,p=best;suggestion=str(p.get('SKU_NAME',''));item={**r,'master_sku':suggestion if score>=threshold else '','suggestion':suggestion,'pts':num(p.get('PTS')),'confidence':round(score),'status':'AUTO MATCH' if score>=threshold else 'REVIEW'}
        else:item={**r,'master_sku':'','suggestion':'','pts':0,'confidence':0,'status':'UNMATCHED'}
        if item['master_sku']:out.append(item)
        reviews.append(item)
    return out,reviews

@app.post('/api/identify')
@app.post('/identify')
async def identify_endpoint(file:UploadFile=File(...),master_json:str=Form('[]')):
    try:
        data=await file.read()
        if len(data)>4_450_000:raise HTTPException(413,'Statement exceeds Vercel request limit (~4.4 MB).')
        master_rows=json.loads(master_json or '[]')
        ext=Path(file.filename).suffix.lower()
        if ext=='.pdf': identity=extract_pdf_identity(data)
        elif ext in ('.xlsx','.xlsm','.xls'): identity=extract_excel_identity(data,ext)
        else:return {'file':file.filename,'customer':'','hq':'','confidence':0,'match_status':'UNSUPPORTED','source':'Manual selection'}
        return identify_against_master(identity,master_rows)
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f'{type(e).__name__}: {e}')

@app.post('/api/analyze')
@app.post('/analyze')
async def analyze_endpoint(file:UploadFile=File(...),pool_json:str=Form(...),aliases_json:str=Form('{}'),threshold:int=Form(82)):
    try:
        data=await file.read()
        if len(data)>4_450_000:raise HTTPException(413,'Statement exceeds Vercel request limit (~4.4 MB).')
        pool=json.loads(pool_json);aliases=json.loads(aliases_json or '{}')
        rows=parse_statement(data,file.filename)
        matched,reviews=match_rows(rows,pool,aliases,threshold)
        return {'file':file.filename,'rows':len(rows),'sec_total':sum(num(x['sec']) for x in rows),'close_total':sum(num(x['close']) for x in rows),'matched':matched,'reviews':reviews}
    except HTTPException:raise
    except Exception as e:raise HTTPException(500,f'{type(e).__name__}: {e}')
