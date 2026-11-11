#!/usr/bin/env python3
"""Varsity Zone 11 Nov 2026 re-audit: 2.1 headings and 2.2 metadata.
Place the seven named XLSX files alongside this script and run with Python.
Produces date-stamped CSV reports and ZIP without overwriting earlier runs.
"""
from collections import Counter
from pathlib import Path
from urllib.parse import urlsplit,unquote
import csv,re,zipfile,xml.etree.ElementTree as ET

BASE=Path(__file__).resolve().parent
OUTPUT_STEM='vz_migration_audits_2_1_2_2_reports_11-Nov-2026'
OUT=BASE/OUTPUT_STEM
run=2
while OUT.exists() or (BASE/(OUT.name+'.zip')).exists():
    OUT=BASE/f'{OUTPUT_STEM}_run{run}'
    run+=1
OUT.mkdir()
FILES={
'p_html':'production_crawl2_internal_html.xlsx',
'p_h1':'production_crawl2_all_h1.xlsx',
'p_h2':'production_crawl2_all_h2.xlsx',
's_html':'vz-staging-crawl4-internal-11-nov-2026.xlsx',
's_h1':'staging_crawl4_all_h1.xlsx',
's_h2':'staging_crawl4_all_h2.xlsx',
}
INPUT=BASE/'vz-staging-crawl4-internal-11-nov-2026.xlsx'
missing=[str(BASE/name) for name in FILES.values() if not (BASE/name).is_file()]
if not INPUT.is_file():missing.append(str(INPUT))
if missing:raise FileNotFoundError('Missing required input files:\n'+'\n'.join(missing))
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def excel_rows(path):
    """Stream first XLSX worksheet without third-party dependencies."""
    with zipfile.ZipFile(path) as z:
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                for _,el in ET.iterparse(f,events=('end',)):
                    if el.tag==NS+'si':
                        shared.append(''.join(t.text or '' for t in el.iter(NS+'t')));el.clear()
        with z.open('xl/worksheets/sheet1.xml') as f:
            for _,el in ET.iterparse(f,events=('end',)):
                if el.tag!=NS+'row':continue
                cells={}
                for c in el.iter(NS+'c'):
                    ref=c.get('r','');match=re.match(r'[A-Z]+',ref)
                    if not match:continue
                    idx=0
                    for ch in match.group():idx=idx*26+ord(ch)-64
                    v=c.find(NS+'v');typ=c.get('t');inline=c.find(NS+'is')
                    val=v.text if v is not None and v.text is not None else ''
                    if typ=='s' and val:val=shared[int(val)]
                    elif typ=='inlineStr' and inline is not None:val=''.join(t.text or '' for t in inline.iter(NS+'t'))
                    cells[idx-1]=val
                if cells:yield [cells.get(i,'') for i in range(max(cells)+1)]
                el.clear()

def records(name):
    rows=excel_rows(BASE/FILES[name]);headers=next(rows)
    for row in rows:
        yield {h:(row[i] if i<len(row) else '') for i,h in enumerate(headers)}

def key(url):
    p=urlsplit(str(url).strip());host=p.netloc.lower().split(':')[0]
    if host not in ('www.varsityzone.com','varsityzone.com','varsityzonedev.wpenginepowered.com'):return None
    return unquote(p.path or '/').rstrip('/').lower() or '/'

def write(name,headers,rows):
    path=OUT/name
    with path.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f);w.writerow(headers);w.writerows(rows)
    print(name,len(rows));return path

rows=excel_rows(INPUT)
headers=next(rows)
column=headers.index('Address') if 'Address' in headers else 0
scope=set()
for row in rows:
    if len(row)>column:
        k=key(row[column])
        if k is not None:scope.add(k)
print('Crawl 4 URL scope:',len(scope))

def html(name):
    result={}
    for r in records(name):
        k=key(r.get('Address',''))
        if k in scope and 'html' in r.get('Content Type','').lower():
            result[k]=r
    return result
P=html('p_html');S=html('s_html')

def headings(name,kind):
    result={}
    for r in records(name):
        k=key(r.get('Address',''))
        if k not in scope:continue
        columns=sorted((c for c in r if re.fullmatch(kind+r'-\d+',c)),key=lambda c:int(c.split('-')[1]))
        result[k]=([r[c].strip() for c in columns if r[c].strip()],r.get('Occurrences',''))
    return result
PH1=headings('p_h1','H1');SH1=headings('s_h1','H1')
PH2=headings('p_h2','H2');SH2=headings('s_h2','H2')

def status(k):
    if k not in P:return 'NOT_IN_PRODUCTION_CRAWL2'
    if k not in S:return 'STAGING_HTML_MISSING'
    if str(S[k].get('Status Code','')).split('.')[0]!='200':return 'STAGING_NON_200'
    return 'COMPARABLE'

coverage=[];hrows=[];mrows=[];exceptions=[]
for k in sorted(scope):
    state=status(k);pr=P.get(k,{});sr=S.get(k,{})
    coverage.append([k,state,pr.get('Status Code',''),sr.get('Status Code',''),pr.get('Address',''),sr.get('Address','')])
    if state!='COMPARABLE':
        exceptions.append(['COVERAGE',k,state,'','', 'Investigate URL response or baseline coverage'])
    for label,prod,stag in [('H1',PH1,SH1),('H2',PH2,SH2)]:
        a,ac=prod.get(k,([],''));b,bc=stag.get(k,([],''))
        if state!='COMPARABLE':diff='NOT_COMPARABLE'
        elif k not in prod or k not in stag:diff='HEADING_EXPORT_MISSING'
        elif a==b:diff='MATCH'
        elif Counter(a)==Counter(b):diff='ORDER_CHANGED'
        elif not a and b:diff='ADDED'
        elif a and not b:diff='REMOVED'
        else:diff='CONTENT_OR_COUNT_CHANGED'
        hrows.append([k,label,diff,len(a),len(b),' | '.join(a),' | '.join(b),ac,bc])
        if diff not in ('MATCH','NOT_COMPARABLE'):
            exceptions.append(['2.1',k,label+'_'+diff,' | '.join(a),' | '.join(b),'Review heading parity'])
    for field,label in [('Title 1','TITLE'),('Meta Description 1','META_DESCRIPTION')]:
        a=str(pr.get(field,'')).strip();b=str(sr.get(field,'')).strip()
        if state!='COMPARABLE':diff='NOT_COMPARABLE'
        elif a==b:diff='MATCH'
        elif a and not b:diff='MISSING_ON_STAGING'
        elif not a and b:diff='ADDED_ON_STAGING'
        else:diff='CHANGED'
        mrows.append([k,label,diff,a,b])
        if diff not in ('MATCH','NOT_COMPARABLE'):
            exceptions.append(['2.2',k,label+'_'+diff,a,b,'Review metadata parity'])

write('00_url_coverage.csv',['Path','Status','Production HTTP','Staging HTTP','Production URL','Staging URL'],coverage)
write('audit_2_1_headings.csv',['Path','Heading','Classification','Production Count','Staging Count','Production Headings','Staging Headings','Production Reported Occurrences','Staging Reported Occurrences'],hrows)
write('audit_2_2_metadata.csv',['Path','Field','Classification','Production Value','Staging Value'],mrows)
write('developer_exceptions.csv',['Audit','Path','Classification','Production','Staging','Action'],exceptions)
summary=[['Metric','Value'],['Input URLs',len(scope)],['Production HTML pages in scope',len(P)],['Staging HTML pages in scope',len(S)],['Comparable pages',sum(status(k)=='COMPARABLE' for k in scope)],['Heading exceptions',sum(r[2] not in ('MATCH','NOT_COMPARABLE') for r in hrows)],['Metadata exceptions',sum(r[2] not in ('MATCH','NOT_COMPARABLE') for r in mrows)]]
write('summary.csv',summary[0],summary[1:])
ZIP_PATH=BASE/(OUT.name+'.zip')
with zipfile.ZipFile(ZIP_PATH,'w',compression=zipfile.ZIP_DEFLATED) as z:
    for p in OUT.glob('*.csv'):z.write(p,arcname=p.name)
print('Complete:',ZIP_PATH)
