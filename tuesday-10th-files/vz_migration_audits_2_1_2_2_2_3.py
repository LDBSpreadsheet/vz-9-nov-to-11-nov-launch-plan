#!/usr/bin/env python3
"""Varsity Zone migration parity audits 2.1 (headings), 2.2 (metadata), 2.3 (image alt).
Run: python vz_migration_audits_2_1_2_2_2_3.py
Place the eight named Screaming Frog XLSX files and the Crawl 3 URL CSV beside this script.
Uses Python standard library only; produces Excel-compatible UTF-8 CSV reports and a ZIP.
"""
from collections import defaultdict,Counter
from pathlib import Path
from urllib.parse import urlsplit,unquote
import csv,re,zipfile,xml.etree.ElementTree as ET

BASE=Path(__file__).resolve().parent
OUT=BASE/'vz_migration_audits_2_reports';OUT.mkdir(exist_ok=True)
FILES={
'p_html':'production_crawl2_internal_html.xlsx','s_html':'staging_crawl3_internal_html.xlsx',
'p_h1':'production_crawl2_all_h1.xlsx','s_h1':'staging_crawl3_all_h1.xlsx',
'p_h2':'production_crawl2_all_h2.xlsx','s_h2':'staging_crawl3_all_h2.xlsx',
'p_img':'production_crawl1_all_image_inlinks.xlsx','s_img':'staging_crawl3_all_image_inlinks.xlsx'}
INPUT=BASE/'vz-staging-crawl3-urls-10-nov-2026.csv'
if not INPUT.exists():INPUT=BASE/'vz-staging-crawl3-urls.csv'
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

if not INPUT.exists():raise FileNotFoundError('Missing Crawl 3 URL CSV next to script')
with INPUT.open(encoding='utf-8-sig',newline='') as f:
    r=csv.DictReader(f);column='Address' if 'Address' in r.fieldnames else r.fieldnames[0]
    scope={key(x[column]) for x in r if key(x[column]) is not None}
print('Crawl 3 URL scope:',len(scope))

def preferred_url(url, environment):
    """Prefer production trailing slash and staging no trailing slash (root excepted)."""
    parsed=urlsplit(str(url).strip())
    path=parsed.path or '/'
    if environment=='production':
        return path=='/' or path.endswith('/')
    return path=='/' or not path.endswith('/')

def select_record(result, k, r, environment):
    """Do not let discovered redirect variants overwrite the intended URL."""
    if k not in result:
        result[k]=r
        return
    old=result[k]
    new_pref=preferred_url(r.get('Address',''),environment)
    old_pref=preferred_url(old.get('Address',''),environment)
    if new_pref and not old_pref:
        result[k]=r
    elif new_pref==old_pref:
        new_200=str(r.get('Status Code','')).split('.')[0]=='200'
        old_200=str(old.get('Status Code','')).split('.')[0]=='200'
        if new_200 and not old_200:
            result[k]=r

def html(name):
    result={}
    environment='production' if name.startswith('p_') else 'staging'
    for r in records(name):
        k=key(r.get('Address',''))
        if k in scope and 'html' in r.get('Content Type','').lower():
            select_record(result,k,r,environment)
    return result
P=html('p_html');S=html('s_html')

def headings(name,kind):
    result={}
    environment='production' if name.startswith('p_') else 'staging'
    for r in records(name):
        k=key(r.get('Address',''))
        if k not in scope:continue
        columns=sorted((c for c in r if re.fullmatch(kind+r'-\d+',c)),key=lambda c:int(c.split('-')[1]))
        value=([r[c].strip() for c in columns if r[c].strip()],r.get('Occurrences',''))
        # H1/H2 exports may contain both canonical and redirected variants.
        if k not in result or (preferred_url(r.get('Address',''),environment) and
                               not preferred_url(result[k][2],environment)):
            result[k]=(value[0],value[1],r.get('Address',''))
    return {k:(v[0],v[1]) for k,v in result.items()}

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

def image_identity(url):
    path=unquote(urlsplit(url).path).lower();basename=path.rsplit('/',1)[-1]
    basename=re.sub(r'-(?:\d{2,5}x\d{2,5}|scaled)(?=\.)','',basename)
    basename=re.sub(r'\.(?:webp|avif|jpe?g|png|gif|svg)$','',basename)
    return basename

def images(name):
    bypage=defaultdict(list);seen=set()
    for r in records(name):
        k=key(r.get('Source',''))
        if k not in scope:continue
        environment='production' if name.startswith('p_') else 'staging'
        if not preferred_url(r.get('Source',''),environment):continue
        dest=r.get('Destination','').strip();alt=r.get('Alt Text','').strip()
        if not dest:continue
        token=(k,dest,alt)
        if token in seen:continue
        seen.add(token);bypage[k].append((dest,alt,image_identity(dest)))
    return bypage
PI=images('p_img');SI=images('s_img')
imgrows=[]
for k in sorted(scope):
    if status(k)!='COMPARABLE':continue
    a=PI.get(k,[]);b=SI.get(k,[])
    bybase=defaultdict(list)
    for j,item in enumerate(b):bybase[item[2]].append(j)
    used=set()
    for pu,pa,pbase in a:
        candidates=[j for j in bybase[pbase] if j not in used]
        if len(candidates)==1:
            j=candidates[0];used.add(j);su,sa,_=b[j]
            diff='MATCH' if pa==sa else ('ALT_ADDED' if not pa else 'ALT_REMOVED' if not sa else 'ALT_CHANGED')
            confidence='BASENAME_MATCH'
        elif len(candidates)>1:
            su=sa='';diff='AMBIGUOUS_IMAGE_MATCH';confidence='MULTIPLE_STAGING_IMAGES_SAME_BASENAME'
        else:
            su=sa='';diff='UNMATCHED_PRODUCTION_IMAGE';confidence='URL_OR_FILENAME_MAY_HAVE_CHANGED'
        imgrows.append([k,diff,confidence,pu,su,pa,sa])
        if diff!='MATCH':exceptions.append(['2.3',k,diff,pu+' | ALT: '+pa,su+' | ALT: '+sa,'Investigate image identity; do not assume missing image from filename alone'])
    for j,(su,sa,_) in enumerate(b):
        if j not in used:
            imgrows.append([k,'UNMATCHED_STAGING_IMAGE','MAY_BE_NEW_OR_RENAMED','',su,'',sa])
# Notes: images are matched conservatively by normalized filename, not visual identity.
write('00_url_coverage.csv',['Path','Status','Production HTTP','Staging HTTP','Production URL','Staging URL'],coverage)
write('audit_2_1_headings.csv',['Path','Heading','Classification','Production Count','Staging Count','Production Headings','Staging Headings','Production Reported Occurrences','Staging Reported Occurrences'],hrows)
write('audit_2_2_metadata.csv',['Path','Field','Classification','Production Value','Staging Value'],mrows)
write('audit_2_3_image_alt.csv',['Path','Classification','Match Confidence','Production Image','Staging Image','Production Alt','Staging Alt'],imgrows)
write('developer_exceptions.csv',['Audit','Path','Classification','Production','Staging','Action'],exceptions)
summary=[['Metric','Value'],['Input URLs',len(scope)],['Production HTML pages in scope',len(P)],['Staging HTML pages in scope',len(S)],['Comparable pages',sum(status(k)=='COMPARABLE' for k in scope)],['Heading exceptions',sum(r[2] not in ('MATCH','NOT_COMPARABLE') for r in hrows)],['Metadata exceptions',sum(r[2] not in ('MATCH','NOT_COMPARABLE') for r in mrows)],['Image rows requiring investigation',sum(r[1]!='MATCH' for r in imgrows)],['Important','Image basename matching is heuristic; unmatched does not prove missing']]
write('summary.csv',summary[0],summary[1:])
with zipfile.ZipFile(BASE/'vz_migration_audits_2_reports.zip','w',compression=zipfile.ZIP_DEFLATED) as z:
    for p in OUT.glob('*.csv'):z.write(p,arcname=p.name)
print('Complete:',BASE/'vz_migration_audits_2_reports.zip')
