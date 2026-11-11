"""VZ Audit 3: staging XML/HTML sitemap allocation reaudit and Monday regression.
Run: %run vz_audit3_xml_sitemap_reaudit_11_nov_2026.py
Dependencies: pandas, openpyxl, beautifulsoup4, lxml.
"""
from pathlib import Path
from urllib.parse import urlsplit, unquote
from collections import defaultdict, Counter
import re, zipfile, xml.etree.ElementTree as ET
import pandas as pd
from bs4 import BeautifulSoup

BASE=Path.cwd()
STAGING='https://varsityzonedev.wpenginepowered.com'
PRODUCTION_XML=BASE/'scorpion_xml_sitemap.xml'
LOCATIONS_HTML=BASE/'vz-locations-rendered-html.html'
SITEMAP_ZIP=BASE/'WordPress-11-Nov-2026.zip'
SITEMAP_DIRECTORY=None  # Optional folder override; normally use the dated ZIP above.
MONDAY_REPORT=BASE/'vz-staging-audit3-xml-sitemap-allocation-09-Nov-2026.xlsx'
OUTPUT=BASE/'vz-staging-audit3-xml-sitemap-allocation-reaudit-11-Nov-2026.xlsx'
AUDIT1_URL_LIST=None  # Set to fixed 1,310-URL CSV if Monday Audit 3 used that inventory.


def path_of(value):
    value=str(value).strip(); p=urlsplit(value)
    path=unquote(p.path if (p.scheme or p.netloc) else value.split('?',1)[0].split('#',1)[0])
    return (re.sub('/+','/','/'+path.lstrip('/')).rstrip('/') or '/')

def normalized(value):
    p=path_of(value)
    return STAGING+('' if p=='/' else p)

def parse_xml(data):
    root=ET.fromstring(data)
    kind=root.tag.rsplit('}',1)[-1].lower()
    if kind not in ('urlset','sitemapindex'): raise ValueError('Unexpected XML root: '+kind)
    return kind,[x.text.strip() for x in root.iter() if x.tag.rsplit('}',1)[-1].lower()=='loc' and x.text and x.text.strip()]

def parse_saved_html(data,filename):
    soup=BeautifulSoup(data,'lxml')
    # Browser-saved XML often becomes an HTML document containing <div class="..."> wrappers
    # around the original XML markup; first recover actual <loc> nodes when available.
    locs=[x.get_text(' ',strip=True) for x in soup.find_all('loc')]
    if locs:
        is_index=bool(soup.find('sitemapindex'))
        return ('sitemapindex' if is_index else 'urlset'),locs
    # Rank Math browser-rendered HTML sitemaps expose URLs as table links, not <loc> tags.
    table_links=[]
    for a in soup.select('table a[href]'):
        href=(a.get('href') or '').strip()
        if href.startswith(('https://','http://')) and urlsplit(href).netloc.lower()==urlsplit(STAGING).netloc.lower():
            table_links.append(href)
    if table_links:
        is_index=('index' in filename.lower() or
                  all(urlsplit(u).path.lower().endswith('.xml') for u in table_links))
        return ('sitemapindex' if is_index else 'urlset'),list(dict.fromkeys(table_links))
    # Chrome XML viewer renders each <loc> as escaped markup in page text.
    text=soup.get_text(' ',strip=True)
    decoded=__import__('html').unescape(data.decode('utf-8',errors='replace'))
    regex=re.findall(r'<loc(?:\s[^>]*)?>(.*?)</loc\s*>',decoded,re.I|re.S)
    if regex:
        return ('sitemapindex' if re.search(r'<sitemapindex\b',decoded,re.I) else 'urlset'),[BeautifulSoup(x,'html.parser').get_text(' ',strip=True) for x in regex]
    # Chrome's XML tree viewer may split loc into HTML spans; inspect plain text for URL strings.
    if 'xml' in text.lower() and ('<loc' in text.lower()):
        raise ValueError('Saved HTML contains visible XML but no recoverable <loc> elements')
    raise ValueError('No XML <loc> entries recoverable from saved HTML; save actual XML or browser-rendered XML page')

def parse_sitemap(data,filename):
    try: return parse_xml(data)
    except (ET.ParseError,ValueError):
        if filename.lower().endswith(('.html','.htm')) or data.lstrip().lower().startswith((b'<!doctype html',b'<html')):
            return parse_saved_html(data,filename)
        raise

def load_expected():
    if AUDIT1_URL_LIST:
        p=Path(AUDIT1_URL_LIST);df=pd.read_excel(p) if p.suffix.lower() in ('.xlsx','.xls') else pd.read_csv(p)
        cols=[c for c in df.columns if 'address' in str(c).lower() or 'url' in str(c).lower()]
        if not cols: raise ValueError('No URL/Address column in '+p.name)
        urls=df[cols[0]].dropna().astype(str)
    else:
        if not PRODUCTION_XML.exists():
            alt=BASE/'scorpion_xml_sitemap(1).xml'
            if alt.exists(): p=alt
            else: raise FileNotFoundError('Missing scorpion_xml_sitemap.xml')
        else:p=PRODUCTION_XML
        kind,urls=parse_xml(p.read_bytes())
        if kind!='urlset':raise ValueError('Production XML must be a complete urlset; use AUDIT1_URL_LIST otherwise')
    return sorted({normalized(u) for u in urls if str(u).strip().startswith(('http://','https://','/'))})

def load_locations():
    if not LOCATIONS_HTML.exists():raise FileNotFoundError(LOCATIONS_HTML)
    soup=BeautifulSoup(LOCATIONS_HTML.read_bytes(),'html.parser')
    wrappers=soup.find_all('div',class_=lambda v:v and any(x.startswith('sfa') for x in (v if isinstance(v,list) else str(v).split())))
    if not wrappers:raise ValueError('No SFA location container found in locations HTML')
    found=set()
    for wrapper in wrappers:
        for a in wrapper.select('a[href]'):
            u=urlsplit(a['href'])
            if u.netloc and u.netloc.lower() not in ('varsityzonedev.wpenginepowered.com','www.varsityzone.com','varsityzone.com'):continue
            parts=[x for x in path_of(a['href']).split('/') if x]
            if len(parts)==1 and re.fullmatch(r'[a-z0-9-]+-[a-z]{2}',parts[0].lower()):found.add(parts[0].lower())
    if not found:raise ValueError('No location folders discovered')
    return sorted(found)

def load_sources():
    if SITEMAP_DIRECTORY is not None:
        root=Path(SITEMAP_DIRECTORY)
        if not root.is_dir():raise FileNotFoundError(root)
        files=[(str(p.relative_to(root)),p.read_bytes()) for p in root.rglob('*') if p.suffix.lower() in ('.xml','.html','.htm')]
    else:
        if not SITEMAP_ZIP.exists():raise FileNotFoundError('Missing Wednesday sitemap ZIP: '+str(SITEMAP_ZIP))
        with zipfile.ZipFile(SITEMAP_ZIP) as z:
            files=[(n,z.read(n)) for n in z.namelist() if not n.endswith('/') and Path(n).suffix.lower() in ('.xml','.html','.htm')]
    sources={}
    for name,data in files:
        key=Path(name).stem.lower()  # .xml and saved .html are equivalent for matching
        if key in sources:raise ValueError(f'Duplicate sitemap basename {key}: {name} and {sources[key][0]}')
        sources[key]=(Path(name).name,data)
    if not sources:raise ValueError('No XML/HTML sitemap files found')
    return sources

def make_key(row):
    # Stable identity: one issue per URL per expected/actual sitemap. Ignore details that can vary by run.
    return (str(row.get('Issue','')).strip(),normalized(row['URL']) if str(row.get('URL','')).strip() else '',
            str(row.get('Expected sitemap','')).strip().lower().removesuffix('.xml').removesuffix('.html'),
            str(row.get('Actual sitemap(s)','')).strip().lower().replace('.xml','').replace('.html',''))

def action(issue):
    return {'MISSING':'Add expected URL to its correct sitemap; confirm page exists',
            'WRONG SITEMAP':'Move URL to correct national/location sitemap',
            'DUPLICATE':'Remove duplicate sitemap entry, retaining correct allocation',
            'UNEXPECTED':'Investigate whether page exists on Scorpion; confirm intended migration scope before removing',
            'WRONG DOMAIN':'Change sitemap loc to staging domain',
            'TRAILING SLASH':'Use slashless staging URL convention',
            'SITEMAP FILE MISSING':'Supply missing child sitemap and rerun audit',
            'PARSE ERROR':'Fix sitemap export/format and rerun audit',
            'UNLISTED SITEMAP FILE':'Confirm index references the intended child sitemap',
            'UNKNOWN LOCATION':'Confirm correct location folder and sitemap allocation'}.get(issue,'Investigate and confirm sitemap parity')

def main():
    expected=load_expected(); locations=load_locations();locset=set(locations);sources=load_sources()
    print(f'Expected URLs: {len(expected)}; locations: {len(locations)}; supplied sitemap files: {len(sources)}')
    rows=[]; errors=[]; index_names=[]; index_seen=False
    for stem,(filename,data) in sorted(sources.items()):
        try:kind,urls=parse_sitemap(data,filename)
        except Exception as exc:
            errors.append({'File':filename,'Issue':str(exc)});continue
        if kind=='sitemapindex':
            index_seen=True
            index_names.extend(Path(urlsplit(u).path).stem.lower() for u in urls)
            continue
        for u in urls:
            rows.append({'Sitemap':stem+'.xml','Raw URL':u,'Normalized URL':normalized(u),
                         'Domain issue':urlsplit(u).netloc.lower()!=urlsplit(STAGING).netloc.lower(),
                         'Trailing slash issue':urlsplit(u).path.endswith('/') and urlsplit(u).path!='/'})
    if not rows:raise ValueError('No staging sitemap URL entries parsed; inspect source ZIP')
    index_names=list(dict.fromkeys(index_names)); indexed=set(index_names)
    missing_files=sorted(indexed-set(sources))
    by_url=defaultdict(list)
    for row in rows:by_url[row['Normalized URL']].append(row)
    def classify(u):
        first=path_of(u).strip('/').split('/')[0].lower()
        if first in locset:return 'Local',first,first+'-sitemap.xml'
        return 'National','','national-sitemap.xml'
    audit=[];exc=[]
    for u in expected:
        typ,slug,target=classify(u);occ=by_url.get(u,[]);actual=sorted({x['Sitemap'] for x in occ})
        status=('MISSING' if not occ else 'WRONG SITEMAP' if target not in actual else 'DUPLICATE' if len(occ)>1 else 'PASS')
        audit.append({'Expected staging URL':u,'Classification':typ,'Location folder':slug,'Expected sitemap':target,'Actual sitemap(s)':'; '.join(actual),'Occurrences':len(occ),'Status':status})
        if status!='PASS':exc.append({'Issue':status,'URL':u,'Expected sitemap':target,'Actual sitemap(s)':'; '.join(actual),'Details':''})
    for u,occ in by_url.items():
        typ,slug,target=classify(u);actual=sorted({x['Sitemap'] for x in occ})
        if u not in set(expected):
            exc.append({'Issue':'UNEXPECTED','URL':u,'Expected sitemap':target,'Actual sitemap(s)':'; '.join(actual),'Details':'Not in production inventory'})
            if len(occ)>1:exc.append({'Issue':'DUPLICATE','URL':u,'Expected sitemap':target,'Actual sitemap(s)':'; '.join(actual),'Details':f'{len(occ)} entries'})
            if target not in actual:exc.append({'Issue':'WRONG SITEMAP','URL':u,'Expected sitemap':target,'Actual sitemap(s)':'; '.join(actual),'Details':'Unexpected URL misallocated'})
    for r in rows:
        for flag,issue in [('Domain issue','WRONG DOMAIN'),('Trailing slash issue','TRAILING SLASH')]:
            if r[flag]:exc.append({'Issue':issue,'URL':r['Raw URL'],'Expected sitemap':'','Actual sitemap(s)':r['Sitemap'],'Details':'Raw XML loc differs from expected convention'})
    for name in missing_files:exc.append({'Issue':'SITEMAP FILE MISSING','URL':'','Expected sitemap':name+'.xml','Actual sitemap(s)':'','Details':'Index references missing child sitemap'})
    for e in errors:exc.append({'Issue':'PARSE ERROR','URL':'','Expected sitemap':e['File'],'Actual sitemap(s)':'','Details':e['Issue']})
    if index_seen:
        for stem in sorted(set(sources)-indexed):
            if stem not in ('wp_sitemap_index','sitemap_index'):
                exc.append({'Issue':'UNLISTED SITEMAP FILE','URL':'','Expected sitemap':stem+'.xml','Actual sitemap(s)':'','Details':'File not referenced in index'})
    else:
        exc.append({'Issue':'PARSE ERROR','URL':'','Expected sitemap':'Sitemap index','Actual sitemap(s)':'','Details':'No sitemap index parsed; allocation completeness cannot be fully verified'})
    audit_df=pd.DataFrame(audit)
    exception_df=pd.DataFrame(exc,columns=['Issue','URL','Expected sitemap','Actual sitemap(s)','Details'])
    inventory=pd.DataFrame([{'Sitemap file':stem+'.xml','In index':stem in indexed,'File supplied':stem in sources,'Parsed URL entries':sum(x['Sitemap']==stem+'.xml' for x in rows)} for stem in sorted(indexed|set(sources)) if stem not in ('wp_sitemap_index','sitemap_index')])
    valid_current=(not missing_files and not errors and index_seen)
    regression=[];baseline_available=False
    if MONDAY_REPORT.exists():
        old=pd.read_excel(MONDAY_REPORT,sheet_name='Exceptions').fillna('')
        required={'Issue','URL','Expected sitemap','Actual sitemap(s)'}
        if not required.issubset(old.columns):raise ValueError('Monday Exceptions sheet missing columns: '+str(required-set(old.columns)))
        baseline_available=True
        previous={make_key(x):x for x in old.to_dict('records')}
        current={make_key(x):x for x in exception_df.fillna('').to_dict('records')}
        for key in sorted(set(previous)|set(current)):
            status='Outstanding' if key in previous and key in current else 'New' if key in current else 'Resolved'
            source=current.get(key,previous.get(key))
            regression.append({'Regression':status,**{c:source.get(c,'') for c in exception_df.columns},'Developer action':action(source['Issue']) if status!='Resolved' else 'No action — verify remains fixed'})
    regression_df=pd.DataFrame(regression,columns=['Regression',*exception_df.columns,'Developer action'])
    if baseline_available:
        todo=regression_df[regression_df['Regression'].isin(['New','Outstanding'])].copy()
        todo.insert(0,'Priority',todo['Issue'].map(lambda x:'P1' if x in ('MISSING','WRONG SITEMAP','SITEMAP FILE MISSING','PARSE ERROR') else 'P2'))
    else:
        todo=exception_df.copy();todo.insert(0,'Priority',todo['Issue'].map(lambda x:'P1' if x in ('MISSING','WRONG SITEMAP','SITEMAP FILE MISSING','PARSE ERROR') else 'P2'))
        todo.insert(1,'Regression','Baseline unavailable');todo['Developer action']=todo['Issue'].map(action)
    summary=[('Audit','03 — XML Sitemap Completeness & Allocation'),('Run','11 November 2026'),('Expected URLs',len(expected)),('Valid local folders',len(locations)),('Sitemap entries',len(rows)),('Unique sitemap URLs',len(by_url)),('Sitemap index parsed',index_seen),('Missing child sitemap files',len(missing_files)),('Parse errors',len(errors)),('Current input complete',valid_current),('Monday report found',baseline_available),('Regression reliable',valid_current and baseline_available)]
    summary += [(f'Expected URLs: {k}',int(v)) for k,v in sorted(audit_df['Status'].value_counts().items())]
    summary += [(f'Exceptions: {k}',int(v)) for k,v in sorted(exception_df['Issue'].value_counts().items())]
    if baseline_available:summary += [(f'Regression: {k}',int(v)) for k,v in sorted(regression_df['Regression'].value_counts().items())]
    summary_df=pd.DataFrame(summary,columns=['Metric','Value'])
    with pd.ExcelWriter(OUTPUT,engine='openpyxl') as writer:
        sheets=[('Summary',summary_df),('Developer To Do',todo),('Monday vs Wednesday',regression_df),('Exceptions',exception_df),('Full URL Audit',audit_df),('Sitemap Inventory',inventory),('Raw Sitemap Entries',pd.DataFrame(rows)),('Location Folders',pd.DataFrame({'Location folder':locations}))]
        for name,df in sheets:
            df.to_excel(writer,sheet_name=name,index=False)
            ws=writer.sheets[name];ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
            from openpyxl.styles import Font,PatternFill
            for cell in ws[1]:cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='17365D')
            for col in ws.columns:
                letter=col[0].column_letter;mx=max((len(str(c.value or '')) for c in list(col)[:1000]),default=10)
                ws.column_dimensions[letter].width=min(max(mx+2,14),75)
    print('Report:',OUTPUT)
    print('Exceptions:',len(exception_df),'Developer actions:',len(todo))
    if not valid_current:print('WARNING: Wednesday sitemap inventory incomplete — do not interpret missing URLs as confirmed regressions')
    if not baseline_available:print('WARNING: Monday workbook missing — regression comparison skipped')
    return {'summary':summary_df,'exceptions':exception_df,'regression':regression_df,'developer_actions':todo,'output':OUTPUT}

if __name__=='__main__':
    audit3=main()
