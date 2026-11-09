"""Audit #4: live HTML sitemap allocation reconciliation.
Install: pip install requests beautifulsoup4 openpyxl
Place the Audit #3 workbook and vz-locations-rendered-html.html beside this script.
Run: python vz_staging_audit4_html_sitemap_query_safe.py
"""
from __future__ import annotations
import argparse, collections, datetime, pathlib, random, re, time
from urllib.parse import urljoin, urlsplit, unquote
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = 'https://varsityzonedev.wpenginepowered.com'
OUT = 'vz-staging-audit4-html-sitemap-allocation-09-Nov-2026.xlsx'
HEADERS = {'User-Agent': 'Mozilla/5.0 (compatible; VZMigrationAudit/4.0; internal QA)'}

def norm(raw):
    if not raw: return ''
    p = urlsplit(urljoin(BASE+'/', str(raw).strip()))
    if p.netloc.lower() not in {'varsityzonedev.wpenginepowered.com', 'www.varsityzone.com', 'varsityzone.com'}: return ''
    path = re.sub('/+', '/', unquote(p.path or '/')).rstrip('/') or '/'
    return BASE + path.lower()

def pathof(u): return urlsplit(u).path

def read_sheet(file):
    book = load_workbook(file, read_only=True, data_only=True)
    try:
        for sh in book.worksheets[:15]:
            rows = list(sh.iter_rows(min_row=1, max_row=min(sh.max_row, 12000),
                                     max_col=min(sh.max_column, 26), values_only=True))
            if rows and any('url' in str(v).lower() or 'address' in str(v).lower()
                            for v in rows[0]):
                yield sh.title, rows
    finally:
        book.close()

def expected_from_audit3(file):
    for name, rows in read_sheet(file):
        if not any(s in name.lower() for s in ('full', 'audit', 'allocation', 'expected')): continue
        for head_idx, row in enumerate(rows[:10]):
            headers = [str(v or '').strip().lower() for v in row]
            candidates = [j for j, h in enumerate(headers) if ('url' in h or 'address' in h) and ('sitemap' not in h and 'source' not in h)]
            if not candidates: continue
            col = candidates[0]
            urls = {norm(r[col]) for r in rows[head_idx+1:] if len(r)>col and r[col]}
            urls = {u for u in urls if u and not pathof(u).endswith('/site-map')}
            if urls: return urls, f'{file.name} / {name} / {headers[col]}'
    raise ValueError('Could not find expected staging URL column in Audit #3 workbook. Inspect its headers.')

def locations_from_html(file):
    soup = BeautifulSoup(file.read_text(encoding='utf-8',errors='replace'), 'html.parser')
    scopes = soup.select('div.sfa') or [soup]
    locations = set()
    for scope in scopes:
        for a in scope.select('a[href]'):
            p = pathof(norm(a.get('href'))) if norm(a.get('href')) else ''
            parts = [x for x in p.split('/') if x]
            if parts and re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*-[a-z]{2}', parts[0]) and parts[0] not in ('about-us', 'contact-us'): locations.add(parts[0])
    if not locations: raise ValueError('No location folders discovered from locations HTML. Check the file/DOM.')
    return sorted(locations)

def get_live(session, url, delay, retries):
    err=''; status=None
    for attempt in range(retries):
        time.sleep(delay + random.uniform(0, .35))
        try:
            r=session.get(url,timeout=(10,30),allow_redirects=True)
            status=r.status_code
            if r.status_code == 200 and '<html' in r.text[:4000].lower():
                return r.text, status, r.url, ''
            err=f'HTTP {status} or response not HTML'
            if status not in (403,408,429,500,502,503,504): break
            if status==429 and r.headers.get('Retry-After','').isdigit():
                time.sleep(min(60,int(r.headers['Retry-After'])))
        except requests.RequestException as e: err=str(e)[:240]
        if attempt+1<retries: time.sleep(min(30,2**attempt*2))
    return None,status,'',err

def extract(html, sitemap_url, locations):
    # Only actual /site-map endpoints may be used as sitemap sources.
    if not re.fullmatch(r'/(?:[^/]+/)?site-map', pathof(norm(sitemap_url))):
        raise ValueError(f'Not a sitemap endpoint: {sitemap_url}')
    soup=BeautifulSoup(html,'html.parser')
    # Extract ONLY links inside the verified sitemap list container.
    # A missing container is an extraction failure, not evidence of missing URLs.
    roots = soup.select('div.sf-sitemap__list')
    if not roots:
        raise ValueError('Required sitemap container missing: div.sf-sitemap__list')
    links = []
    query_links = []
    for a in soup.select('div.sf-sitemap__list a.sf-sitemap__link[href]'):
        # Preserve query-bearing URLs separately: norm() intentionally strips queries
        # for page-path reconciliation, which otherwise causes false duplicates.
        resolved = urljoin(sitemap_url, a['href'])
        if urlsplit(resolved).query:
            query_links.append((resolved, a.get_text(' ', strip=True)))
            continue
        u = norm(resolved)
        if not u:
            continue
        p = pathof(u)
        if p.endswith('/site-map') or p in ('/wp-admin', '/wp-login.php', '/locations'):
            continue
        if re.search(r'\.(?:pdf|jpg|jpeg|png|gif|webp|svg|zip|xml|css|js)$', p):
            continue
        links.append(u)
    return links, query_links

def allocation(u, locations):
    parts=[p for p in pathof(u).split('/') if p]
    return parts[0] if parts and parts[0] in locations else 'NATIONAL'

def write_report(output, datasets):
    book = Workbook()
    book.remove(book.active)
    for title, headers, records in datasets:
        sh = book.create_sheet(title)
        sh.append(list(headers))
        for record in records:
            sh.append(list(record))
        sh.freeze_panes = 'A2'
        sh.auto_filter.ref = sh.dimensions
        for cell in sh[1]:
            cell.fill = PatternFill('solid', fgColor='17365D')
            cell.font = Font(name='Calibri', bold=True, color='FFFFFF')
            cell.alignment = Alignment(wrap_text=True, vertical='center')
        sh.row_dimensions[1].height = 28
        for col in range(1, len(headers) + 1):
            sh.column_dimensions[get_column_letter(col)].width = (
                72 if col == 1 and title in
                ('Full URL Audit', 'Exceptions', 'Raw Entries') else 28
            )
        sh.sheet_view.showGridLines = False
    book.save(output)

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--directory',default='.',help='Folder containing inputs and output')
    parser.add_argument('--delay',type=float,default=2)
    parser.add_argument('--retries',type=int,default=3)
    args=parser.parse_args()
    folder=pathlib.Path(args.directory)
    locations_file=folder/'vz-locations-rendered-html.html'
    if not locations_file.exists(): raise FileNotFoundError(locations_file)
    matches=sorted(folder.glob('*audit3*.xlsx'))
    if not matches: raise FileNotFoundError('Place the Audit #3 XLSX workbook in the script folder (filename containing audit3).')
    expected, baseline=expected_from_audit3(matches[0])
    locations=locations_from_html(locations_file)
    targets=[('NATIONAL',BASE+'/site-map')]+[(loc,f'{BASE}/{loc}/site-map') for loc in locations]
    print(f'Expected URLs: {len(expected)} | Local sitemaps: {len(locations)} | Baseline: {baseline}',flush=True)
    session=requests.Session();session.headers.update(HEADERS)
    inventory=[]; raw=[]; query_entries=[]; failed=set(); by_url=collections.defaultdict(list)
    # Avoid processing a sitemap more than once if location discovery repeats it.
    targets=list(dict.fromkeys(targets))
    for index,(owner,url) in enumerate(targets,1):
        html,status,final,error=get_live(session,url,args.delay,args.retries)
        entries=[]; skipped_queries=[]
        if html:
            try:
                # Redirects to ordinary pages (e.g. about-us/contact-us) are failures,
                # not successful sitemap retrievals.
                if pathof(norm(final)) != pathof(norm(url)):
                    raise ValueError(f'Sitemap redirected to nonmatching URL: {final}')
                entries,skipped_queries=extract(html,url,locations)
            except Exception as e: error=str(e); html=None
        if not html: failed.add(owner)
        inventory.append([owner,url,status or '',final,'OK' if html else 'RETRIEVAL_FAILED',len(entries),len(set(entries)),error])
        for query_url, link_text in skipped_queries:
            query_entries.append([query_url,owner,url,link_text,'QUERY_STRING_EXCLUDED'])
        for entry in entries:
            raw.append([entry,owner,url])
            by_url[entry].append(owner)
        print(f'[{index}/{len(targets)}] {owner}: {"OK" if html else "FAILED"} ({len(entries)} links)',flush=True)
    full=[]; exceptions=[]
    for u in sorted(expected):
        owner=allocation(u,locations); actual=by_url.get(u,[])
        own_count=actual.count(owner); other=[x for x in actual if x!=owner]
        if owner in failed and not actual: status='UNVERIFIABLE_SITEMAP_FAILED'
        elif not actual: status='MISSING_FROM_HTML_SITEMAPS'
        elif not own_count: status='WRONG_ALLOCATION'
        elif other: status='CROSS_SITEMAP_DUPLICATE'
        elif own_count>1: status='DUPLICATE_IN_SITEMAP'
        else: status='PASS'
        row=[u,owner,', '.join(sorted(set(actual))),len(actual),status]
        full.append(row)
        if status!='PASS': exceptions.append(row)
    for u,owners in sorted(by_url.items()):
        if u not in expected:
            owner=allocation(u,locations)
            status='UNEXPECTED_HTML_SITEMAP_URL'
            exceptions.append([u,owner,', '.join(sorted(set(owners))),len(owners),status])
    for owner,url in targets:
        if owner in failed: exceptions.append([url,owner,'',0,'RETRIEVAL_FAILED'])
    counts=collections.Counter(r[4] for r in full)
    counts.update(collections.Counter(r[4] for r in exceptions if r[4] in ('UNEXPECTED_HTML_SITEMAP_URL','RETRIEVAL_FAILED')))
    summary=[['Run UTC',datetime.datetime.now(datetime.timezone.utc).isoformat()],['Expected baseline',baseline],['Expected URLs',len(expected)],['Locations',len(locations)],['Sitemaps attempted',len(targets)],['Sitemaps failed',len(failed)],['Unique sitemap URLs',len(by_url)],['Query-string links excluded',len(query_entries)],['Total exceptions',len(exceptions)]]
    summary += [[key,value] for key,value in sorted(counts.items())]
    headers=['Normalized URL','Expected sitemap allocation','Found in sitemap allocation(s)','Occurrences','Status']
    datasets=[('Summary',['Metric','Value'],summary),('Exceptions',headers,exceptions),('Full URL Audit',headers,full),('Sitemap Inventory',['Allocation','Requested sitemap','HTTP','Final URL','Status','Links','Unique links','Error'],inventory),('Raw Entries',['Normalized URL','Sitemap allocation','Source sitemap'],raw),('Excluded Query URLs',['Original URL','Sitemap allocation','Source sitemap','Anchor text','Reason'],query_entries),('Location Folders',['Location folder','HTML sitemap URL'],[[x,f'{BASE}/{x}/site-map'] for x in locations])]
    output=folder/OUT;write_report(output,datasets)
    print(f'\nReport: {output}\nExceptions: {len(exceptions)} | Failed sitemaps: {len(failed)}',flush=True)

if __name__=='__main__':main()
