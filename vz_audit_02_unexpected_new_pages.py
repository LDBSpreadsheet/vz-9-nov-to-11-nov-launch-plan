# VZ Migration — Audit #2: Unexpected New Pages
# Jupyter: %pip install requests beautifulsoup4 lxml pandas openpyxl
from pathlib import Path
from urllib.parse import urljoin, urlsplit, urlunsplit, unquote
from collections import deque
import time, re, random, xml.etree.ElementTree as ET
import requests
import pandas as pd
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

STAGING = 'https://varsityzonedev.wpenginepowered.com'
PRODUCTION = 'https://www.varsityzone.com'
BASE_DIR = Path.cwd()
LOCATIONS_FILE = BASE_DIR / 'vz-locations-rendered-html.html'
PRODUCTION_XML_FALLBACK = BASE_DIR / 'scorpion_xml_sitemap.xml'
OUTPUT_FILE = BASE_DIR / 'vz-audit-02-unexpected-new-pages.xlsx'
PRODUCTION_XML_CANDIDATES = [PRODUCTION + '/sitemap.xml', PRODUCTION + '/sitemap_index.xml']
NATIONAL_SITEMAP = STAGING + '/site-map'
REQUEST_TIMEOUT = (12, 35)
MAX_ATTEMPTS = 3
SLEEP_BETWEEN_REQUESTS = 0.7
RETRY_BACKOFF_SECONDS = 1.5
# Set to True only if the HTML sitemap content selector is known to include the page list.
# Otherwise inspect 'Sitemap Links' output and adjust CONTENT_SELECTORS.
CONTENT_SELECTORS = ['main', '[role="main"]', '.entry-content', '.page-content', 'article']
# Optional manual additions, if some location is absent from the saved HTML.
ADDITIONAL_SITEMAPS = []
EXCLUDED_EXTENSIONS = ('.jpg','.jpeg','.png','.gif','.webp','.svg','.pdf','.zip','.css','.js','.ico','.mp4','.webm','.xml','.txt')


def normalize_url(url, target_domain=PRODUCTION, trailing_slash=True):
    """Normalize paths for comparison; production-style URLs end in /.

    Ignores query strings/fragments and treats /page and /page/ identically.
    Set trailing_slash=False for staging URL display only.
    """
    p = urlsplit(str(url).strip())
    path = re.sub(r'/+', '/', unquote(p.path or '/'))
    if path != '/':
        path = path.rstrip('/')
        if trailing_slash:
            path += '/'
    return target_domain.rstrip('/') + path


def eligible_page_url(url):
    p = urlsplit(url)
    if p.scheme not in ('http','https') or p.netloc.lower() != urlsplit(STAGING).netloc.lower():
        return False
    path = p.path.lower().rstrip('/')
    if path.endswith(EXCLUDED_EXTENSIONS): return False
    if any(path.startswith(prefix) for prefix in ('/wp-admin','/wp-json','/wp-content','/wp-includes')): return False
    if path.endswith('/feed') or path == '/feed': return False
    return True


session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (compatible; VZ-Migration-Audit/2.0)', 'Accept': 'text/html,application/xhtml+xml,application/xml,text/xml;q=0.9,*/*;q=0.8'})
# Retries are controlled explicitly so every failed attempt is visible in diagnostics.


def fetch(url):
    errors = []
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            if response.status_code == 200:
                return response, attempt, ''
            errors.append(f'attempt {attempt}: HTTP {response.status_code}')
            if response.status_code not in (403, 408, 429, 500, 502, 503, 504):
                break
        except requests.RequestException as exc:
            errors.append(f'attempt {attempt}: {type(exc).__name__}: {str(exc)[:140]}')
        if attempt < MAX_ATTEMPTS:
            time.sleep(RETRY_BACKOFF_SECONDS * attempt)
    return None, len(errors), '; '.join(errors)


def read_local_html():
    if not LOCATIONS_FILE.is_file():
        raise FileNotFoundError(f'Missing required local file: {LOCATIONS_FILE}')
    return BeautifulSoup(LOCATIONS_FILE.read_text(encoding='utf-8', errors='replace'), 'lxml')


def discover_location_sitemaps():
    soup = read_local_html()
    container = soup.select_one('div.sfa-locations-results')
    if container is None:
        raise RuntimeError('Could not find div.sfa-locations-results in saved HTML')
    cards = container.select('article.location-card')
    if not cards:
        raise RuntimeError('No article.location-card elements in saved HTML')
    rows, issues = [], []
    for index, card in enumerate(cards, 1):
        heading = card.select_one('h2')
        name = heading.get_text(' ', strip=True) if heading else f'Card {index}'
        matches = [a for a in card.select('a[href]') if a.get_text(' ', strip=True).casefold() == 'website']
        if len(matches) != 1:
            issues.append({'Location':name,'Issue':f'Expected 1 Website link, found {len(matches)}'})
            continue
        website = urljoin(STAGING, matches[0]['href'])
        if not eligible_page_url(website):
            issues.append({'Location':name,'Issue':f'Unexpected Website URL: {website}'})
            continue
        site_map = normalize_url(website, STAGING).rstrip('/') + '/site-map'
        rows.append({'Location':name,'Website URL':website,'HTML Sitemap URL':site_map,'Source':'Location card'})
    if issues:
        raise RuntimeError(f'Location extraction incomplete: {issues[:10]}')
    result = pd.DataFrame(rows).drop_duplicates('HTML Sitemap URL')
    if len(result) != len(cards):
        raise RuntimeError(f'{len(cards)} location cards but only {len(result)} unique sitemaps. Check duplicate Website links.')
    national = pd.DataFrame([{'Location':'National','Website URL':STAGING,'HTML Sitemap URL':NATIONAL_SITEMAP,'Source':'National seed'}])
    extras = pd.DataFrame([{'Location':'Manual','Website URL':'','HTML Sitemap URL':u,'Source':'Manual seed'} for u in ADDITIONAL_SITEMAPS])
    return pd.concat([national, result, extras], ignore_index=True).drop_duplicates('HTML Sitemap URL'), len(cards)


def xml_locs(xml_bytes):
    root = ET.fromstring(xml_bytes)
    return [(node.text or '').strip() for node in root.iter() if node.tag.rsplit('}',1)[-1] == 'loc' and (node.text or '').strip()], root.tag.rsplit('}',1)[-1].lower()


def load_production_xml():
    """Try live XML + sitemap indexes; if unavailable, use local XML snapshot."""
    errors, urls, sources = [], set(), []
    queue = deque(PRODUCTION_XML_CANDIDATES)
    seen = set()
    while queue and len(seen) < 80:
        url = queue.popleft()
        if url in seen: continue
        seen.add(url)
        response, tries, err = fetch(url)
        if response is None:
            errors.append({'URL':url,'Error':err})
            continue
        try:
            locs, root_type = xml_locs(response.content)
            if root_type == 'sitemapindex':
                queue.extend(u for u in locs if urlsplit(u).netloc == urlsplit(PRODUCTION).netloc)
            elif root_type == 'urlset':
                urls.update(normalize_url(u) for u in locs if urlsplit(u).netloc == urlsplit(PRODUCTION).netloc)
                sources.append(url)
            else:
                errors.append({'URL':url,'Error':f'Unexpected XML root: {root_type}'})
        except ET.ParseError as exc:
            errors.append({'URL':url,'Error':f'Invalid XML: {exc}'})
        time.sleep(SLEEP_BETWEEN_REQUESTS)
    if urls:
        return urls, 'Live Scorpion XML', sources, errors
    if not PRODUCTION_XML_FALLBACK.is_file():
        raise RuntimeError('Live Scorpion XML unavailable and local fallback file missing: ' + str(PRODUCTION_XML_FALLBACK))
    data = PRODUCTION_XML_FALLBACK.read_bytes()
    locs, root_type = xml_locs(data)
    if root_type != 'urlset':
        raise RuntimeError('Local XML fallback is not a complete <urlset>. Provide a full consolidated XML sitemap, not a sitemap index.')
    urls = {normalize_url(u) for u in locs if urlsplit(u).netloc == urlsplit(PRODUCTION).netloc}
    if not urls:
        raise RuntimeError('Local XML fallback contains no valid production URLs')
    return urls, 'Local Scorpion XML fallback', [str(PRODUCTION_XML_FALLBACK)], errors


def extract_html_sitemap_links(html, base_url):
    soup = BeautifulSoup(html, 'lxml')
    # Avoid sitewide header/footer navigation: prefer content region(s).
    main = soup.select_one('main') or soup.select_one('[role="main"]')
    if main is None:
        raise ValueError('No <main> or [role=main] element; cannot safely distinguish sitemap from sitewide navigation')
    for element in main.select('header, footer, nav, script, style'):
        element.decompose()
    links = []
    for anchor in main.select('a[href]'):
        absolute = urljoin(base_url, anchor.get('href',''))
        if not eligible_page_url(absolute): continue
        if urlsplit(absolute).path.rstrip('/').endswith('/site-map'):
            # Sitemap pages are discovery infrastructure, not ordinary page inventory.
            continue
        links.append({'Staging URL':normalize_url(absolute, STAGING, trailing_slash=False),
                      'Production Equivalent':normalize_url(absolute, PRODUCTION),
                      'Anchor Text':anchor.get_text(' ', strip=True),
                      'Source Sitemap':base_url})
    return links


def run_audit():
    seeds, cards = discover_location_sitemaps()
    print(f'Found {cards} location cards; {len(seeds)} HTML sitemap URLs including national.')
    production, xml_source, xml_sources, xml_errors = load_production_xml()
    print(f'Production reference: {len(production)} URLs ({xml_source}).')
    sitemap_rows, link_rows = [], []
    for idx, seed in seeds.iterrows():
        url = seed['HTML Sitemap URL']
        print(f'[{idx+1}/{len(seeds)}] {url}')
        response, tries, error = fetch(url)
        row = {**seed.to_dict(), 'Attempts':tries, 'HTTP Status':response.status_code if response is not None else '',
               'Final URL':response.url if response is not None else '', 'Extracted Links':0, 'Status':'FAILED', 'Error':error}
        if response is not None:
            try:
                if 'html' not in response.headers.get('Content-Type','').lower() and not response.text.lstrip().startswith(('<!DOCTYPE','<html')):
                    raise ValueError('Response is not HTML')
                links = extract_html_sitemap_links(response.text, response.url)
                if not links:
                    raise ValueError('No page links found in main content; inspect sitemap template')
                link_rows.extend(links)
                row.update({'Extracted Links':len(links), 'Status':'OK', 'Error':''})
            except Exception as exc:
                row['Error'] = f'{type(exc).__name__}: {exc}'
        sitemap_rows.append(row)
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    sitemap_df = pd.DataFrame(sitemap_rows)
    all_links = pd.DataFrame(link_rows, columns=['Staging URL','Production Equivalent','Anchor Text','Source Sitemap'])
    if not all_links.empty:
        all_links = all_links.drop_duplicates(['Staging URL','Source Sitemap'])
        pages = all_links.drop_duplicates('Production Equivalent').copy()
        pages['In Production XML'] = pages['Production Equivalent'].isin(production)
        exceptions = pages[~pages['In Production XML']].copy()
        exceptions.insert(0, 'Classification', 'Not in production XML — investigate')
    else:
        pages = pd.DataFrame(columns=['Staging URL','Production Equivalent','Anchor Text','Source Sitemap','In Production XML'])
        exceptions = pd.DataFrame(columns=['Classification',*pages.columns])
    failures = sitemap_df[sitemap_df['Status'] != 'OK']
    complete = failures.empty and not pages.empty
    summary = pd.DataFrame([
        ('Audit','02 — Unexpected New Pages'),('Run timestamp',pd.Timestamp.now().isoformat(timespec='seconds')),
        ('Audit status','COMPLETE' if complete else 'INCOMPLETE — sitemap extraction failures'),
        ('Location cards in local HTML',cards),('HTML sitemap URLs expected',len(seeds)),
        ('HTML sitemaps successful',int((sitemap_df['Status']=='OK').sum())),
        ('HTML sitemaps failed',len(failures)),('Unique staging page URLs extracted',len(pages)),
        ('Potential unexpected pages',len(exceptions)),('Production XML URLs',len(production)),
        ('Production XML source',xml_source),('Production XML files loaded','; '.join(xml_sources)),
        ('Comparison rule','Production equivalents and production XML normalized with trailing slash; query/fragment ignored'),
        ('Interpretation','Absent from XML does not prove the URL is absent from live production')
    ],columns=['Metric','Value'])
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for name, df in [('Summary',summary),('Unexpected New Pages',exceptions),('All Staging URLs',pages),
                         ('Sitemap Discovery',sitemap_df),('All Production URLs',pd.DataFrame({'Production URL':sorted(production)})),
                         ('Sitemap Links',all_links),('Production XML Errors',pd.DataFrame(xml_errors,columns=['URL','Error']))]:
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            from openpyxl.styles import Font, PatternFill, Alignment
            for cell in ws[1]:
                cell.font = Font(bold=True,color='FFFFFF')
                cell.fill = PatternFill('solid',fgColor='17365D')
            for col in ws.columns:
                col_letter = col[0].column_letter
                max_len = max((len(str(c.value)) for c in list(col)[:1000] if c.value is not None),default=10)
                ws.column_dimensions[col_letter].width = min(max(max_len+2,16),75)
    print('\nAUDIT STATUS:', 'COMPLETE' if complete else 'INCOMPLETE')
    print('Potential unexpected pages:',len(exceptions))
    print('Output:',OUTPUT_FILE)
    if not complete:
        print('WARNING: Failed sitemap(s). Do not treat exception list as comprehensive.')
    return {'summary':summary,'exceptions':exceptions,'sitemaps':sitemap_df,'pages':pages,'output':OUTPUT_FILE}


if __name__ == '__main__':
    audit = run_audit()
