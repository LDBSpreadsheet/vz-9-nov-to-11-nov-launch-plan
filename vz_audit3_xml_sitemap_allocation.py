# VZ MIGRATION — AUDIT #3: XML SITEMAP COMPLETENESS & ALLOCATION
# Run in Jupyter with: %run vz_audit3_xml_sitemap_allocation_html_xml.py
# Requirements: pip install pandas openpyxl beautifulsoup4 lxml
# Classification: first URL path segment in known location folders => LOCAL; otherwise NATIONAL.
from pathlib import Path
from urllib.parse import urlsplit, unquote
from collections import defaultdict, Counter
from datetime import date
import re, zipfile, io, xml.etree.ElementTree as ET
import pandas as pd
from bs4 import BeautifulSoup

BASE = Path.cwd()
STAGING_DOMAIN = 'https://varsityzonedev.wpenginepowered.com'
PRODUCTION_XML = BASE / 'scorpion_xml_sitemap.xml'
LOCATIONS_HTML = BASE / 'vz-locations-rendered-html.html'
SITEMAP_DIRECTORY = BASE / 'WordPress'  # directory with all staging sitemap XML files
SITEMAP_ZIP = BASE / 'WordPress.zip'      # set to WordPress(1).zip if that is your current archive
OUTPUT = BASE / f'vz-staging-audit3-xml-sitemap-allocation-{date.today():%d-%b-%Y}.xlsx'
# Optional: set this to an Audit #1 URL-list file (CSV or XLSX) to use it instead of production XML.
AUDIT1_URL_LIST = None
# If Audit #1 deliberately excluded some production URLs, supply that exact URL list above.
# Otherwise the original production sitemap is used as the complete expected inventory.


def read_file(path):
    return Path(path).read_bytes()

def norm_path(value):
    value = str(value).strip()
    if not value: return '/'
    p = urlsplit(value)
    path = unquote(p.path if (p.scheme or p.netloc) else value.split('?',1)[0].split('#',1)[0])
    path = '/' + path.lstrip('/')
    path = re.sub(r'/+', '/', path)
    return path.rstrip('/') or '/'

def norm_url(value):
    return STAGING_DOMAIN + ('' if norm_path(value) == '/' else norm_path(value))

def extract_xml_locs(data):
    root = ET.fromstring(data)
    root_tag = root.tag.rsplit('}',1)[-1].lower()
    if root_tag not in ('urlset','sitemapindex'):
        raise ValueError(f'Unexpected XML root: {root_tag}')
    return root_tag, [el.text.strip() for el in root.iter() if el.tag.rsplit('}',1)[-1].lower() == 'loc' and el.text and el.text.strip()]

def load_expected():
    if AUDIT1_URL_LIST:
        p = Path(AUDIT1_URL_LIST)
        df = pd.read_excel(p) if p.suffix.lower() in ('.xlsx','.xls') else pd.read_csv(p)
        possible = [c for c in df.columns if any(x in str(c).lower() for x in ('address','url'))]
        if not possible: raise ValueError(f'No URL column in {p.name}; columns: {list(df.columns)}')
        urls = df[possible[0]].dropna().astype(str).tolist()
    else:
        if not PRODUCTION_XML.exists(): raise FileNotFoundError(f'Missing {PRODUCTION_XML.name}; supply production XML or AUDIT1_URL_LIST')
        kind, urls = extract_xml_locs(read_file(PRODUCTION_XML))
        if kind != 'urlset': raise ValueError('Production input must be a URL sitemap, not an index. Use Audit #1 URL list if production has multiple sitemap files.')
    return sorted({norm_url(u) for u in urls if str(u).strip().startswith(('http://','https://','/'))})

def load_locations():
    if not LOCATIONS_HTML.exists(): raise FileNotFoundError(f'Missing {LOCATIONS_HTML.name}')
    soup = BeautifulSoup(read_file(LOCATIONS_HTML), 'html.parser')
    # Audit #2: location cards use an SFA-prefixed wrapper class.
    wrappers = soup.find_all('div', class_=lambda value: value and any(
        c.startswith('sfa') for c in (value if isinstance(value, list) else str(value).split())
    ))
    if not wrappers:
        raise ValueError('No SFA location container found in saved locations HTML.')
    candidates = set()
    for wrapper in wrappers:
        for a in wrapper.select('a[href]'):
            href = a['href']
            parsed = urlsplit(href)
            if parsed.netloc and parsed.netloc.lower() not in (
                'varsityzonedev.wpenginepowered.com', 'www.varsityzone.com', 'varsityzone.com'
            ):
                continue
            segments = [part for part in norm_path(href).split('/') if part]
            if len(segments) == 1 and re.fullmatch(r'[a-z0-9-]+-[a-z]{2}', segments[0].lower()):
                candidates.add(segments[0].lower())
    if not candidates:
        raise ValueError('No location homepage URLs extracted from SFA location containers.')
    print(f'Location folders discovered: {len(candidates)}')
    return sorted(candidates)

def sitemap_sources():
    """Read XML and saved HTML sitemap files from a directory or ZIP.

    Normalizes the national HTML filename to the index's national-sitemap.xml
    identifier. If both wp_national-sitemap.html and national-sitemap.html exist,
    uses national-sitemap.html once, not twice.
    """
    raw = []
    if SITEMAP_DIRECTORY.is_dir():
        raw = [(str(p.relative_to(SITEMAP_DIRECTORY)), p.read_bytes())
               for p in SITEMAP_DIRECTORY.rglob('*')
               if p.is_file() and p.suffix.lower() in ('.xml', '.html', '.htm')]
        print(f'Reading staging sitemap directory: {SITEMAP_DIRECTORY}')
    elif SITEMAP_ZIP.exists():
        with zipfile.ZipFile(SITEMAP_ZIP) as z:
            raw = [(n, z.read(n)) for n in z.namelist()
                   if not n.endswith('/') and Path(n).suffix.lower() in ('.xml', '.html', '.htm')]
        print(f'Reading staging sitemap ZIP: {SITEMAP_ZIP.name}')
    else:
        raise FileNotFoundError('Supply WordPress/ folder or configure SITEMAP_ZIP')

    sources = {}
    for original_name, data in sorted(raw, key=lambda x: x[0]):
        basename = Path(original_name).name
        low = basename.lower()
        # Ignore unrelated HTML files; accept the known national sitemap aliases.
        if low.endswith(('.html', '.htm')) and low not in (
            'national-sitemap.html', 'wp_national-sitemap.html',
            'national-sitemap.htm', 'wp_national-sitemap.htm'
        ):
            continue
        canonical = ('national-sitemap.xml' if low in (
            'national-sitemap.html', 'wp_national-sitemap.html',
            'national-sitemap.htm', 'wp_national-sitemap.htm'
        ) else basename)
        if canonical in sources:
            existing_name, existing_data = sources[canonical]
            if data != existing_data:
                # Prefer native XML, then the canonical HTML filename.
                prefer_new = (basename.lower() == 'national-sitemap.xml' or
                              (basename.lower().startswith('national-sitemap.') and
                               existing_name.lower().startswith('wp_national-sitemap.')))
                if prefer_new:
                    print(f'WARNING: conflicting copies for {canonical}; using {basename}')
                    sources[canonical] = (basename, data)
                else:
                    print(f'WARNING: conflicting copies for {canonical}; keeping {existing_name}')
            else:
                print(f'Ignoring duplicate copy: {basename} (same as {existing_name})')
            continue
        sources[canonical] = (basename, data)
    return sources


def extract_sitemap(data, filename):
    """Extract sitemap URLs from native XML or a browser-saved Rank Math HTML table."""
    if filename.lower().endswith(('.html', '.htm')):
        soup = BeautifulSoup(data, 'html.parser')
        table = soup.select_one('table#sitemap')
        if table is None:
            raise ValueError('Saved HTML has no table#sitemap; cannot safely extract sitemap URLs')
        urls = []
        for tr in table.select('tbody tr'):
            td = tr.find('td')
            if td is None:
                continue
            a = td.find('a', href=True)
            if a and a['href'].startswith(('http://', 'https://')):
                urls.append(a['href'].strip())
        return 'urlset', urls
    return extract_xml_locs(data)

def main():
    expected = load_expected()
    locations = load_locations()
    locset = set(locations)
    sources = sitemap_sources()
    if len(sources) != len(set(sources)):
        raise ValueError('Duplicate sitemap basenames; preserve unique filenames')
    index_names = []
    rows = []
    parse_errors = []
    for filename, (original_name, data) in sorted(sources.items()):
        try: kind, urls = extract_sitemap(data, original_name)
        except Exception as e:
            parse_errors.append({'File':filename,'Issue':str(e)})
            continue
        if kind == 'sitemapindex':
            index_names += [Path(urlsplit(u).path).name for u in urls]
            continue
        for u in urls:
            rows.append({'Sitemap':filename,'Raw URL':u,'Normalized URL':norm_url(u),
                         'Domain issue':urlsplit(u).netloc.lower() != urlsplit(STAGING_DOMAIN).netloc.lower(),
                         'Trailing slash issue':urlsplit(u).path.endswith('/') and urlsplit(u).path != '/'})
    index_names = list(dict.fromkeys(index_names))
    expected_sitemaps = set(index_names)
    missing_sitemap_files = sorted(expected_sitemaps - set(sources))
    if missing_sitemap_files:
        print('WARNING: sitemap files referenced by index but absent from supplied XML files:',missing_sitemap_files)
    if parse_errors:
        print('WARNING: sitemap parsing errors:',parse_errors)
    if not rows: raise ValueError('No URL entries found in supplied staging XML files')
    entries = pd.DataFrame(rows)
    by_url = defaultdict(list)
    for r in rows: by_url[r['Normalized URL']].append(r)

    def classify(url):
        path = norm_path(url)
        first = path.strip('/').split('/')[0].lower() if path != '/' else ''
        if first in locset: return 'Local',first, f'{first}-sitemap.xml'
        return 'National','', 'national-sitemap.xml'

    audit = []
    for u in expected:
        kind, slug, target = classify(u)
        occurrences = by_url.get(u,[])
        actual = sorted({r['Sitemap'] for r in occurrences})
        if not occurrences: status = 'MISSING'
        elif target not in actual: status = 'WRONG SITEMAP'
        elif len(occurrences)>1: status = 'DUPLICATE'
        else: status = 'PASS'
        audit.append({'Expected staging URL':u,'Classification':kind,'Location folder':slug,
                      'Expected sitemap':target,'Actual sitemap(s)':'; '.join(actual),
                      'Occurrences':len(occurrences),'Status':status})
    audit_df = pd.DataFrame(audit)
    exception_rows = []
    for r in audit:
        if r['Status'] != 'PASS': exception_rows.append({'Issue':r['Status'],'URL':r['Expected staging URL'],
            'Expected sitemap':r['Expected sitemap'],'Actual sitemap(s)':r['Actual sitemap(s)'], 'Details':''})
    expected_set = set(expected)
    for u, occ in by_url.items():
        kind, slug, target = classify(u)
        actual = sorted({r['Sitemap'] for r in occ})
        if u not in expected_set:
            exception_rows.append({'Issue':'UNEXPECTED','URL':u,'Expected sitemap':target,
                'Actual sitemap(s)':'; '.join(actual),'Details':'Present in staging sitemap but not expected production URL inventory'})
        if len(occ)>1 and u not in expected_set:
            exception_rows.append({'Issue':'DUPLICATE','URL':u,'Expected sitemap':target,
                'Actual sitemap(s)':'; '.join(actual),'Details':f'{len(occ)} entries'})
        if target not in actual and u not in expected_set:
            exception_rows.append({'Issue':'WRONG SITEMAP','URL':u,'Expected sitemap':target,
                'Actual sitemap(s)':'; '.join(actual),'Details':'Unexpected URL also appears misallocated'})
    for r in rows:
        problems = []
        if r['Domain issue']: problems.append('WRONG DOMAIN')
        if r['Trailing slash issue']: problems.append('TRAILING SLASH')
        for problem in problems:
            exception_rows.append({'Issue':problem,'URL':r['Raw URL'],'Expected sitemap':'',
                'Actual sitemap(s)':r['Sitemap'],'Details':'Raw XML value differs from expected URL convention'})
    for name in missing_sitemap_files:
        exception_rows.append({'Issue':'SITEMAP FILE MISSING','URL':'','Expected sitemap':name,
            'Actual sitemap(s)':'','Details':'Referenced in sitemap index but XML file not supplied'})
    for err in parse_errors:
        exception_rows.append({'Issue':'PARSE ERROR','URL':'','Expected sitemap':err['File'],
            'Actual sitemap(s)':'','Details':err['Issue']})
    for name in sorted(set(sources)-set(index_names)):
        if name not in ('wp_sitemap_index.xml','sitemap_index.xml'):
            exception_rows.append({'Issue':'UNLISTED SITEMAP FILE','URL':'','Expected sitemap':name,
                'Actual sitemap(s)':'','Details':'File supplied but not referenced in sitemap index'})
    exceptions = pd.DataFrame(exception_rows,columns=['Issue','URL','Expected sitemap','Actual sitemap(s)','Details'])
    inventory = pd.DataFrame([{'Sitemap file':name,'In index':name in expected_sitemaps,
         'File supplied':name in sources,'Parsed URL entries':sum(r['Sitemap']==name for r in rows),
         'Source filename':sources[name][0] if name in sources else ''}
         for name in sorted(expected_sitemaps | set(sources)) if name not in ('wp_sitemap_index.xml','sitemap_index.xml')])
    counts = audit_df['Status'].value_counts()
    summary = pd.DataFrame([{'Metric':'Expected URLs','Value':len(expected)},
        {'Metric':'Valid local folders','Value':len(locations)},
        {'Metric':'Sitemap URL entries','Value':len(rows)},
        {'Metric':'Unique sitemap URLs','Value':len(by_url)},
        {'Metric':'Referenced child sitemaps','Value':len(expected_sitemaps)},
        {'Metric':'Missing referenced sitemap files','Value':len(missing_sitemap_files)},
        {'Metric':'Sitemap parse errors','Value':len(parse_errors)}] +
        [{'Metric':f'Expected URLs: {k}','Value':int(v)} for k,v in sorted(counts.items())] +
        [{'Metric':f'Exceptions: {k}','Value':int(v)} for k,v in sorted(exceptions['Issue'].value_counts().items())])
    with pd.ExcelWriter(OUTPUT,engine='openpyxl') as writer:
        for sheet, df in [('Summary',summary),('Exceptions',exceptions),('Full URL Audit',audit_df),
                          ('Sitemap Inventory',inventory),('Raw Sitemap Entries',entries),
                          ('Location Folders',pd.DataFrame({'Location folder':locations}))]:
            df.to_excel(writer,sheet_name=sheet,index=False)
            ws = writer.sheets[sheet]
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            from openpyxl.styles import PatternFill, Font
            for cell in ws[1]:
                cell.fill = PatternFill('solid',fgColor='17365D')
                cell.font = Font(bold=True,color='FFFFFF')
            for col in ws.columns:
                letter=col[0].column_letter
                max_len=max((len(str(c.value or '')) for c in list(col)[:1500]),default=10)
                ws.column_dimensions[letter].width=min(max(max_len+2,14),75)
    print(f'Completed Audit #3: {OUTPUT}')
    print(summary.to_string(index=False))
    print(f'\nExceptions: {len(exceptions)}')
    return audit_df,exceptions,inventory

if __name__ == '__main__':
    AUDIT, EXCEPTIONS, INVENTORY = main()
