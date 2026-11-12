"""Varsity Zone Audit 3: SSOT-driven local XML sitemap allocation audit.
Place this .py, Audit 1 .xlsx, and all saved sitemap files in one folder.
Run in Jupyter: %run vz_audit3_ssot_xml_sitemap_reaudit_12_nov_2026.py
Requires: pip install openpyxl
"""
from pathlib import Path
from collections import defaultdict, Counter
from urllib.parse import urlsplit, unquote
from xml.etree import ElementTree as ET
import re, csv
from html.parser import HTMLParser
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()
INPUT_NAME = 'vz-wednesday-reaudit1-screaming-frog-comparison-11-nov-2026.xlsx'
OUTPUT_NAME = 'vz-staging-audit3-ssot-xml-sitemap-reaudit-12-nov-2026.xlsx'
# Optional: provide a location slug -> sitemap stem override for nonstandard location naming.
LOCATION_SITEMAP_OVERRIDES = {}  # e.g. {'example-tx': 'example-tx-sitemap'}


def path_of(url):
    path = unquote(urlsplit(str(url or '')).path)
    return '/' + path.strip('/').lower() if path.strip('/') else '/'


def normalized_url(url):
    p = urlsplit(str(url or '').strip())
    return (p.netloc.lower(), path_of(url))


def status_of(value):
    try: return int(float(value))
    except (TypeError, ValueError): return None


def local_slug(path):
    parts = path.strip('/').split('/')
    return parts[0] if parts else ''


def is_round_rock(path):
    return path == '/round-rock-tx' or path.startswith('/round-rock-tx/')


def is_blog_category(path):
    return bool(re.search(r'/blog/categories(?:/[^/]+)?/?$', path))


def is_blog_article_exempt(path):
    return '/blog/' in path and not is_blog_category(path)


def is_space_suffix_exempt(production_url):
    # Inspect the ORIGINAL production URL: unquote preserves trailing spaces.
    raw = unquote(urlsplit(str(production_url)).path)
    final = raw.rstrip('/').split('/')[-1]
    return final.endswith(' ') or final == '-' or raw.rstrip('/').endswith('/-')


class SitemapHTMLParser(HTMLParser):
    """Read URL anchors only from the Rank Math sitemap table."""
    def __init__(self):
        super().__init__()
        self.in_table = False
        self.table_depth = 0
        self.urls = []
    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == 'table' and attrs.get('id') == 'sitemap':
            self.in_table = True
            self.table_depth = 1
        elif tag == 'table' and self.in_table:
            self.table_depth += 1
        elif tag == 'a' and self.in_table:
            href = attrs.get('href', '')
            if href.startswith(('http://', 'https://')):
                self.urls.append(href)
    def handle_endtag(self, tag):
        if tag == 'table' and self.in_table:
            self.table_depth -= 1
            if not self.table_depth:
                self.in_table = False


def parse_sitemap(file):
    raw = file.read_bytes()
    # Browser 'Save As' can transform Rank Math XML into an HTML table.
    # Parse both formats, restricting HTML extraction to the sitemap table.
    try:
        root = ET.fromstring(raw)
        tag = root.tag.rsplit('}', 1)[-1].lower()
        if tag in ('urlset', 'sitemapindex'):
            items = []
            for node in root:
                if node.tag.rsplit('}', 1)[-1].lower() != ('url' if tag == 'urlset' else 'sitemap'):
                    continue
                for child in node:
                    if child.tag.rsplit('}', 1)[-1].lower() == 'loc' and child.text:
                        items.append(child.text.strip())
            return tag, items
    except ET.ParseError:
        pass
    parser = SitemapHTMLParser()
    parser.feed(raw.decode('utf-8-sig', errors='replace'))
    if not parser.urls:
        raise ValueError('No XML <loc> entries or Rank Math HTML sitemap table URLs found')
    # The sitemap index HTML table also has id=sitemap; its links are sitemap files.
    is_index = file.stem.lower() == 'sitemap_index'
    return ('sitemapindex' if is_index else 'urlset'), parser.urls


def find_sitemap_files():
    found = {}
    for f in BASE.iterdir():
        if not f.is_file() or f.suffix.lower() not in ('.xml', '.html', '.htm', '.txt'):
            continue
        stem = f.stem.lower()
        if stem == 'sitemap_index' or stem.endswith('-sitemap'):
            if stem in found: raise ValueError(f'Duplicate sitemap filenames/stems: {found[stem].name} and {f.name}')
            found[stem] = f
    return found


def sheet(wb, name, headers, rows):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for row in rows: ws.append([str(x) if isinstance(x, (list, tuple, set)) else x for x in row])
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 27
    for c in ws[1]:
        c.fill = PatternFill('solid', fgColor='16324F')
        c.font = Font(color='FFFFFF', bold=True)
        c.alignment = Alignment(wrap_text=True, vertical='center')
    for i, h in enumerate(headers, 1):
        width = 75 if 'URL' in h or 'Sitemap' in h or 'Action' in h else 34
        ws.column_dimensions[get_column_letter(i)].width = min(width, 75)
    return ws


def main():
    source = BASE / INPUT_NAME
    if not source.exists(): raise FileNotFoundError(f'Audit 1 SSOT workbook not found: {source}')
    files = find_sitemap_files()
    if not files: raise FileNotFoundError('No saved sitemap XML files found beside script')
    parsed, errors, index_entries = {}, [], []
    for stem, f in sorted(files.items()):
        try:
            kind, urls = parse_sitemap(f)
            if kind == 'sitemapindex':
                index_entries += [(stem, u) for u in urls]
            else: parsed[stem] = urls
        except Exception as e:
            errors.append((f.name, str(e)))
    index_missing = []
    for index_stem, loc in index_entries:
        expected = Path(urlsplit(loc).path).stem.lower()
        if expected not in parsed:
            index_missing.append((index_stem, loc, expected, 'Referenced sitemap not successfully parsed locally'))
    # If no index is supplied, expected sitemap names are still validated per SSOT URL.
    present = defaultdict(list)
    for stem, urls in parsed.items():
        for url in urls: present[path_of(url)].append((stem, url))

    book = load_workbook(source, read_only=True, data_only=True)
    if 'All URLs' not in book.sheetnames: raise ValueError('Audit 1 workbook missing All URLs worksheet')
    ws = book['All URLs']
    headers = [str(c.value or '').strip() for c in ws[1]]
    required = ['URL in Production XML Sitemap', 'Expected Staging URL', 'Staging HTTP Status']
    if any(h not in headers for h in required): raise ValueError('Missing required SSOT columns: '+str([h for h in required if h not in headers]))
    idx = {h: headers.index(h) for h in headers}
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        prod = row[idx['URL in Production XML Sitemap']]
        staging = row[idx['Expected Staging URL']]
        if not staging: continue
        records.append((str(prod or ''), str(staging), status_of(row[idx['Staging HTTP Status']])))

    # All first segments matching a known local sitemap are local; additional local slugs
    # from SSOT are inferred only when matching US city-state slug conventions.
    local_stems = {s[:-8] for s in files if s.endswith('-sitemap') and s != 'national-sitemap'}
    state_codes = 'al ak az ar ca co ct de fl ga hi id il in ia ks ky la me md ma mi mn ms mo mt ne nv nh nj nm ny nc nd oh ok or pa ri sc sd tn tx ut vt va wa wv wi wy dc'.split()
    state_re = re.compile(r'^[a-z0-9]+(?:-[a-z0-9]+)*-(?:'+'|'.join(state_codes)+r')$')
    all_ssot_paths = {path_of(u) for _,u,_ in records if not is_round_rock(path_of(u))}
    details, exceptions, excluded = [], [], []
    expected_stems = set()
    for prod, staging, status in records:
        path = path_of(staging)
        if is_round_rock(path):
            excluded.append((staging, 'Round Rock excluded from migration scope'))
            continue
        slug = local_slug(path)
        local = slug in local_stems or bool(state_re.fullmatch(slug)) or slug in LOCATION_SITEMAP_OVERRIDES
        target = LOCATION_SITEMAP_OVERRIDES.get(slug, slug+'-sitemap') if local else 'national-sitemap'
        cat = is_blog_category(path)
        exempt_reasons = []
        if is_blog_article_exempt(path): exempt_reasons.append('Blog URL without /categories/')
        if is_space_suffix_exempt(prod): exempt_reasons.append('Production terminal segment has trailing space or /-')
        exempt = bool(exempt_reasons)
        included = present.get(path, [])
        stems = [s for s, _ in included]
        should_include = status == 200 and not cat
        if should_include: expected_stems.add(target)
        issues = []
        if cat:
            if included: issues.append('BLOG_CATEGORY_INCLUDED')
        elif status == 200:
            if not included: issues.append('MISSING_FROM_XML')
            if included and any(s != target for s in stems): issues.append('WRONG_XML_SITEMAP')
        elif status == 301:
            if not exempt: issues.append('UNEXPECTED_301')
            if included: issues.append('301_INCLUDED_IN_XML')
        else:
            issues.append('HTTP_EXCEPTION')
            if included: issues.append('NON_200_INCLUDED_IN_XML')
        if len(included) > 1: issues.append('DUPLICATE_SITEMAP_INCLUSION')
        if should_include and target not in parsed:
            issues.append('EXPECTED_SITEMAP_FILE_MISSING')
        action_map = {
            'MISSING_FROM_XML': f'Add 200 URL to {target}',
            'WRONG_XML_SITEMAP': f'Move URL to {target}; remove from incorrect sitemap(s)',
            'UNEXPECTED_301': 'Investigate and restore 200 response unless formally exempted',
            '301_INCLUDED_IN_XML': 'Remove redirecting URL from all XML sitemaps',
            'BLOG_CATEGORY_INCLUDED': 'Remove blog category URL from all XML sitemaps',
            'HTTP_EXCEPTION': 'Investigate unexpected HTTP response / missing crawl status',
            'NON_200_INCLUDED_IN_XML': 'Remove non-200 URL from XML sitemaps',
            'DUPLICATE_SITEMAP_INCLUSION': 'Remove duplicate occurrences; retain only permitted allocation',
            'EXPECTED_SITEMAP_FILE_MISSING': f'Provide/validate local file for {target} before confirming missing URLs'
        }
        # If target sitemap unavailable, absence cannot be definitively confirmed.
        if 'EXPECTED_SITEMAP_FILE_MISSING' in issues and 'MISSING_FROM_XML' in issues:
            issues.remove('MISSING_FROM_XML')
        if errors or index_missing:
            # Record audit coverage caveat in every result, not false clean certification.
            coverage = 'INCOMPLETE - sitemap parse/index errors' if errors or index_missing else 'COMPLETE'
        else: coverage = 'COMPLETE'
        outcome = 'FAIL' if issues else 'PASS'
        entry = (prod, staging, status if status is not None else '', 'LOCAL' if local else 'NATIONAL',
                 'YES' if exempt else 'NO', '; '.join(exempt_reasons), 'YES' if cat else 'NO',
                 target if should_include else '(not eligible)', '; '.join(stems), len(included),
                 outcome, '; '.join(issues), '; '.join(dict.fromkeys(action_map[i] for i in issues)), coverage)
        details.append(entry)
        for issue in issues:
            exceptions.append((issue, staging, status if status is not None else '', target if should_include else '(not eligible)',
                               '; '.join(stems), action_map[issue], 'YES' if exempt else 'NO'))

    # URLs in XML but not in SSOT: category exclusions and out-of-scope Round Rock are special.
    unexpected = []
    for path, occurrences in sorted(present.items()):
        if is_round_rock(path): continue
        if path not in all_ssot_paths:
            for stem, url in occurrences:
                unexpected.append((url, stem, 'UNEXPECTED_NOT_IN_SSOT', 'Investigate and remove if not an approved addition'))
        if len(occurrences) > 1 and path not in all_ssot_paths:
            unexpected.append((occurrences[0][1], '; '.join(s for s,_ in occurrences), 'DUPLICATE_NOT_IN_SSOT', 'Remove duplicate allocations'))

    coverage_issues = [('PARSE_ERROR', a, b) for a,b in errors]
    coverage_issues += [('INDEX_MISSING_LOCAL_FILE', b, d) for _,b,_,d in index_missing]
    for target in sorted(expected_stems):
        if target not in parsed:
            coverage_issues.append(('EXPECTED_SITEMAP_NOT_LOADED', target, 'Missing/unparseable sitemap for eligible SSOT URLs'))
    if 'national-sitemap' not in parsed:
        coverage_issues.append(('NATIONAL_SITEMAP_NOT_LOADED', 'national-sitemap', 'Cannot validate national allocation'))
    if not index_entries:
        coverage_issues.append(('NO_VALID_SITEMAP_INDEX', 'sitemap_index', 'Index validation unavailable'))

    out = Workbook(); out.remove(out.active)
    counts = Counter(x[0] for x in exceptions)
    summary = [
        ('Audit date', '12 November 2026'), ('Audit input', INPUT_NAME),
        ('SSOT rows', len(records)), ('Round Rock excluded', len(excluded)),
        ('SSOT rows in scope', len(details)), ('SSOT PASS', sum(x[10]=='PASS' for x in details)),
        ('SSOT FAIL', sum(x[10]=='FAIL' for x in details)),
        ('Sitemap files parsed', len(parsed)), ('XML URL occurrences', sum(map(len, parsed.values()))),
        ('Unexpected XML URL occurrences', len(unexpected)),
        ('Coverage issues', len(coverage_issues)),
        ('IMPORTANT', 'Do not certify missing URL findings if sitemap coverage is incomplete'),
    ] + [(k, v) for k,v in sorted(counts.items())]
    sheet(out, 'Summary', ['Metric','Value'], summary)
    sheet(out, 'SSOT Allocation', ['Production URL','Staging URL','Staging HTTP','Scope','Exempt','Calculated Exemption Reason','Blog Category','Expected Sitemap','Actual Sitemap(s)','Occurrences','Result','Exception Code(s)','Developer Action(s)','Coverage'], details)
    sheet(out, 'Developer Exceptions', ['Exception Code','Staging URL','HTTP Status','Expected Sitemap','Actual Sitemap(s)','Developer Action','Exempt'], exceptions)
    sheet(out, 'Unexpected XML URLs', ['XML URL','Actual Sitemap','Exception Code','Developer Action'], unexpected)
    sheet(out, 'Sitemap Inventory', ['Sitemap File','Sitemap Stem','Parsed URLs','Kind'], [(f.name, stem, len(parsed.get(stem,[])), 'URLSET' if stem in parsed else 'INDEX/ERROR') for stem,f in sorted(files.items())])
    sheet(out, 'Coverage Issues', ['Issue','File or URL','Explanation'], coverage_issues)
    sheet(out, 'Round Rock Excluded', ['Staging URL','Reason'], excluded)
    sheet(out, 'Rules', ['Rule','Implementation'], [
        ('SSOT','Audit 1 All URLs worksheet is authoritative'),
        ('200','Exactly once in correct sitemap, except blog categories'),
        ('301 exempt','Allowed HTTP outcome; must be absent from XML'),
        ('301 nonexempt','HTTP exception; must be absent from XML'),
        ('Blog articles','/blog/ without /categories/ is HTTP 301 exemption'),
        ('Blog categories','Exclude /blog/categories and /blog/categories/{category-name}; deeper paths are posts'),
        ('Production space suffix','Production final URL segment trailing space or /- is HTTP 301 exemption'),
        ('Round Rock','Exclude /round-rock-tx and descendants from all audit calculations'),
        ('Local','Location slug determined by sitemap names and state-suffix pattern'),
        ('National','Everything not classified local goes to national-sitemap'),
        ('URL matching','Compare normalized decoded paths, ignoring trailing slash and host'),
        ('Saved HTML sitemap','Parse Rank Math HTML table links as well as raw XML <loc> entries'),
        ('Incomplete input','Flag missing or unparseable sitemap files; no false missing claim')])
    dest = BASE / OUTPUT_NAME
    out.save(dest)
    print('AUDIT 3 COMPLETE:', dest)
    print(f'SSOT: {len(records)} | Round Rock excluded: {len(excluded)} | In scope: {len(details)}')
    print(f'Pass: {sum(x[10]=="PASS" for x in details)} | Fail: {sum(x[10]=="FAIL" for x in details)} | Coverage issues: {len(coverage_issues)}')
    for issue, count in sorted(counts.items()): print(' ',issue, count)
    return dest

if __name__ == '__main__': main()
