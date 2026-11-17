#!/usr/bin/env python3
"""Fill the NSF COA template (Table 4, "A:" co-authors) from a NASA ADS link.

Usage:
    python3 ads_to_coa.py ADS_LINK TEMPLATE.xlsx [--person "Feng, Qi"] [options]
    python3 ads_to_coa.py ADS_LINK --person "Feng, Qi"     (downloads the blank NSF template)
    python3 ads_to_coa.py --person "Feng, Qi" --orcid 0000-0001-6674-4238 --institution Utah "Smithsonian|Harvard"
                                                           (search all of ADS; see README for how papers are vetted)

ADS_LINK may be any of:
    https://ui.adsabs.harvard.edu/search/q=docs(library%2F<ID>)&sort=...
    https://ui.adsabs.harvard.edu/public-libraries/<ID>
    https://ui.adsabs.harvard.edu/user/libraries/<ID>
    a raw ADS query, e.g. 'author:"Feng, Qi" year:2022-2026'

No API token is required: the script uses the same anonymous session the ADS
web site uses. If ADS_API_TOKEN is set (or ~/.ads/dev_key exists), that token
is used instead.

Outputs (--out NAME.xlsx sets the file name; default <template>_ADS.xlsx next to
the template, or coa_<Last>_ADS.xlsx in the current folder for a downloaded one):
    NAME.xlsx  copy of the template with new co-authors appended to Table 4
    NAME.json  papers, co-authors, merges and items to review
The input template is never modified.
"""
import argparse, copy, datetime, html, json, os, re, sys, unicodedata
import urllib.parse, urllib.request
import http.cookiejar

try:
    import openpyxl
    from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
except ImportError:
    sys.exit('openpyxl is required: pip install openpyxl')

API = 'https://api.adsabs.harvard.edu/v1'
UI = 'https://ui.adsabs.harvard.edu'
TABLE4 = 'TableD'  # name of Table 4 in the NSF COA template
NSF_SAMPLE = re.compile(r'Alphaman|Samplename|Test University|Test Company|example@example\.com|'
                        r'University of ABCD|Community College, Virginia Beach|State Dept of Natural Resources')
NSF_TEMPLATE_URL = 'https://nsf-gov-resources.nsf.gov/files/coa_template.xlsx?VersionId=uaz022di3.2KKDC7uauIadw7yeI3Xf.1'


# ---------------------------------------------------------------- ADS access
def ads_token():
    tok = os.environ.get('ADS_API_TOKEN') or os.environ.get('ADS_DEV_KEY')
    keyfile = os.path.expanduser('~/.ads/dev_key')
    if not tok and os.path.exists(keyfile):
        tok = open(keyfile).read().strip()
    if tok:
        return tok
    # anonymous token, as issued to the ADS web UI
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    req = urllib.request.Request(f'{UI}/v1/accounts/bootstrap', headers={'User-Agent': 'ads_to_coa'})
    with opener.open(req, timeout=30) as r:
        return json.load(r)['access_token']


def ads_search(query, token, fq=None):
    fl = 'bibcode,title,pubdate,author,aff,doctype,author_count,orcid_pub,orcid_user,orcid_other'
    docs, start = [], 0
    while True:
        params = {'q': query, 'fl': fl, 'rows': 200, 'start': start, 'sort': 'date desc'}
        if fq:
            params['fq'] = fq
        req = urllib.request.Request(f'{API}/search/query?' + urllib.parse.urlencode(params),
                                     headers={'Authorization': 'Bearer ' + token})
        with urllib.request.urlopen(req, timeout=60) as r:
            resp = json.load(r)['response']
        docs += resp['docs']
        start += 200
        if start >= resp['numFound']:
            return docs


def link_to_query(link):
    link = link.strip()
    m = re.search(r'(?:public-libraries|user/libraries)/([\w-]+)', link)
    if m:
        return f'docs(library/{m.group(1)})'
    if 'adsabs.harvard.edu' in link:
        # ADS puts search params in the path: /search/q=...&sort=...
        tail = link.split('/search/', 1)[-1].split('?', 1)
        for part in '&'.join(tail).split('&'):
            if part.startswith('q='):
                return urllib.parse.unquote_plus(part[2:])
        sys.exit(f'Could not find a query in link: {link}')
    return link  # raw ADS query


FIELDS = {'astronomy': 'collection:astronomy', 'physics': 'collection:(astronomy OR physics)', 'any': None}


def search_query(person, orcid, field):
    """ADS query for every paper that could be this person's: their name (ADS also matches initials) or ORCID."""
    q = f'author:"{person}"'
    if orcid:
        q = f'({q} OR orcid:{orcid})'
    return f'{q} {FIELDS[field]}' if FIELDS[field] else q


def vet_papers(docs, is_me, orcid, institutions, accept_sim=0.6, review_sim=0.4):
    """Decide which search hits are really this person's papers.

    accept: their ORCID is on the matching author slot, or that slot's affiliation matches --institution,
            or (no affiliation given, 20+ authors) the author list is nearly the same as one accepted paper's
            (similarity >= accept_sim: collaboration proceedings, errata)
    review: the author list resembles one accepted paper's (similarity >= review_sim), or a short author
            list with no affiliation where at least half the co-authors appear on accepted papers
    reject: everything else (another person's ORCID on the slot, other institution, unrelated co-authors),
            or the most similar accepted paper also has a namesake who fits the name on this paper
            (e.g. "Smith, J." when the accepted paper lists both "Smith, J. A." and "Smith, J. R.")
    A preprint whose title matches an accepted paper is marked as a duplicate.

    Similarity is |A & B| / |A | B| of the two author lists, compared paper by paper rather than against
    everyone on all accepted papers: a joint paper between two collaborations must not vouch for every
    proceeding of the larger one that lists a "Smith, J.".
    """
    inst = [re.compile(p, re.I) for p in institutions]
    for d in docs:
        authors = d.get('author') or []
        slots = [i for i, a in enumerate(authors) if is_me(a)]
        d['_slot_aff'], d['_status'], d['_why'] = '', 'reject', 'name not in author list'
        d['_slot_names'] = [authors[i] for i in slots]
        if orcid:  # ORCID can identify the slot even when the name is spelled differently
            for f in ('orcid_pub', 'orcid_user', 'orcid_other'):
                slots += [i for i, o in enumerate(d.get(f) or []) if o == orcid and i not in slots]
        for i in slots:
            ids = {(d.get(f) or [''] * len(authors))[i] for f in ('orcid_pub', 'orcid_user', 'orcid_other')} - {'-', ''}
            aff = (d.get('aff') or [''] * len(authors))[i]
            aff = '' if aff == '-' else html.unescape(aff)
            d['_slot_aff'] = aff
            if orcid and orcid in ids:
                d['_status'], d['_why'] = 'accept', 'ORCID match'
                break
            if orcid and ids:
                d['_why'] = f'different ORCID on this author ({", ".join(sorted(ids))})'
                continue
            hit = next((p.pattern for p in inst if p.search(aff)), None)
            if hit:
                d['_status'], d['_why'] = 'accept', f'affiliation matches "{hit}"'
                break
            d['_why'] = 'no affiliation listed' if not aff else 'affiliation does not match --institution'

    # co-author overlap with accepted papers, for those still undecided
    def key(a):
        l, f = split_name(a)
        return last_key(l) + '|' + ''.join(t[0] for t in given(f))
    def coauthors(d):
        return {key(a) for a in d.get('author') or [] if not is_me(a)}
    accepted = [(d['bibcode'], coauthors(d)) for d in docs if d['_status'] == 'accept']
    by_bib = {d['bibcode']: d for d in docs}

    def namesakes(slot_names, bibcode):
        out = []
        for a in by_bib[bibcode].get('author') or []:
            l, f = split_name(a)
            if is_me(a):
                continue
            for n in slot_names:
                nl, nf = split_name(n)
                if last_key(l) == last_key(nl) and given(f) and given(nf) and compatible(given(f), given(nf)):
                    out.append(a)
        return out
    known = set().union(*(s for _, s in accepted))
    title = lambda d: re.sub(r'[^a-z0-9]', '', fold((d.get('title') or [''])[0]))
    titles = {title(d): d['bibcode'] for d in docs if d['_status'] == 'accept'}
    for d in docs:
        if d['_status'] == 'accept' or d['_why'].startswith('different ORCID') or d['_why'] == 'name not in author list':
            continue
        if title(d) in titles:
            d['_why'] = f'duplicate (e.g. preprint) of {titles[title(d)]}'
            continue
        others = coauthors(d)
        sim, like = max(((len(others & s) / max(1, len(others | s)), b) for b, s in accepted), default=(0, ''))
        shared = len(others & known)
        d['_overlap'] = sim
        no_aff = d['_why'] == 'no affiliation listed'
        twins = namesakes(d['_slot_names'], like) if like and sim >= review_sim else []
        if twins:
            d['_why'] += f'; probably {twins[0]}, who is also on the similar paper {like}'
            continue
        if no_aff and len(others) >= 20 and sim >= accept_sim:
            d['_status'] = 'accept'
        elif sim >= review_sim or (no_aff and len(others) < 10 and shared >= 2 and shared >= 0.5 * len(others)):
            d['_status'] = 'review'
        d['_why'] += f'; author list {sim:.0%} like {like}' if like else ''
        d['_why'] += f'; {shared} of {len(others)} co-authors on accepted papers'
    return docs


def orcids_seen(docs, is_me):
    """ORCID iDs on this person's author slot in accepted papers, most common first."""
    seen = {}
    for d in docs:
        if d['_status'] != 'accept':
            continue
        for i, a in enumerate(d.get('author') or []):
            if is_me(a):
                for f in ('orcid_pub', 'orcid_user', 'orcid_other'):
                    o = (d.get(f) or [])[i:i + 1]
                    if o and o[0] not in ('-', ''):
                        seen.setdefault(o[0], set()).add(d['bibcode'])
    return sorted(((o, len(b)) for o, b in seen.items()), key=lambda x: -x[1])


# ---------------------------------------------------------------- names
def fold(s):
    return unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode().lower()


def split_name(name):
    if ',' in name:
        last, first = [p.strip() for p in name.split(',', 1)]
    else:  # "First Last"
        parts = name.split()
        last, first = parts[-1], ' '.join(parts[:-1])
    return last, first


def last_key(last):
    return re.sub(r'[^a-z]', '', fold(last))


def given(first):
    return [t for t in re.split(r'[\s.\-]+', fold(first)) if t]


def compatible(g1, g2):
    """Given-name token lists agree on initials, and on full names where both have them."""
    if not g1 or not g2:
        return True
    for a, b in zip(g1, g2):
        if a[0] != b[0]:
            return False
        if len(a) > 1 and len(b) > 1 and not (a.startswith(b) or b.startswith(a)):
            return False
    return True


LATIN1_MAP = str.maketrans({'ł': 'l', 'Ł': 'L', 'đ': 'd', 'Đ': 'D', 'ı': 'i', 'ø': 'o', 'Ø': 'O',
                            '∣': '|', '─': '-', '–': '-', '—': '-', '‐': '-', '‑': '-', '−': '-',
                            '‘': "'", '’': "'", '“': '"', '”': '"', '•': '-', '​': '', '﻿': ''})


def latin1(s):
    """Research.gov rejects most characters outside Latin-1 (ć, š, ł, ∣, ...): fold them to plain letters,
    keep Latin-1 accents (é, ü, á), drop control characters."""
    out = []
    for ch in unicodedata.normalize('NFC', s.translate(LATIN1_MAP)):
        if ord(ch) > 255:
            ch = ''.join(c for c in unicodedata.normalize('NFKD', ch) if ord(c) < 256 and not unicodedata.combining(c))
        elif ord(ch) < 32 or 127 <= ord(ch) < 160:
            ch = ' ' if ch in '\t\n\r\xa0' else ''
        out.append(ch)
    return re.sub(r'\s+', ' ', ''.join(out)).strip(' ,;')


IS_GROUP = re.compile(r'collaboration|consortium|^the |team$|group$', re.I)

STREET = re.compile(r'\b(str\.|strasse|straße|street|st\.|avenue|ave\.?|road|rd\.|blvd|boulevard|bag|box|c/|via|'
                    r'calle|rue|chaussée|plaza|place|pl\.|allee|weg|ul\.|cedex|campus|bldg|building|p\.o\.)', re.I)


def clean_aff(aff):
    """First affiliation only, without street address or postcode (NSF: no street address, <=255 chars)."""
    # ADS separates multiple affiliations with ';' but also HTML-escapes '&' as '&amp;'
    aff = re.sub(r'(&(?:amp|lt|gt|quot|#\d+));', '\\1\0', aff or '')
    aff = html.unescape(aff.split(';')[0].replace('\0', ';')).strip()
    aff = re.sub(r'\s+', ' ', aff)
    if aff in ('', '-'):
        return ''
    segs = [s.strip() for s in aff.split(',') if s.strip()]
    keep = [s for s in segs if not re.search(r'\d', s) and not STREET.search(s)] or segs[:1]
    out = ', '.join(keep)
    if len(out) > 150 and len(keep) > 2:
        out = ', '.join(keep[:2] + keep[-1:])
    return out[:255]


def collect_people(docs, is_me, max_authors):
    """One record per unique person; most recent paper wins for affiliation and date."""
    recs = {}
    for d in sorted(docs, key=lambda d: d['pubdate']):
        authors = d.get('author') or []
        if max_authors and len(authors) > max_authors:
            continue
        affs = d.get('aff') or [''] * len(authors)
        date = d['pubdate'][:7]
        for a, af in zip(authors, affs):
            if IS_GROUP.search(a) or is_me(a):
                continue
            af = clean_aff(af)
            r = recs.setdefault(a, dict(name=a, aff='', date=date, papers=[]))
            r['date'] = date
            r['aff'] = af or r['aff']
            r['papers'].append(d['bibcode'])

    # cluster spelling variants: same folded last name + compatible given names
    clusters, by_last = [], {}
    for r in recs.values():
        last, first = split_name(r['name'])
        r['lk'], r['g'] = last_key(last), given(first)
        for c in by_last.get(r['lk'], []):
            if all(compatible(r['g'], m['g']) for m in c):
                c.append(r)
                break
        else:
            c = [r]
            by_last.setdefault(r['lk'], []).append(c)
            clusters.append(c)

    people = []
    for c in clusters:
        best = max(c, key=lambda m: (sum(len(t) for t in m['g']), len(m['name'])))
        newest = sorted(c, key=lambda m: m['date'], reverse=True)
        people.append(dict(
            name=best['name'], affiliation=next((m['aff'] for m in newest if m['aff']), ''),
            last_active=newest[0]['date'], variants=sorted(m['name'] for m in c),
            papers=sorted({p for m in c for p in m['papers']}), lk=best['lk'], g=best['g']))
    return people


# ---------------------------------------------------------------- workbook
def table1_name(template):
    """The "Your Name" entry of Table 1 (TableA), if filled in."""
    ws = openpyxl.load_workbook(template).worksheets[0]
    if 'TableA' not in ws.tables:
        return None
    t = CellRange(ws.tables['TableA'].ref)
    v = ws.cell(t.min_row + 1, t.min_col + 1).value
    return v.strip() if isinstance(v, str) and v.strip() else None


def shift_below(ws, wb, row, n):
    """Insert n rows after `row`, moving everything below and fixing tables, validations and merges."""
    max_row, max_col = ws.max_row, ws.max_column
    if max_row > row:
        heights = {r: ws.row_dimensions[r].height for r in range(row + 1, max_row + 1)}
        last_col = openpyxl.utils.get_column_letter(max_col)
        ws.move_range(f'A{row + 1}:{last_col}{max_row}', rows=n)
        for r in range(max_row, row, -1):
            ws.row_dimensions[r + n].height = heights[r]
        for r in range(row + 1, row + n + 1):
            ws.row_dimensions[r].height = None

    def moved(ref):
        cr = CellRange(ref)
        if cr.min_row > row:
            cr.shift(row_shift=n)
        elif cr.max_row >= row:  # range ends at or spans the insertion point: grow it
            cr.expand(down=n)
        return cr.coord

    for t in ws.tables.values():
        if t.name == TABLE4 or CellRange(t.ref).min_row > row:
            t.ref = moved(t.ref)
            if t.autoFilter is not None:
                t.autoFilter.ref = t.ref
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange(' '.join(moved(str(c)) for c in dv.sqref.ranges))
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row > row:
            ws.merged_cells.remove(mc)
            mc.shift(row_shift=n)
            ws.merged_cells.add(mc)


def download_template(path):
    """Fetch the blank NSF COA template (kept for reuse if it is already there)."""
    if not os.path.exists(path):
        print(f'Downloading blank NSF COA template to {path}')
        req = urllib.request.Request(NSF_TEMPLATE_URL, headers={'User-Agent': 'ads_to_coa'})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = r.read()
        if not data.startswith(b'PK'):
            sys.exit(f'Download from {NSF_TEMPLATE_URL} is not an .xlsx file; download it by hand and pass its path.')
        with open(path, 'wb') as f:
            f.write(data)
    return path


def fill_template(template, out, people, report, person=None, clear_samples=False):
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    if TABLE4 not in ws.tables:
        sys.exit(f'{template}: no table named {TABLE4!r}; is this the NSF COA template?')
    # the blank NSF template ships with example rows ("Alphaman, Alan K", ...): remove them from any template
    for t in ws.tables.values():
        ref = CellRange(t.ref)
        for r in range(ref.min_row + 1, ref.max_row + 1):
            cells = [ws.cell(r, c) for c in range(ref.min_col, ref.max_col + 1)]
            text = ' '.join(str(c.value) for c in cells if c.value is not None)
            if clear_samples or NSF_SAMPLE.search(text) or any(c.hyperlink and 'example.com' in str(c.hyperlink.target)
                                                               for c in cells):
                if text and not clear_samples:
                    report['sample_rows_removed'].append(dict(row=r, text=text))
                for c in cells:
                    c.value = None
                    c.hyperlink = None
    if person and 'TableA' in ws.tables:  # put the name in a blank Table 1
        t1 = CellRange(ws.tables['TableA'].ref)
        name_cell = ws.cell(t1.min_row + 1, t1.min_col + 1)
        if not (isinstance(name_cell.value, str) and name_cell.value.strip()):
            name_cell.value = latin1(person)
    t4 = CellRange(ws.tables[TABLE4].ref)
    first, last = t4.min_row + 1, t4.max_row

    # everyone already named anywhere on the sheet (column B)
    existing = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.strip() and not v.strip().endswith(':') and len(v) < 80:
            l, f = split_name(v.strip())
            existing.append(dict(row=r, name=v.strip(), lk=last_key(l), g=given(f)))

    new = []
    for p in people:
        same = [e for e in existing if e['lk'] == p['lk'] and e['g'] and p['g'] and e['g'][0][0] == p['g'][0][0]]
        hit = next((e for e in same if compatible(e['g'], p['g'])), None)
        if not hit:
            if same:  # e.g. sheet "Smith, Andrew W." vs ADS "Smith, A. J.": add, but flag it
                report['possible_duplicates_of_sheet'].append(dict(sheet=same[0]['name'], ads=p['name'],
                                                                   sheet_row=same[0]['row']))
            new.append(p)
            continue
        report['already_listed'].append(dict(sheet=hit['name'], ads=p['name'], row=hit['row'],
                                             spelling_differs=hit['name'] != p['name']))
        # bring a stale Table 4 "Last Active" date forward
        if first <= hit['row'] <= last:
            cur, nd = ws.cell(hit['row'], 5).value, datetime.datetime.strptime(p['last_active'], '%Y-%m')
            if isinstance(cur, datetime.datetime) and cur < nd:
                ws.cell(hit['row'], 5).value = nd
                report['date_updates'].append(dict(name=hit['name'], old=cur.strftime('%m/%d/%Y'),
                                                   new=nd.strftime('%m/%d/%Y')))
    new.sort(key=lambda p: fold(p['name']))

    # use blank rows inside Table 4 first, then grow the table
    blanks = [r for r in range(first, last + 1)
              if not any(ws.cell(r, c).value not in (None, '') for c in (1, 2, 3))]
    extra = max(0, len(new) - len(blanks))
    if extra:
        shift_below(ws, wb, last, extra)
    rows = blanks + list(range(last + 1, last + 1 + extra))

    style_row = last if any(ws.cell(last, c).value for c in (1, 2)) else first
    tpl = {c: copy.copy(ws.cell(style_row, c)._style) for c in range(1, 6)}
    for r, p in zip(rows, new):
        p['ads_name'], p['name'], p['affiliation'] = p['name'], latin1(p['name']), latin1(p['affiliation'])
        vals = ['A:', p['name'], p['affiliation'], None,
                datetime.datetime.strptime(p['last_active'], '%Y-%m')]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c)
            cell._style = copy.copy(tpl[c])
            cell.value = v
        ws.cell(r, 5).number_format = 'mm/dd/yyyy'
        p['row'] = r

    # entries already in the template must pass Research.gov's character check too
    for t in ws.tables.values():
        ref = CellRange(t.ref)
        for r in range(ref.min_row + 1, ref.max_row + 1):
            for c in range(ref.min_col, ref.max_col + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and latin1(v) != v.strip():
                    ws.cell(r, c).value = latin1(v)
                    if r not in rows:
                        report['characters_replaced'].append(dict(row=r, old=v, new=latin1(v)))
    wb.save(out)
    return new


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('link', nargs='?', help='ADS library/search URL or raw ADS query. Leave out to search all of '
                    'ADS for --person (then give --orcid and/or --institution)')
    ap.add_argument('template', nargs='?',
                    help='NSF COA template .xlsx (blank or already partly filled). If omitted, the current '
                         'blank template is downloaded from nsf.gov (needs --person)')
    ap.add_argument('--person', help='the person this COA is for, "Last, First" (e.g. "Feng, Qi"); excluded '
                    'from the co-author list. Default: the "Your Name" cell in Table 1 of the template')
    ap.add_argument('--months', type=int, default=48, help='look-back window (default 48, per NSF)')
    ap.add_argument('--max-authors', type=int, default=0,
                    help='skip papers with more authors than this (default 0 = keep all)')
    ap.add_argument('--orcid', help='search mode: your ORCID iD; papers carrying it on your author slot are accepted')
    ap.add_argument('--institution', nargs='+', default=[], metavar='PATTERN',
                    help='search mode: your institutions in the window, as text or regular expressions '
                         '(e.g. "Utah" "Smithsonian|Harvard"); matched against your own affiliation on each paper')
    ap.add_argument('--field', choices=sorted(FIELDS), default='physics',
                    help='search mode: ADS collections to search (physics = astronomy + physics; default)')
    ap.add_argument('--include', nargs='+', default=[], metavar='BIBCODE',
                    help='search mode: papers to accept anyway (e.g. confirmed from the "review" list)')
    ap.add_argument('--exclude', nargs='+', default=[], metavar='BIBCODE', help='papers to leave out')
    ap.add_argument('--out', help='output .xlsx file name; the report is written next to it with .json '
                    '(default: <template>_ADS.xlsx, or coa_<Last>_ADS.xlsx for a downloaded template)')
    a = ap.parse_args()

    today = datetime.date.today()
    y, m = divmod(today.year * 12 + today.month - 1 - a.months, 12)
    since = f'{y}-{m + 1:02d}'
    if a.link and a.link.lower().endswith('.xlsx') and not a.template:
        a.link, a.template = None, a.link
    if not a.link and not a.person:
        sys.exit('To search all of ADS, give --person "Last, First" (plus --orcid and/or --institution).')
    if not a.link and not (a.orcid or a.institution):
        sys.exit('Searching all of ADS by name alone is not reliable: add --orcid and/or --institution.')
    if a.orcid and not re.fullmatch(r'\d{4}-\d{4}-\d{4}-\d{3}[\dX]', a.orcid):
        sys.exit(f'--orcid should look like 0000-0001-6674-4238, not {a.orcid!r}')
    if not a.template and not a.person:
        sys.exit('Without a template, give --person "Last, First".')
    person = a.person or table1_name(a.template)
    if not person:
        sys.exit('Give --person "Last, First", or fill in "Your Name" in Table 1 of the template.')
    downloaded = not a.template
    if downloaded:
        last = re.sub(r'[^A-Za-z]', '', latin1(split_name(person)[0]))
        a.out = a.out or os.path.join(os.getcwd(), f'coa_{last}_ADS.xlsx')
        a.template = download_template(os.path.join(os.path.dirname(os.path.abspath(a.out)), 'coa_template_NSF_blank.xlsx'))
    me_last, me_first = split_name(person)
    me_lk, me_g = last_key(me_last), given(me_first)

    def is_me(name):  # "Feng, Q." and "Feng, Qi" match "Feng, Qi"; "Feng, Qian" does not; accents ignored
        last, first = split_name(name)
        g = given(first)
        return last_key(last) == me_lk and bool(g) and all(
            x == y if len(x) > 1 and len(y) > 1 else x[0] == y[0] for x, y in zip(g, me_g))

    query = link_to_query(a.link) if a.link else search_query(person, a.orcid, a.field)
    print(f'COA for: {person}\nADS query: {query}\nPapers since {since}')
    token = ads_token()
    found = ads_search(query, token, fq=f'pubdate:[{since} TO 9999-12]')
    if a.link:  # a library or query you chose: every paper counts
        for d in found:
            d['_status'], d['_why'], d['_slot_aff'] = 'accept', 'from ADS link', ''
    else:
        vet_papers(found, is_me, a.orcid, a.institution)
    for d in found:
        if d['bibcode'] in a.include:
            d['_status'], d['_why'] = 'accept', '--include'
    missing = [b for b in a.include if b not in {d['bibcode'] for d in found}]
    if missing:
        for d in ads_search('bibcode:(' + ' OR '.join(missing) + ')', token):
            d['_status'], d['_why'], d['_slot_aff'] = 'accept', '--include', ''
            found.append(d)
    for d in found:
        if d['bibcode'] in a.exclude:
            d['_status'], d['_why'] = 'reject', '--exclude'
    docs = [d for d in found if d['_status'] == 'accept']
    people = collect_people(docs, is_me, a.max_authors)

    out_xlsx = a.out or os.path.splitext(a.template)[0] + '_ADS.xlsx'
    if not out_xlsx.lower().endswith('.xlsx'):
        out_xlsx += '.xlsx'
    if os.path.abspath(out_xlsx) == os.path.abspath(a.template):
        sys.exit('--out must differ from the template; the template is never overwritten.')
    out_json = os.path.splitext(out_xlsx)[0] + '.json'
    report = dict(person=person, query=query, since=since, generated=today.isoformat(), max_authors=a.max_authors,
                  papers=[dict(status=d['_status'], why=d['_why'], bibcode=d['bibcode'], date=d['pubdate'][:7],
                               n_authors=len(d.get('author') or []), your_affiliation=d['_slot_aff'],
                               title=(d.get('title') or [''])[0], link=f'{UI}/abs/{d["bibcode"]}/abstract')
                          for d in sorted(found, key=lambda d: (('accept', 'review', 'reject').index(d['_status']),
                                                                -d.get('_overlap', 0), d['bibcode']))],
                  already_listed=[], date_updates=[], characters_replaced=[], sample_rows_removed=[], possible_duplicates_of_sheet=[])
    new = fill_template(a.template, out_xlsx, people, report, person, clear_samples=downloaded)

    by_last = {}
    for p in people:
        by_last.setdefault(p['lk'], []).append(p)
    report.update(
        added=[{k: p[k] for k in ('name', 'ads_name', 'affiliation', 'last_active', 'variants', 'papers', 'row')} for p in new],
        review=dict(
            initials_only_merges=[p['variants'] for p in people if len(p['variants']) > 1 and
                                  any(all(len(t) == 1 for t in given(split_name(v)[1])) for v in p['variants'])
                                  and len({v for v in p['variants']}) > 1],
            same_last_name_and_initial=[[q['name'] for q in g] for g in by_last.values() if len(g) > 1 and
                                        len({q['g'][0][0] for q in g if q['g']}) < len(g)],
            possible_duplicates_of_sheet=report['possible_duplicates_of_sheet'],
            no_affiliation=[p['name'] for p in new if not p['affiliation']]))
    seen = orcids_seen(found, is_me) if not a.link else []
    report['orcids_on_accepted_papers'] = seen
    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=1, ensure_ascii=False)

    n = {s: sum(d['_status'] == s for d in found) for s in ('accept', 'review', 'reject')}
    if not a.link:
        print(f'{len(found)} papers found: {n["accept"]} accepted, {n["review"]} to review, {n["reject"]} rejected')
        if n['review']:
            print('  Check the "review" papers in the .json; add yours with --include BIBCODE ...')
        if seen and not a.orcid:
            print(f'  ORCID on your accepted papers: {seen[0][0]} ({seen[0][1]} papers); '
                  f'if it is yours, rerun with --orcid {seen[0][0]}')
        elif seen and seen[0][0] != a.orcid:
            print(f'  Warning: accepted papers mostly carry ORCID {seen[0][0]}, not {a.orcid}')
    print(f'{len(docs)} papers used, {len(people)} unique co-authors, {len(report["already_listed"])} already listed, '
          f'{len(new)} added, {len(report["date_updates"])} dates updated')
    print(f'Wrote {out_xlsx} and {out_json}  (open the .xlsx in Excel and re-save once before uploading)')


if __name__ == '__main__':
    main()
