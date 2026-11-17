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
    fl = 'bibcode,title,pub,pubdate,author,aff,doctype,author_count,orcid_pub,orcid_user,orcid_other'
    docs, start = [], 0
    while True:
        # bibcode breaks ties: sorting by date alone lets ADS shuffle same-date papers between pages
        params = {'q': query, 'fl': fl, 'rows': 200, 'start': start, 'sort': 'date desc, bibcode desc'}
        if fq:
            params['fq'] = fq
        req = urllib.request.Request(f'{API}/search/query?' + urllib.parse.urlencode(params),
                                     headers={'Authorization': 'Bearer ' + token})
        with urllib.request.urlopen(req, timeout=60) as r:
            resp = json.load(r)['response']
        docs += resp['docs']
        start += 200
        if start >= resp['numFound'] or not resp['docs']:
            seen = set()
            return [d for d in docs if not (d['bibcode'] in seen or seen.add(d['bibcode']))]


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
    def coauthors(d):  # people only: a "VERITAS Collaboration" author entry says nothing about who wrote it
        return {key(a) for a in d.get('author') or [] if not is_me(a) and not IS_GROUP.search(a)}
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
        elif sim >= review_sim or (no_aff and len(others) < 10 and shared >= 1 and shared >= 0.5 * len(others)):
            d['_status'] = 'review'  # e.g. a 2-author ATel whose co-author is on your accepted papers
        d['_why'] += f'; author list {sim:.0%} like {like}' if like else ''
        d['_why'] += f'; {shared} of {len(others)} co-authors on accepted papers'
        if not others:
            d['_why'] += ' (no individual co-authors: does not change the COA)'
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


IS_GROUP = re.compile(r'collaboration|consortium|\bproject\b|^the |team$|group$|'
                      r'^(ctao?|veritas|icecube|h\.?e\.?s\.?s\.?|magic|hawc|fermi[- ]lat|lhaaso|eht|ligo|virgo|kagra|'
                      r'km3net|antares|auger|nanograv|rno-g)$', re.I)

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
            r = recs.setdefault(a, dict(name=a, aff='', date=date, papers=[], raw={}))
            raw = html.unescape(af or '').strip()
            if raw not in ('', '-'):
                r['raw'][raw] = d['bibcode']
            af = clean_aff(af)
            r['date'] = date
            r['aff'] = af or r['aff']
            r['papers'].append(d['bibcode'])

    # cluster spelling variants: same folded last name + compatible given names
    clusters, by_last = [], {}
    for r in recs.values():
        last, first = split_name(r['name'])
        r['lk'], r['g'] = last_key(last), given(first)
        for c in by_last.get(r['lk'], []):
            # same initials AND, when both list one, affiliations naming the same place: "Williams, D." at Glasgow
            # is not merged into "Williams, D. R." at Alabama (a missing COA entry is worse than an extra row)
            if all(compatible(r['g'], m['g']) and (not r['aff'] or not m['aff'] or same_place([r['aff'], m['aff']]))
                   for m in c):
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
            papers=sorted({p for m in c for p in m['papers']}), lk=best['lk'], g=best['g'],
            raw={k: v for m in c for k, v in m['raw'].items()},
            variant_affs={m['name']: m['aff'] for m in c}))
    return people


# ---------------------------------------------------------------- paper list for checking
# Collaboration names, matched against group authors ("VERITAS Collaboration") and titles.
# Source names such as "HESS J1745-290" or "MAGIC J0223+430" are not counted.
NOT_SRC = r'(?![\s-]*J\d)'
COLLABS = [
    ('CTAO', r'\bCTAO?\b|Cherenkov Telescope Array|\bLST-?\d?\b|Large[- ]Sized Telescope|Schwarzschild[- ]Couder|\bp?SCT\b'),
    ('VERITAS', r'\bVERITAS\b'),
    ('IceCube', r'IceCube'),
    ('LIGO-Virgo-KAGRA', r'\bLIGO\b|\bKAGRA\b|Virgo Collaboration'),
    ('EHT', r'Event Horizon Telescope|\bEHT\b'),
    ('Fermi-LAT', r'Fermi[- ]LAT|Fermi Large Area Telescope'),
    ('H.E.S.S.', r'H\.\s?E\.\s?S\.\s?S\.' + NOT_SRC + r'|\bHESS\b' + NOT_SRC),
    ('MAGIC', r'\bMAGIC\b' + NOT_SRC),
    ('HAWC', r'\bHAWC\b' + NOT_SRC),
    ('LHAASO', r'\bLHAASO\b' + NOT_SRC),
    ('KM3NeT', r'KM3NeT|ANTARES'),
    ('Pierre Auger', r'Pierre Auger'),
    ('NANOGrav', r'NANOGrav'),
    ('RNO-G', r'RNO-G'),
]


def classify_papers(docs, is_me, min_authors=20, min_overlap=0.3):
    """Tag each paper with the collaborations it belongs to (d['_collab']).

    1. A collaboration named as a group author counts; one named in the title counts for papers with
       min_authors+ authors (a short paper about VERITAS data is a small-team paper, not a collaboration one).
    2. A paper with min_authors+ authors and no name match takes the collaborations of the most similar
       tagged paper, if their author lists overlap by min_overlap or more (marked d['_collab_how'] = 'overlap').
    """
    pats = [(n, re.compile(r, re.I if n not in ('MAGIC', 'HAWC', 'EHT', 'CTAO') else 0)) for n, r in COLLABS]
    def key(a):
        l, f = split_name(a)
        return last_key(l) + '|' + ''.join(t[0] for t in given(f))
    for d in docs:
        groups = ' ; '.join(a for a in d.get('author') or [] if IS_GROUP.search(a))
        big = len(d.get('author') or []) >= min_authors
        text = groups + (' ; ' + ' '.join(d.get('title') or ['']) if big else '')
        d['_collab'] = [n for n, p in pats if p.search(text)]
        if big:  # e.g. "Japanese MAGIC Group: ICRR" on 65 authors of a paper whose only group author is VERITAS
            affs = [a for a in d.get('aff') or [] if a and a != '-']
            d['_collab'] += [n for n, p in pats if n not in d['_collab'] and sum(bool(p.search(a)) for a in affs) >= 5]
            d['_collab'].sort(key=[n for n, _ in COLLABS].index)
        d['_collab_how'] = 'name' if d['_collab'] else ''
        d['_size_note'] = ''
        d['_keys'] = {key(a) for a in d.get('author') or [] if not is_me(a) and not IS_GROUP.search(a)}
    tagged = [d for d in docs if d['_collab']]
    for d in docs:
        if d['_collab'] or len(d['_keys']) < min_authors:
            continue
        sim, best = max(((len(d['_keys'] & t['_keys']) / max(1, len(d['_keys'] | t['_keys'])), t) for t in tagged),
                        default=(0, None), key=lambda x: x[0])
        if best and sim >= min_overlap:
            d['_collab'], d['_collab_how'] = list(best['_collab']), 'overlap'
    # an erratum or preprint shares its paper's collaborations (ADS often has no affiliations on errata)
    base = lambda d: re.sub(r'[^a-z0-9]', '', re.sub(r'^\s*(correction|erratum|corrigendum|addendum)\s*(to)?\s*:?', '',
                                                      fold((d.get('title') or [''])[0])))
    by_title = {}
    for d in docs:
        by_title.setdefault(base(d), []).append(d)
    for same in by_title.values():
        if len(same) > 1:
            union = sorted({c for d in same for c in d['_collab']}, key=[n for n, _ in COLLABS].index)
            for d in same:
                if union and set(d['_collab']) != set(union):
                    d['_collab'], d['_collab_how'] = union, d['_collab_how'] or 'name'

    # a paper far larger than the collaboration's other papers is probably a joint paper not named as one
    sizes = {}
    for d in docs:
        if len(d['_collab']) == 1:
            sizes.setdefault(d['_collab'][0], []).append(len(d.get('author') or []))
    for d in docs:
        if len(d['_collab']) == 1 and len(sizes[d['_collab'][0]]) >= 3:
            med = sorted(sizes[d['_collab'][0]])[len(sizes[d['_collab'][0]]) // 2]
            n = len(d.get('author') or [])
            if n > 3 * med and n - med > 100:
                d['_size_note'] = (f'{n} authors vs typically {med} on {d["_collab"][0]} papers: '
                                   f'probably a joint paper with other collaborations')
    for d in docs:
        del d['_keys']
    return docs


SMALL_TEAM = 20  # papers with fewer authors and no collaboration tag count as small-team papers


def paper_sections(docs, max_authors, review=()):
    """Group and number papers: [(section, subsection, [docs])]. Sets d['_num'] and d['_flags'] on each."""
    order = [n for n, _ in COLLABS]
    single, joint, other = {}, {}, []
    for d in sorted(docs, key=lambda d: d['pubdate'], reverse=True):
        c = d.get('_collab') or []
        if len(c) == 1:
            single.setdefault(c[0], []).append(d)
        elif c:
            joint.setdefault(' + '.join(sorted(c, key=order.index)), []).append(d)
        else:
            other.append(d)
    secs = [(n, '', single[n]) for n in sorted(single, key=lambda n: (order.index(n) > 1, -len(single[n]),
                                                                       order.index(n)))]
    secs += [('Joint collaboration papers', n, joint[n]) for n in sorted(joint, key=lambda n: -len(joint[n]))]
    if other:
        secs.append(('Small-team and other papers', '', other))
    if review:
        secs.append(('Not used: to review', '', sorted(review, key=lambda d: d['pubdate'], reverse=True)))
    seen, i = {}, 0
    for _, _, ds in secs:
        for d in ds:
            i += 1
            d['_num'] = i
            na = len(d.get('author') or [])
            d['_title'] = re.sub(r'<[^>]+>', '', html.unescape((d.get('title') or [''])[0]))
            flags = ['collaboration guessed from author overlap'] if d.get('_collab_how') == 'overlap' else []
            if d.get('_size_note'):
                flags.append(d['_size_note'])
            if max_authors and na > max_authors:
                flags.append(f'skipped: {na} authors > {max_authors}')
            tk = re.sub(r'[^a-z0-9]', '', re.sub(r'^\s*(correction|erratum|corrigendum|addendum)\s*(to)?\s*:?', '',
                                                 fold(d['_title'])))
            if tk in seen:
                flags.append(f'same title as #{seen[tk]}')
            else:
                seen[tk] = i
            d['_flags'] = flags
    return secs


def paper_list(docs, max_authors, person, review=()):
    """Markdown list of the papers used, numbered and grouped by collaboration."""
    secs = paper_sections(docs, max_authors, review)
    n_used = sum(not (max_authors and len(d.get('author') or []) > max_authors) for d in docs)
    out = [f'# Papers used for the COA of {person}', '',
           f'{len(docs)} accepted papers; {n_used} used for co-authors'
           + (f' ({len(docs) - n_used} skipped with more than {max_authors} authors)' if n_used < len(docs) else ''),
           '', 'Grouped by collaboration (named as an author or in the title; "~" = guessed from author overlap).', '']
    def line(d):
        flags = ''.join(' ~' if f.startswith('collaboration guessed') else f' **[{f}]**' for f in d['_flags'])
        out_ = (f'{d["_num"]}. {d["pubdate"][:7]} | {d["_title"]} | {d.get("pub") or ""} | '
                f'{len(d.get("author") or [])} authors | [{d["bibcode"]}]({UI}/abs/{d["bibcode"]}/abstract){flags}')
        if '_out' in d:  # review papers: who is and is not already in the spreadsheet
            out_ += ('\n   - NOT in spreadsheet: ' + ('; '.join(a for _, a in sorted(d['_out'], key=lambda x: fold(x[1])))
                                                     or 'none') +
                     '\n   - already in spreadsheet: ' + ('; '.join(f'{a} (row {r})' for r, a in
                                                                  sorted(d['_in'], key=lambda x: fold(x[1]))) or 'none'))
        return out_
    last = None
    for sec, sub, ds in secs:
        if sec != last:
            n = sum(len(x) for s_, _, x in secs if s_ == sec)
            out += [f'## {sec} ({n})', '']
            if sec.startswith('Not used'):
                out += ['Add yours with --include BIBCODE ...', '']
            last = sec
        if sub:
            out += [f'### {sub} ({len(ds)})', '']
        out += [line(d) for d in ds] + ['']
    return '\n'.join(out)


# ---------------------------------------------------------------- co-author groups and checks
GROUP_DIRECT = 'Direct collaborators (small-team papers)'
GROUP_SEVERAL = 'Several collaborations'
GROUP_GUESSED = 'Questionable: collaboration guessed'
GROUP_UNTAGGED = 'Questionable: large paper, no collaboration'


def assign_groups(people, docs):
    """Give each co-author one group (p['group']) from the papers they share with you."""
    by_bib = {d['bibcode']: d for d in docs}
    for p in people:
        ps = [by_bib[b] for b in p['papers'] if b in by_bib]
        p['paper_nums'] = sorted(d['_num'] for d in ps if '_num' in d)
        if any(not d['_collab'] and len(d.get('author') or []) < SMALL_TEAM for d in ps):
            p['group'] = GROUP_DIRECT
            continue
        tagged = [d for d in ps if d['_collab']]
        if not tagged:
            p['group'] = GROUP_UNTAGGED
            continue
        if all(d['_collab_how'] == 'overlap' for d in tagged):
            p['group'] = GROUP_GUESSED
            continue
        # the collaboration on most of their papers; a tie between collaborations that always appear
        # together means they are only on joint papers (e.g. an IceCube member on VERITAS + IceCube papers)
        count = {}
        for d in tagged:
            for c in d['_collab']:
                count[c] = count.get(c, 0) + 1
        top = max(count.values())
        best = [c for c in count if count[c] == top]
        p['also'] = sorted(set(count) - set(best), key=[n for n, _ in COLLABS].index)
        if len(best) == 1:
            p['group'] = best[0]
        elif all(set(best) <= set(d['_collab']) for d in tagged if set(best) & set(d['_collab'])):
            p['group'] = 'Joint: ' + ' + '.join(sorted(best, key=[n for n, _ in COLLABS].index))
        else:
            p['group'] = GROUP_SEVERAL
    return people


def group_order(name):
    order = [n for n, _ in COLLABS]
    if name == GROUP_DIRECT:
        return (0, 0, name)
    if name in order:
        return (1, order.index(name), name)
    if name == GROUP_SEVERAL:
        return (2, 0, name)
    if name.startswith('Joint'):
        return (3, 0, name)
    return (4, 0, name)


# Affiliations in these countries are listed for a manual research-security check. Matching is on text, so
# false positives (e.g. Moscow, Idaho) and misses (an institute named without its country) are both possible.
CONCERN = [
    ('China', r'\bChina\b|\bP\.\s?R\.\s?C\b|\bPRC\b|Chinese Academy|Beijing|Shanghai|Nanjing|Hefei|Wuhan|Chengdu|'
              r'Kunming|Tianjin|Guangzhou|Shenzhen|Xi\'?an\b|Lanzhou|Harbin|Changchun|Jinan|Hangzhou|Tsinghua|'
              r'Peking|Shandong|Yunnan|Sichuan|Zhejiang|Guangxi|Hebei|Xinjiang|Tibet|Qinghai'),
    ('Hong Kong', r'Hong Kong'),
    ('Macau', r'\bMaca[uo]\b'),
    ('Cuba', r'\bCuba\b|Havana|La Habana'),
    ('Iran', r'\bIran\b|Tehran|Isfahan|Shiraz|Tabriz|Mashhad'),
    ('North Korea', r'North Korea|\bDPRK\b|Pyongyang|Democratic People'),
    ('Russia', r'Russia|Moscow|Novosibirsk|Petersburg|Dubna|Troitsk|Protvino|Kazan|Yekaterinburg|Tomsk|Irkutsk|'
               r'Nizhny|\bJINR\b|Lebedev|Ioffe'),
    ('Venezuela', r'Venezuela|Caracas'),
]
NOT_CONCERN = {'China': r'Taiwan|Taipei|Hsinchu|Taoyuan|Tainan|Kaohsiung',
               'Russia': r'Idaho|\bID\b|Florida|\bFL\b|USA|United States'}


def concern_matches(people):
    """People with any affiliation (on any shared paper, any position) that names a listed country."""
    pats = [(c, re.compile(r)) for c, r in CONCERN]
    nots = {c: re.compile(r) for c, r in NOT_CONCERN.items()}
    hits = []
    for p in people:
        found = []
        for raw, bib in p.get('raw', {}).items():
            for part in re.split(r';(?![^&]*;)', raw):
                for c, pat in pats:
                    if pat.search(part) and not (c in nots and nots[c].search(part)):
                        found.append((c, part.strip(), bib))
        if found:
            hits.append(dict(name=p['name'], countries=sorted({c for c, _, _ in found}),
                             affiliations=sorted({a for _, a, _ in found}), group=p.get('group', ''),
                             row=p.get('row'), papers=p.get('paper_nums', []),
                             current_affiliation=p['affiliation']))
    return sorted(hits, key=lambda h: (h['countries'], fold(h['name'])))


def review_coauthors(review, people, is_me):
    """For each paper sent to review: its co-authors already in the spreadsheet (from used papers) and the ones
    who are not, i.e. who would be added by --include."""
    index = {}
    for p in people:
        index.setdefault(p['lk'], []).append(p)
    for d in review:
        d['_in'], d['_out'] = [], []
        for a in d.get('author') or []:
            if IS_GROUP.search(a) or is_me(a):
                continue
            l, f = split_name(a)
            g = given(f)
            hit = next((p for p in index.get(last_key(l), []) if g and p['g'] and compatible(g, p['g'])), None)
            (d['_in'] if hit else d['_out']).append((hit.get('row') if hit else None, a))
    return review


AFF_STOP = set("""university universita universitat universite universidad universidade institute institut instituto istituto
national nazionale nacional department dipartimento departamento faculty school college center centre centro laboratory
laboratorio physics fisica physik physique astronomy astrophysics astrophysical astrofisica sciences science research
group sezione section studi studies degli della delle dell interateneo division technology technical engineering
observatory osservatorio observatoire academy and the for with magic infn inaf cnrs csic italy spain germany france
japan usa united states kingdom china republic""".split())


def same_place(affs):
    """True if every pair of non-empty affiliations shares a distinctive word (city, university name, ...)."""
    toks = [{w for w in re.findall(r'[a-z]{4,}', fold(a)) if w not in AFF_STOP} for a in affs if a]
    return all(t1 & t2 for i, t1 in enumerate(toks) for t2 in toks[i + 1:])


def row_ranges(rows):
    rows = sorted(r for r in rows if r)
    out = []
    for r in rows:
        if out and r == out[-1][1] + 1:
            out[-1][1] = r
        else:
            out.append([r, r])
    return ', '.join(f'{a}' if a == b else f'{a}-{b}' for a, b in out)


def short_ranges(rows, limit=60, prefix=''):
    """row_ranges, or 'N rows, first-last (not contiguous)' when that would not fit on a slide."""
    rows = sorted({r for r in rows if r})
    full = prefix + row_ranges(rows).replace(', ', ', ' + prefix)
    if not rows or len(full) <= limit:
        return full
    return f'{len(rows)} {"papers" if prefix else "rows"} in {prefix}{rows[0]}-{prefix}{rows[-1]}'


def num_ranges(nums):
    return '#' + row_ranges(nums).replace(', ', ', #') if nums else ''


# ---------------------------------------------------------------- slides
def make_slides(path, person, report, secs, people, concerns, flags, review=()):
    """A .pptx summary for checking the COA: papers by collaboration, co-authors by group, items to check."""
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
    except ImportError:
        sys.exit('--slides needs python-pptx: pip install python-pptx')
    NAVY, GREY, RED, LIGHT = RGBColor(0x1E, 0x27, 0x61), RGBColor(0x55, 0x5B, 0x6E), RGBColor(0xB0, 0x1E, 0x2D), \
        RGBColor(0xEE, 0xF1, 0xF8)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    def text(slide, x, y, w, h, lines, size=12, color=GREY, bold=False):
        tf = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
        tf.word_wrap = True
        for i, ln in enumerate(lines if isinstance(lines, list) else [lines]):
            par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            par.text = ln
            par.font.size, par.font.color.rgb, par.font.bold, par.font.name = Pt(size), color, bold, 'Calibri'
        return tf

    def new(title, sub=''):
        sl = prs.slides.add_slide(blank)
        text(sl, 0.5, 0.3, 12.3, 0.7, title, 26, NAVY, True)
        if sub:
            text(sl, 0.5, 0.95, 12.3, 0.4, sub, 13, GREY)
        return sl

    def table(sl, header, rows, widths, y=1.45, size=10):
        shp = sl.shapes.add_table(len(rows) + 1, len(header), Inches(0.5), Inches(y),
                                  Inches(sum(widths)), Inches(0.3 * (len(rows) + 1)))
        t = shp.table
        for j, w in enumerate(widths):
            t.columns[j].width = Inches(w)
        for i, row in enumerate([header] + rows):
            for j, v in enumerate(row):
                c = t.cell(i, j)
                c.text = str(v)
                c.margin_top = c.margin_bottom = Inches(0.02)
                for par in c.text_frame.paragraphs:
                    par.font.size, par.font.name = Pt(size), 'Calibri'
                    par.font.bold = i == 0
                    par.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if i == 0 else RGBColor(0x22, 0x22, 0x22)
                c.fill.solid()
                c.fill.fore_color.rgb = NAVY if i == 0 else (LIGHT if i % 2 else RGBColor(0xFF, 0xFF, 0xFF))
        return t

    def paged(title, sub, header, rows, widths, per=16, size=10):
        if not rows:
            return
        n = (len(rows) + per - 1) // per
        for k in range(n):
            table(new(title + (f' ({k + 1}/{n})' if n > 1 else ''), sub), header, rows[k * per:(k + 1) * per],
                  widths, size=size)

    n_used = sum(len(ds) for sec, _, ds in secs if not sec.startswith('Not used'))
    by_group = {}
    for p in people:
        by_group.setdefault(p['group'], []).append(p)
    groups = sorted(by_group, key=group_order)

    # 1. title
    sl = prs.slides.add_slide(blank)
    text(sl, 0.8, 2.2, 11.7, 1.0, f'COA check: {person}', 40, NAVY, True)
    text(sl, 0.8, 3.3, 11.7, 2.5, [
        f'Co-authors from NASA ADS, papers since {report["since"]}  |  generated {report["generated"]}',
        f'{n_used} papers used, '
        f'{len(people)} unique co-authors, {len(report.get("added", []))} added to Table 4, '
        f'{len(report["already_listed"])} already listed',
        f'Query: {report["query"]}',
        f'Countries-of-concern matches: {len(concerns)}   |   items to review: {sum(len(v) for v in flags.values())}'],
        16)
    un = report.get('unmatched_affiliations') or []
    if un:
        text(sl, 0.8, 5.4, 11.7, 1.9,
             [f'Check: {sum(n for _, n in un)} papers were left out because your affiliation on them did not match '
              f'--institution. If these are you, add them and rerun:'] +
             [f'   {n} papers:  {aff[:110]}' for aff, n in un[:4]], 13, RED)

    # 2. papers by collaboration
    rows = []
    for sec, sub, ds in secs:
        na = [len(d.get('author') or []) for d in ds]
        names = {p['name'] for p in people for b in p['papers'] if b in {d['bibcode'] for d in ds}}
        odd = [d['_num'] for d in ds if d.get('_size_note')]
        rows.append([(sub or sec) + (f'  (check {", ".join("#" + str(n) for n in odd)}: unusually large)' if odd else ''),
                     len(ds), short_ranges([d['_num'] for d in ds], 45, '#'),
                     f'{min(na)}' if min(na) == max(na) else f'{min(na)}-{max(na)}',
                     '' if sec.startswith('Not used') else len(names)])
    paged('Papers used, by collaboration',
          'Author-list length: fewest-most authors on one paper in that row. Unique co-authors: different people '
          '(not you) on those papers together, counted once.',
          ['Collaboration', 'Papers', 'Paper #', 'Author-list length', 'Unique co-authors'], rows,
          [4.4, 0.9, 3.6, 1.8, 1.6], per=14, size=11)

    # 3. co-authors by group
    rows = [[g, len(by_group[g]), short_ranges([p.get('row') for p in by_group[g]], 40),
             short_ranges(sorted({n for p in by_group[g] for n in p['paper_nums']}), 45, '#')] for g in groups]
    paged('Co-authors, by group', 'Each co-author is placed in one group from the papers you share. '
          'Rows are Table 4 rows in the .xlsx (contiguous with --order collab).',
          ['Group', 'Co-authors', 'Table 4 rows', 'From papers'], rows, [4.4, 1.2, 3.4, 3.3], per=14, size=11)

    # 4. countries of concern
    sub = ('Any affiliation on any shared paper naming China (incl. Hong Kong, Macau), Cuba, Iran, North Korea, '
           'Russia or Venezuela. Text match: verify each.')
    if concerns:
        paged(f'Countries of concern: {len(concerns)} co-authors', sub,
              ['Name', 'Country', 'Matching affiliation', 'Row', 'Group', 'Papers'],
              [[h['name'], ', '.join(h['countries']), h['affiliations'][0][:110] +
                (f' (+{len(h["affiliations"]) - 1} more)' if len(h['affiliations']) > 1 else ''),
                h['row'] or '', h['group'][:28], short_ranges(h['papers'], 30, '#')] for h in concerns],
              [2.2, 1.1, 5.2, 0.6, 1.8, 1.4], per=13, size=9)
    else:
        text(new('Countries of concern: none found', sub), 0.5, 1.6, 12, 1, 'No affiliation matched.', 16)

    # 5. questionable
    q = [p for g in groups if g.startswith('Questionable') for p in by_group[g]]
    paged(f'Questionable: {len(q)} co-authors without a clear collaboration',
          'Only on large papers with no collaboration named, or on papers whose collaboration was guessed',
          ['Name', 'Affiliation', 'Row', 'Group', 'Papers'],
          [[p['name'], p['affiliation'][:80], p.get('row') or '', p['group'].split(': ')[1],
            short_ranges(p['paper_nums'], 30, '#')] for p in sorted(q, key=lambda p: fold(p['name']))],
          [2.4, 5.2, 0.7, 2.6, 1.4], per=15, size=9)
    m = flags.get('initials_only_merges', [])
    m = sorted(m, key=lambda x: (x['same_affiliation'], x['row'] or 0))  # different affiliations first
    n_diff = sum(not x['same_affiliation'] for x in m)
    paged(f'Merged spellings: {len(m)} co-authors, one row each ({n_diff} with different affiliations, listed first)',
          'Spellings combined into ONE row because the initials agree. Same affiliation: almost surely one person. '
          'Different: check, and add the other by hand if needed.',
          ['Row', 'Spellings combined into this row', 'Affiliation of each spelling (most recent paper)'],
          [[x['row'] or '', '  +  '.join(x['merged']),
            ('same: ' + next((a for a in x['affiliations'].values() if a), '(none)')[:85]) if x['same_affiliation'] else
            '  |  '.join(f'{v.split(",")[1].strip() if "," in v else v}: {(a or "(none)")[:45]}'
                         for v, a in x['affiliations'].items())]
           for x in m], [0.7, 3.9, 7.7], per=14, size=9)
    sep = flags.get('same_last_name_and_initial', [])
    paged(f'Kept separate: {len(sep)} sets of people with the same surname and first initial',
          'Each has their OWN row: initials or given names differ, or affiliations differ. '
          'Delete a row only if two are really the same person (e.g. someone who moved).',
          ['Rows', 'Separate people (row: name, affiliation)'],
          [[', '.join(str(x['row'] or '-') for x in g),
            '   |   '.join(f'{x["row"] or "-"}: {x["name"]} ({(x.get("affiliation") or "no affiliation")[:40]})'
                           for x in g)]
           for g in sep], [1.4, 10.9], per=12, size=9)
    dup = flags.get('possible_duplicates_of_sheet', [])
    paged(f'Near-duplicates of names already on the sheet: {len(dup)}',
          'Added as a new row, but close to an existing entry. Delete one if they are the same person.',
          ['Existing row', 'Already on sheet', 'Added from ADS'],
          [[x['sheet_row'], x['sheet'], x['ads']] for x in dup], [1.4, 5.4, 5.5], per=16, size=10)
    na = flags.get('no_affiliation', [])
    paged(f'No affiliation in ADS: {len(na)} co-authors', 'Fill in the affiliation by hand',
          ['Row', 'Name'], [[x['row'] or '', x['name']] for x in na], [1.0, 11.3], per=16, size=10)

    # 5b. papers sent to review: why, and which of their co-authors are (not) in the spreadsheet
    if review:
        paged(f'To review: {len(review)} papers not used',
              'Not confirmed as yours, so not used. Their co-authors are in the spreadsheet only if they are also on a '
              'used paper. If a paper is yours, rerun with --include BIBCODE.',
              ['#', 'Title', 'Authors', 'Your affiliation on it', 'Why not used', 'Not in sheet'],
              [[d['_num'], d['_title'][:90], len(d.get('author') or []), (d.get('_slot_aff') or 'none listed')[:60],
                d['_why'][:150], len(d['_out'])] for d in review],
              [0.5, 4.0, 0.8, 2.6, 3.4, 1.0], per=10, size=9)
        per = 2 * 24
        for d in review:
            out = sorted(d['_out'], key=lambda x: fold(x[1]))
            inn = sorted(d['_in'], key=lambda x: (x[0] or 0, fold(x[1])))
            show_in = len(inn) <= per  # a long list of people already in the sheet is shown as a count
            n = max(1, (max(len(out), len(inn) if show_in else 0) + per - 1) // per)
            for k in range(n):
                sl = new(f'To review #{d["_num"]}: co-authors' + (f' ({k + 1}/{n})' if n > 1 else ''),
                         f'{d["_title"][:120]}  |  {d["bibcode"]}  |  {len(d.get("author") or [])} authors')
                for x0, head, items, fmt, col in (
                        (0.5, f'NOT in spreadsheet ({len(out)}): added only if you --include this paper', out,
                         lambda r, a: a, RED),
                        (6.9, f'Already in spreadsheet ({len(inn)}): from used papers, with Table 4 row', inn,
                         lambda r, a: f'{r or "-"}  {a}', NAVY)):
                    text(sl, x0, 1.4, 6.1, 0.4, head, 12, col, True)
                    if items is inn and not show_in:
                        if k == 0:
                            text(sl, x0, 1.85, 6, 0.8, f'{len(inn)} people, all already in the spreadsheet from used '
                                 f'papers (list in the .json under review_papers)', 10)
                        continue
                    chunk = items[k * per:(k + 1) * per]
                    for c in range(2):
                        part = chunk[c * 24:(c + 1) * 24]
                        if part:
                            text(sl, x0 + c * 3.05, 1.85, 3.0, 5.4, [fmt(r, a) for r, a in part], 9,
                                 RGBColor(0x22, 0x22, 0x22))
                    if not items and k == 0:
                        text(sl, x0, 1.85, 6, 0.4, 'none', 10)

    # 6. paper list, one run of slides per section (joint papers share slides, labelled by combination)
    merged = {}
    for sec, sub, ds in secs:
        merged.setdefault(sec, []).extend((sub, d) for d in ds)
    def names(lst, k=8, rows=False):
        lst = sorted(lst, key=lambda x: fold(x[1]))
        txt = '; '.join((f'{a} (row {r})' if rows and r else a) for r, a in lst[:k])
        return (txt + f'; +{len(lst) - k} more' if len(lst) > k else txt) or 'none'
    for sec, items in merged.items():
        if sec.startswith('Not used'):
            paged(sec, f'{len(items)} papers, not used: their co-authors are in the spreadsheet only if they are also '
                       f'on a used paper. --include BIBCODE adds the "not in spreadsheet" people.',
                  ['#', 'Date', 'Title / bibcode', 'Authors', 'NOT in spreadsheet', 'Already in spreadsheet'],
                  [[d['_num'], d['pubdate'][:7], d['_title'][:80] + '\n' + d['bibcode'], len(d.get('author') or []),
                    names(d.get('_out', [])), names(d.get('_in', []), 6, rows=True)] for _, d in items],
                  [0.5, 0.8, 4.2, 0.8, 3.0, 3.0], per=7, size=9)
            continue
        paged(sec, f'{len(items)} papers', ['#', 'Date', 'Title', 'Journal', 'Authors', 'Bibcode / notes'],
              [[d['_num'], d['pubdate'][:7], (f'[{sub}] ' if sub else '') + d['_title'][:120],
                (d.get('pub') or '')[:38], len(d.get('author') or []),
                d['bibcode'] + ('  ' + '; '.join(d['_flags']) if d['_flags'] else '')]
               for sub, d in items], [0.5, 0.8, 5.6, 2.3, 0.8, 2.3], per=12, size=9)

    # 7. co-author names by group
    per_col, ncol = 26, 4
    for g in groups:
        ps = sorted(by_group[g], key=lambda p: (p.get('row') or 0, fold(p['name'])))
        n = (len(ps) + per_col * ncol - 1) // (per_col * ncol)
        for k in range(n):
            chunk = ps[k * per_col * ncol:(k + 1) * per_col * ncol]
            sl = new(f'{g} ({len(ps)})' + (f' ({k + 1}/{n})' if n > 1 else ''),
                     f'Table 4 rows {short_ranges([p.get("row") for p in ps], 110)}'
                     + ('' if len(short_ranges([p.get("row") for p in ps], 10**6)) <= 110 else
                        '  (rerun with --order collab to make each group one block of rows)'))
            for c in range(ncol):
                col = chunk[c * per_col:(c + 1) * per_col]
                if col:
                    text(sl, 0.5 + c * 3.1, 1.45, 3.05, 5.9,
                         [f'{p.get("row") or "-"}  {p["name"]}' for p in col], 9, RGBColor(0x22, 0x22, 0x22))
    prs.save(path)
    return len(prs.slides)


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


def fill_template(template, out, people, report, person=None, clear_samples=False, by_group=False):
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
        p['row'] = hit['row']
        report['already_listed'].append(dict(sheet=hit['name'], ads=p['name'], row=hit['row'],
                                             spelling_differs=hit['name'] != p['name']))
        # bring a stale Table 4 "Last Active" date forward
        if first <= hit['row'] <= last:
            cur, nd = ws.cell(hit['row'], 5).value, datetime.datetime.strptime(p['last_active'], '%Y-%m')
            if isinstance(cur, datetime.datetime) and cur < nd:
                ws.cell(hit['row'], 5).value = nd
                report['date_updates'].append(dict(name=hit['name'], old=cur.strftime('%m/%d/%Y'),
                                                   new=nd.strftime('%m/%d/%Y')))
    new.sort(key=lambda p: (group_order(p['group']) if by_group else (), fold(p['name'])))

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
    ap.add_argument('--list-papers', action='store_true',
                    help='print a numbered list of the papers used, grouped by collaboration (CTAO, VERITAS, '
                         'IceCube, ...), and write it to NAME_papers.md for checking')
    ap.add_argument('--slides', action='store_true',
                    help='write NAME_check.pptx: papers by collaboration, co-authors by group with their Table 4 '
                         'rows, questionable names and countries-of-concern matches (needs python-pptx)')
    ap.add_argument('--order', choices=('name', 'collab'), default='name',
                    help='order of new Table 4 rows: alphabetical (default) or grouped by collaboration, '
                         'so each group is one block of rows')
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
    classify_papers(found, is_me)
    docs = [d for d in found if d['_status'] == 'accept']
    people = collect_people(docs, is_me, a.max_authors)
    review_docs = [d for d in found if d['_status'] == 'review']
    secs = paper_sections(docs, a.max_authors, review_docs)
    assign_groups(people, docs)

    out_xlsx = a.out or os.path.splitext(a.template)[0] + '_ADS.xlsx'
    if not out_xlsx.lower().endswith('.xlsx'):
        out_xlsx += '.xlsx'
    if os.path.abspath(out_xlsx) == os.path.abspath(a.template):
        sys.exit('--out must differ from the template; the template is never overwritten.')
    out_json = os.path.splitext(out_xlsx)[0] + '.json'
    report = dict(person=person, query=query, since=since, generated=today.isoformat(), max_authors=a.max_authors,
                  papers=[dict(status=d['_status'], why=d['_why'], bibcode=d['bibcode'], date=d['pubdate'][:7],
                               collaboration=' + '.join(d['_collab']) + (' (by author overlap)'
                                                                         if d['_collab_how'] == 'overlap' else ''),
                               n_authors=len(d.get('author') or []), your_affiliation=d['_slot_aff'],
                               title=(d.get('title') or [''])[0], link=f'{UI}/abs/{d["bibcode"]}/abstract')
                          for d in sorted(found, key=lambda d: (('accept', 'review', 'reject').index(d['_status']),
                                                                -d.get('_overlap', 0), d['bibcode']))],
                  already_listed=[], date_updates=[], characters_replaced=[], sample_rows_removed=[], possible_duplicates_of_sheet=[])
    new = fill_template(a.template, out_xlsx, people, report, person, clear_samples=downloaded,
                        by_group=a.order == 'collab')

    by_last = {}
    for p in people:  # same surname and same first initial, kept as separate people
        if p['g']:
            by_last.setdefault((p['lk'], p['g'][0][0]), []).append(p)
    report.update(
        added=[{k: p[k] for k in ('name', 'ads_name', 'affiliation', 'last_active', 'group', 'variants', 'papers',
                                  'paper_nums', 'row')} for p in new],
        review=dict(
            # several spellings merged into ONE row, at least one of them initials only (e.g. "Kim, J." + "Kim, J.-Y.")
            initials_only_merges=[dict(row=p.get('row'), name=p['name'], merged=p['variants'],
                                       affiliations={v: p['variant_affs'].get(v, '') for v in p['variants']},
                                       same_affiliation=same_place(p['variant_affs'].values()))
                                  for p in people
                                  if len(p['variants']) > 1 and
                                  any(all(len(t) == 1 for t in given(split_name(v)[1])) for v in p['variants'])],
            # SEPARATE rows for people with the same surname and first initial (e.g. "Williams, D. A." and "D. R.")
            same_last_name_and_initial=[[dict(row=q.get('row'), name=q['name'], affiliation=q['affiliation']) for q in
                                         sorted(g, key=lambda q: (q.get('row') or 0, fold(q['name'])))]
                                        for g in by_last.values() if len(g) > 1],
            possible_duplicates_of_sheet=report['possible_duplicates_of_sheet'],
            no_affiliation=[dict(row=p.get('row'), name=p['name']) for p in new if not p['affiliation']]))
    review_coauthors(review_docs, people, is_me)
    report['review_papers'] = [dict(number=d['_num'], bibcode=d['bibcode'], title=d['_title'], why=d['_why'],
                                    your_affiliation=d['_slot_aff'],
                                    coauthors_in_sheet=[dict(row=r, name=a) for r, a in d['_in']],
                                    coauthors_not_in_sheet=[a for _, a in d['_out']]) for d in review_docs]
    concerns = concern_matches(people)
    report['countries_of_concern'] = concerns
    if a.slides:
        out_pptx = os.path.splitext(out_xlsx)[0] + '_check.pptx'
        n_sl = make_slides(out_pptx, person, report, secs, people, concerns, report['review'], review_docs)
        print(f'Slides written to {out_pptx} ({n_sl} slides)')
    unmatched = {}
    for d in found:  # your own affiliation on papers rejected only because it did not match --institution
        if d['_status'] != 'accept' and d['_why'].startswith('affiliation does not match'):
            k = clean_aff(d['_slot_aff'])
            unmatched[k] = unmatched.get(k, 0) + 1
    report['unmatched_affiliations'] = sorted(unmatched.items(), key=lambda x: -x[1])
    seen = orcids_seen(found, is_me) if not a.link else []
    report['orcids_on_accepted_papers'] = seen
    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=1, ensure_ascii=False)

    if a.list_papers:
        listing = paper_list(docs, a.max_authors, person, review_docs)
        out_md = os.path.splitext(out_xlsx)[0] + '_papers.md'
        with open(out_md, 'w', encoding='utf-8') as f:
            f.write(listing + '\n')
        print('\n' + listing + '\n')
        print(f'Paper list written to {out_md}')

    n = {s: sum(d['_status'] == s for d in found) for s in ('accept', 'review', 'reject')}
    if not a.link:
        print(f'{len(found)} papers found: {n["accept"]} accepted, {n["review"]} to review, {n["reject"]} rejected')
        if n['review']:
            print('  Check the "review" papers in the .json; add yours with --include BIBCODE ...')
            for d in review_docs:
                new_ = sorted((a_ for _, a_ in d['_out']), key=fold)
                print(f'    #{d["_num"]} {d["bibcode"]} ({len(d.get("author") or [])} authors): '
                      f'{len(d["_in"])} already in spreadsheet; not in spreadsheet: '
                      + ('; '.join(new_[:6]) + (f'; +{len(new_) - 6} more' if len(new_) > 6 else '') if new_ else 'none'))
        if report['unmatched_affiliations']:
            n_un = sum(n for _, n in report['unmatched_affiliations'])
            print(f'  {n_un} papers were left out because your affiliation on them did not match --institution:')
            for aff, k in report['unmatched_affiliations'][:8]:
                print(f'    {k:4d}  {aff[:110]}')
            print('  If any of these are you, add a pattern for it (e.g. --institution "Santa Cruz" "East Bay") and rerun.')
        if seen and not a.orcid:
            print(f'  ORCID on your accepted papers: {seen[0][0]} ({seen[0][1]} papers); '
                  f'if it is yours, rerun with --orcid {seen[0][0]}')
        elif seen and seen[0][0] != a.orcid:
            print(f'  Warning: accepted papers mostly carry ORCID {seen[0][0]}, not {a.orcid}')
    print(f'{len(docs)} papers used, {len(people)} unique co-authors, {len(report["already_listed"])} already listed, '
          f'{len(new)} added, {len(report["date_updates"])} dates updated')
    if concerns:
        print(f'{len(concerns)} co-authors have an affiliation in a listed country of concern; '
              f'see "countries_of_concern" in the .json' + (' and the slides' if a.slides else ''))
    print(f'Wrote {out_xlsx} and {out_json}  (open the .xlsx in Excel and re-save once before uploading)')


if __name__ == '__main__':
    main()
