> **Disclaimer:** This tool was written with the help of Claude (Anthropic's AI assistant, via Claude Code).
> It is provided as is, without warranty of any kind. Use it at your own discretion: check every entry in the
> generated COA spreadsheet against your own records and your funding agency's current rules before submitting.
> The authors and Anthropic are not responsible for errors, omissions or rejected submissions.

# ads_to_coa.py — fill the NSF COA template from NASA ADS

Adds every co-author from your recent papers to **Table 4** ("A:" rows) of the
official NSF Collaborators & Other Affiliations template, and writes a JSON
report of what it did. Your template is never modified; a new copy is written.

You can give it a list of papers (an ADS library or search link), or let it search
all of ADS for you and work out which papers are really yours.

## Requirements

Python 3.8+ and openpyxl (`pip install openpyxl`). No ADS API token is needed.
If you have one, set `ADS_API_TOKEN` (or put it in `~/.ads/dev_key`) and it is used instead.

## Usage

```
# 1. from a list of papers you chose
python3 ads_to_coa.py ADS_LINK [TEMPLATE.xlsx] [--person "Last, First"] [options]

# 2. search all of ADS for one person
python3 ads_to_coa.py [TEMPLATE.xlsx] --person "Last, First M." --orcid ORCID --institution PATTERN ... [options]
```

### Input

- `ADS_LINK`: an ADS library or search URL, e.g.
  `https://ui.adsabs.harvard.edu/public-libraries/<ID>`,
  `https://ui.adsabs.harvard.edu/search/q=docs(library%2F<ID>)&sort=...`,
  or a raw query such as `'author:"^Feng, Qi" year:2022-2026'`.
  A library must be **public** (ADS library → Settings → Make public). Every paper in it is used.
- Without `ADS_LINK`, the script searches all of ADS (see **Search mode** below).
- `TEMPLATE.xlsx`: your NSF COA template, blank or already partly filled. Optional.
  - If you leave it out, the script downloads the current blank template from
    [nsf.gov](https://nsf-gov-resources.nsf.gov/files/coa_template.xlsx?VersionId=uaz022di3.2KKDC7uauIadw7yeI3Xf.1)
    as `coa_template_NSF_blank.xlsx` next to the output, and reuses that copy on later runs.
  - NSF's sample rows ("Alphaman, Alan K", "Test University XYZ", ...) are removed from any template.
    Your name goes into a blank Table 1.

### Options

- `--person`: whose COA this is; that person is left out of the co-author list.
  With your own template it defaults to the "Your Name" cell of Table 1. Give the middle initial if you use one:
  `"Smith, John A."` does not match `Smith, J. R.`.
- `--months`: look-back window, 48 by default (NSF rule for co-authors).
- `--max-authors N`: skip papers with more than N authors. See the large-collaboration note below.
- `--out NAME.xlsx`: output file name. The report is written next to it as `NAME.json`.
  The default is `<template>_ADS.xlsx`, or `coa_<Last>_ADS.xlsx` in the current folder when the template was downloaded.
  The output can't be the template itself.

## Search mode

Searching ADS by name alone doesn't work for common names. For example, there are over 900 "Feng, Qi" papers since 2022
in all fields. So search mode needs at least one of these:

- `--orcid 0000-0001-6674-4238`: any paper with your ORCID on your author slot is accepted.
- `--institution "Utah" "Smithsonian|Harvard" "Barnard|Columbia"`: your institutions during the window, as text or
  regular expressions. Each is checked against **your own** affiliation on each paper, not against your co-authors'.

Other search options:
- `--field`: `physics` (astronomy + physics collections, default), `astronomy`, or `any`.
- `--include BIBCODE ...`: accept papers you confirmed from the review list.
- `--exclude BIBCODE ...`: drop papers.

Each paper found is sorted into one of three groups:

| Status | When                                                                                                                                                                                                                                                                                                                            |
| ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| accept | Your ORCID is on your author slot. Or your affiliation on that paper matches `--institution`. Or, when no affiliation is listed and there are 20+ authors, its author list is at least 60% the same as one accepted paper's (collaboration proceedings, errata, arXiv versions).                                                |
| review | Its author list is 40–60% like an accepted paper's. Or it is a short paper with no affiliation where half the co-authors appear on accepted papers.                                                                                                                                                                             |
| reject | Someone else's ORCID is on that author slot. Or the affiliation is a different institution and co-authors are unrelated. Or the similar accepted paper has a **namesake**: e.g. a proceedings paper listing just "Smith, J." is rejected when the closest accepted paper lists both "Smith, J. A." and "Smith, J. R.". |

Only accepted papers go into the spreadsheet. The `.json` lists every paper with its status, the reason, your affiliation
on it and an ADS link, with the review papers first. If you ran with `--institution` only, the script also prints the ORCID
found on your accepted papers so you can rerun with `--orcid`.

Tested (Nov 2026):
- **Qi Feng** (`--institution` only): 60 papers accepted, including all 36 in a hand-curated library.
  The other 24 were also his (dark matter papers, software releases, CTAO proceedings). 5 went to review; 0 wrong papers were accepted.
- **A second researcher with a common surname** (`--orcid` + `--institution`): 61 accepted and 7 to review.
  162 large-collaboration records listing only the surname and first initial were correctly rejected as a namesake's.

## Examples

```
# search mode, blank NSF template
python3 ads_to_coa.py --person "Smith, John A." --orcid 0000-0000-0000-0000 --institution "Example University|EU" --out coa_Smith_2026.xlsx

# then add papers you confirmed from the review list
python3 ads_to_coa.py --person "Smith, John A." --orcid 0000-0000-0000-0000 --institution "Example University|EU" --include BIBCODE1 BIBCODE2 --out coa_Smith_2026.xlsx

# from your own ADS library, adding to a COA you already started
python3 ads_to_coa.py "https://ui.adsabs.harvard.edu/public-libraries/<LIBRARY_ID>" my_coa.xlsx --out my_coa_updated.xlsx
```

## What it does with the papers

- Lists each unique co-author once.
  Spelling variants such as `Hinton, J.` and `Hinton, J. A.`, or names with and without accents, count as one person.
- Skips anyone already named on the sheet with a compatible name (same last name and matching given names or initials).
  If an existing Table 4 row has an older "Last Active" date, it moves that date forward.
  A near-match such as sheet "Smith, Andrew W." vs ADS "Smith, A. J." is added and flagged.
- Uses the affiliation from each person's most recent paper.
  It keeps only the first affiliation and removes the street address and postcode, as the template requires.
  "Last Active" is the first day of that paper's publication month.
- Fills blank Table 4 rows first, then adds rows.
  Table 5, the tables' ranges and the dropdowns move down with the new rows, and the template's own row styles are kept.
- Converts characters that Research.gov rejects to plain letters (ć→c, ł→l, š→s, ∣→|, dashes→-).
  Latin-1 accents such as é and ü are left alone. This applies to entries already in the template too.

## Before you upload

1. In search mode, look at the `review` papers in the `.json` and rerun with `--include` for yours.
2. Open the `.xlsx` in Excel and save it once, so Excel rewrites its internal metadata.
3. Read the `review` section at the end of the `.json`:
   - `initials_only_merges`: e.g. `Kim, J.` merged with `Kim, J.-Y.`. These might be two different people.
   - `same_last_name_and_initial`: kept as separate people; check none is a duplicate.
   - `possible_duplicates_of_sheet`: added, but close to a name already on the sheet.
   - `no_affiliation`: ADS had no affiliation; fill it in by hand.

ADS only covers what it indexes. Papers in journals outside ADS still have to be added by hand.

## Large collaborations

The general NSF COA FAQ says to list all co-authors from the last 48 months.
The MPS Physics solicitation (NSF 26-523) says that for large collaborations or large author lists, you should list only
people you worked with "in a direct and substantive way." Follow your solicitation and ask the program officer if unsure.
`--max-authors 100` keeps VERITAS-sized papers and drops CTAO/EHT/IceCube-sized ones. Add the people you worked with
directly on those big papers by hand.
