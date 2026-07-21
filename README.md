# Plorabunt: Attacks on Places of Worship

An interactive map of violence against worshippers, places of worship, and religious
personnel, 1982 to the present.

**Live map:** https://amin-al-maktaba.github.io/plorabunt-map/

---

## What this is

A single self-contained HTML page. It carries the dataset inside it, renders a world map
with D3, and lets a reader filter by victims' religion, mode of attack, year range, and free
text, size the points by incidents or by casualties, and open individual records with their
sources.

There is no server, no database, and no login. That is deliberate: it makes the map cheap to
host, trivial to archive, and easy for other institutions to mirror or embed.

## Repository contents

```
index.html                        the map interface (about 34 KB)
data.js                           the dataset, one incident per line
data/
  Plorabunt_cleaned_full-work_FINAL-2.xlsx   working dataset
  Plorabunt_dataset.csv                      flat export
vendor/                           local copies of the libraries (see Setup step 2)
.github/ISSUE_TEMPLATE/           submission form for proposed new cases
```

## Setup

### 1. Point the submission button at this repository

Open `index.html`, find this line near the top of the main script, and replace the
placeholder with your own `username/repository`:

```js
const GITHUB_REPO = 'Amin-Al-Maktaba/plorabunt-map';
```

Do the same in `.github/ISSUE_TEMPLATE/config.yml`, and in the live-map URL at the top of
this file.

### 2. Vendor the external libraries (recommended)

As shipped, the page loads four things from public CDNs. It works, but it will break the day
any of them changes, which matters for something you intend to cite. Download these four
files into `vendor/`:

| File | Source |
| --- | --- |
| `d3.min.js` | https://unpkg.com/d3@7.9.0/dist/d3.min.js |
| `topojson-client.min.js` | https://unpkg.com/topojson-client@3.1.0/dist/topojson-client.min.js |
| `countries-110m.json` | https://cdn.jsdelivr.net/npm/world-atlas@2.0.2/countries-110m.json |
| `PT_Serif.woff2` | linked from https://fonts.googleapis.com/css2?family=PT+Serif |

Then in `index.html` replace the three CDN URLs with `vendor/d3.min.js`,
`vendor/topojson-client.min.js`, and `vendor/countries-110m.json`, and either self-host the
font or delete the Google Fonts link (the page falls back to Georgia, which is close enough).

### 3. Publish

1. Create a new repository on GitHub, public.
2. Upload the contents of this folder (drag and drop into the web interface is fine, no git
   required).
3. Go to **Settings → Pages**.
4. Under *Source*, choose **Deploy from a branch**; branch `main`, folder `/ (root)`.
5. Save. After a minute or two the map is live at
   `https://amin-al-maktaba.github.io/plorabunt-map/`.

## Updating the map

Everything lives in `index.html`, so any change follows the same path: edit the file, upload
the new version, wait about a minute, hard-refresh the browser.

**Visual and interface changes** (colours, layout, labels, new filters) touch only the
`<style>` block and the script at the bottom of `index.html`. Nothing else in the repository
needs to change, and the published URL stays the same, so anyone who has linked to or
embedded the map keeps seeing the current version automatically.

**Data changes** go into `data.js`, which holds one incident per line. To add a record,
open `data.js` on GitHub, click the pencil, scroll to the bottom, add a comma to the current
last line, and paste your new line above the closing `];`. Commit. That is the whole
procedure.

A record looks like this:

```json
{"d": "20220605", "y": 2022, "ev": "Owo church attack", "city": "Owo", "country": "Nigeria",
 "cc": "Nigeria", "rel": "Christianity", "relRaw": "Christianity", "worship": "Christian",
 "mode": "Armed Assault", "vic": "40-50", "inj": "80", "grp": "ISWAP (alleged)",
 "src": "https://…", "modeB": "Armed Assault", "vicN": 40}
```

`vic` and `inj` are free text so that reported ranges and uncertainty survive intact; `vicN`
is the numeric floor used for sizing and totals. `cc` must be spelled exactly as other
records spell that country, or the dot will not be placed. `modeB` must be one of the five
grouped modes: Bombing/Explosion, Armed Assault, Assassination, Hostage Taking, Other.

Keep `data/` in step with `data.js`, and tag a release whenever the dataset changes so that
citations stay stable.

## Workflow for an accepted submission

1. Verify the source in the issue independently.
2. Normalise the fields: country spelling, mode mapped to one of the five groups, casualty
   range preserved in `vic` with its floor in `vicN`.
3. Add the line to `data.js`, and write `Closes #12` in the commit message so the issue is
   closed automatically and the record is permanently linked to its review.
4. If the case is rejected, close the issue with a short reason. The refusal is part of the
   documentation.

## Handling submissions

The **Submit a case** button opens a pre-filled report in this repository's Issues. Nothing
is written to the map automatically. The curators verify the source, decide whether the
incident falls within scope, add the record to `index.html`, and close the issue. The full
review history stays publicly visible, which is a feature rather than an overhead: it
documents why each record was accepted.

## Embedding in another institution's website

Yes, and this is the easy case. Because the map is a static page, an institution can adopt it
in three ways.

**Iframe.** Anyone can embed the live map with one line, and it stays in step with your
updates automatically:

```html
<iframe src="https://amin-al-maktaba.github.io/plorabunt-map/"
        width="100%" height="800" style="border:0"
        title="Plorabunt: Attacks on Places of Worship"
        loading="lazy"></iframe>
```

GitHub Pages sets no framing restrictions, so this works out of the box. The layout needs
roughly 1000 px of width for the sidebar and map to sit comfortably side by side.

**Self-hosting a copy.** They drop `index.html`, `data.js` and `vendor/` onto their own web server.
It runs from any directory, needs no build step and no server-side language. The trade-off is
that their copy freezes at the version they took, so agree in advance who re-syncs it and
when.

**Deeper integration.** If they want the map inside their own page shell rather than in a
frame, the sidebar and map are separable, but at that point it becomes a small development
job on their side rather than a copy-and-paste.

Whichever route, settle the licence and the required citation first, and make sure the
funding acknowledgement travels with the map.

## Licence and citation

TO BE COMPLETED before publication. Decide separately for the code and for the data; CC BY
4.0 for the dataset and MIT for the code is a common and workable pairing. Add a `CITATION.cff`
file so that GitHub renders a "Cite this repository" button, and mint a DOI by linking the
repository to Zenodo, which archives a snapshot on every release.

## Acknowledgements

TO BE COMPLETED. ITSERR / RESILIENCE funding statement and institutional affiliations.

## Importing an updated workbook

`tools/update_data.py` reads the Excel workbook and updates `data.js`.

```bash
pip install openpyxl
python3 tools/update_data.py data/Plorabunt_cleaned_full-work_FINAL-2.xlsx --dry-run
python3 tools/update_data.py data/Plorabunt_cleaned_full-work_FINAL-2.xlsx
```

Always run `--dry-run` first: it reports what would change and writes nothing.

By default the script **merges**. It appends only the incidents not already present and
leaves every existing line untouched, so corrections made directly in `data.js` survive.
`--rebuild` regenerates the file from scratch and discards those corrections; use it only
when the workbook is definitive. A backup is written to `data.js.backup` before any change.

The script also holds back rows that look like near-duplicates of records already present,
that is same date, same country, and near-identical wording. In practice this catches the
case where the workbook still carries a typo that was corrected in `data.js`. Fix the
workbook and re-run, or pass `--force` to import them regardless.

Normalisation happens in three tables at the top of the script: `COUNTRY_TO_MAP_NAME` maps
workbook country strings to the spellings the map can place, `RELIGION_BUCKETS` collapses
the raw religion field, and `MODE_RULES` maps free-text attack descriptions onto the five
grouped modes. When the script meets a country it cannot place, it says so and names it.
