#!/usr/bin/env python3
"""
Update data.js from the Plorabunt Excel workbook.

Usage
-----
    python3 tools/update_data.py Plorabunt_cleaned_full-work_FINAL-2.xlsx

        Default (merge) mode. Reads the workbook, compares it against the
        current data.js, and appends only the incidents that are not already
        there. Existing lines are left untouched, so any correction you made
        by hand after the last import survives.

    python3 tools/update_data.py workbook.xlsx --rebuild

        Regenerates data.js from scratch. Faster to reason about, but it
        DISCARDS every manual correction made in data.js since the last
        import. Use only when the workbook is the definitive version.

    python3 tools/update_data.py workbook.xlsx --dry-run

        Reports what would change and writes nothing.

The script never overwrites data.js without first saving data.js.backup.

Requires: pip install openpyxl
"""

import argparse
import csv
import difflib
import json
import re
import shutil
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is not installed. Run: pip install openpyxl")


# --------------------------------------------------------------------------
# Normalisation rules
#
# These live in ../normalisation.json so that this script and the curator panel
# in index.html apply exactly the same rules. Edit the JSON, never these
# fallbacks, which exist only so the script still runs if the file is missing.
# --------------------------------------------------------------------------

def _load_shared_rules():
    path = Path(__file__).resolve().parent.parent / "normalisation.json"
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except FileNotFoundError:
        print(f"  ! {path.name} not found, falling back to the built-in tables")
        return None


# Spellings the map can place. Anything not listed passes through unchanged;
# if the result is not a name the map knows, the dot is silently dropped, so
# the script warns about every country it has not seen before.
COUNTRY_TO_MAP_NAME = {
    "Ancash, Peru": "Peru",
    "Ayacucho, Peru": "Peru",
    "Huamanga, Peru": "Peru",
    "Lambayeque, Peru": "Peru",
    "Borno State, Nigeria": "Nigeria",
    "Gombe (State), Nigeria": "Nigeria",
    "Kaduna State, Nigeria": "Nigeria",
    "Ondo State, Nigeria": "Nigeria",
    "Punjab, India": "India",
    "Rajasthan, India": "India",
    "Tamil Nadu, India": "India",
    "California, United States": "United States of America",
    "Texas, United States": "United States of America",
    "United States": "United States of America",
    "Chihuahua State, Mexico": "Mexico",
    "Magallanes And Chilean Antarctica, Chile": "Chile",
    "Central Mali": "Mali",
    "Northern Ireland": "United Kingdom",
    "Northern Ireland, United Kingdom": "United Kingdom",
    "Bosnia-Herzegovina": "Bosnia and Herz.",
    "Central African Republic": "Central African Rep.",
    "Democratic Republic of the Congo": "Dem. Rep. Congo",
    "South Sudan": "S. Sudan",
    "Saint Lucia": "__SaintLucia__",
    "Khyber Agency In Federally Administered Tribal Areas (Fata)": "Pakistan",
    "Kurram Agency In Federally Administered Tribal Areas (Fata)": "Pakistan",
}

RELIGION_BUCKETS = ["Islam", "Christianity", "Hinduism", "Buddhism", "Sikhism", "Judaism"]

# Substring tests applied in order to the raw MODE cell, case-insensitive.
MODE_RULES = [
    ("Bombing/Explosion", ["bomb", "explos", "blast", "grenade", "ied", "mortar", "incendiary"]),
    ("Assassination",     ["assassin"]),
    ("Hostage Taking",    ["hostage", "kidnap", "abduct", "barricade", "hijack"]),
    ("Armed Assault",     ["armed assault", "shoot", "gunfire", "firearm", "firing", "fire on",
                           "stab", "killing", "massacre", "ambush", "assault", "attack with",
                           "melee", "beat", "lynch", "arson attack"]),
]


_SHARED = _load_shared_rules()
if _SHARED:
    COUNTRY_TO_MAP_NAME = _SHARED["countryToMapName"]
    RELIGION_BUCKETS = _SHARED["religionBuckets"]
    MODE_RULES = [(r["label"], r["needles"]) for r in _SHARED["modeRules"]]
    DUPLICATE_CUTOFF = _SHARED.get("duplicateCutoff", 0.82)
    OCCASION_RULES = [(r["label"], [re.compile(x, re.I) for x in r["patterns"]])
                      for r in _SHARED.get("occasionRules", [])]
    def _rules(key):
        return [(r["label"], [re.compile(x, re.I) for x in r["patterns"]]) for r in _SHARED.get(key, [])]
    GROUP_RULES, IDEOLOGY_RULES, MOTIVE_RULES = _rules("groupAliases"), _rules("ideologyRules"), _rules("motiveRules")
    GROUP_TYPE_RULES = _rules("groupTypeRules")
    WEAPON_RULES = [(r["label"], bool(r.get("generic")), [re.compile(x, re.I) for x in r["patterns"]])
                    for r in _SHARED.get("weaponRules", [])]
else:
    GROUP_RULES = IDEOLOGY_RULES = MOTIVE_RULES = GROUP_TYPE_RULES = []
    WEAPON_RULES = []
    DUPLICATE_CUTOFF = 0.82
    OCCASION_RULES = []     # without normalisation.json every occasion falls into "Other"


def bucket_religion(raw: str) -> str:
    r = (raw or "").strip()
    if r in RELIGION_BUCKETS:
        return r
    low = r.lower()
    hits = [b for b in RELIGION_BUCKETS if b.lower().rstrip("ity").rstrip("ism") in low]
    if len(hits) == 1:
        return hits[0]
    return "Other/Unknown"


def bucket_mode(raw: str) -> str:
    m = (raw or "").strip().lower()
    if not m:
        return "Other"
    for label, needles in MODE_RULES:
        if any(n in m for n in needles):
            return label
    return "Other"


def bucket_occasion(raw: str) -> str:
    """Group the free-text PARTICULAR OCCASION into a few categories (rules in
    normalisation.json, shared with the curator panel)."""
    t = (raw or "").strip()
    if not t or t.lower().startswith("unknown") or t.lower() in ("no", "none", "n/a"):
        return "Not recorded"
    for label, pats in OCCASION_RULES:
        if any(p.search(t) for p in pats):
            return label
    return "Other"


# --------------------------------------------------------------------------
# Perpetrator fields. Only aggregate, analysable attributes are imported.
# Names of individuals (KILLERS' NAMES) are deliberately never read.
# --------------------------------------------------------------------------

NOT_RECORDED = "Not recorded"
PERP_FIELDS = ("nk", "nkB", "survB", "krel", "grpB", "ideB", "motB", "claimB", "justB")
PERP_DEFAULTS = {"nk": None, "nkB": NOT_RECORDED, "survB": NOT_RECORDED, "krel": NOT_RECORDED,
                 "grpB": NOT_RECORDED, "ideB": NOT_RECORDED, "motB": NOT_RECORDED,
                 "claimB": NOT_RECORDED, "justB": NOT_RECORDED}
# Optional columns. If the workbook has a header with one of these names, the value is
# imported as it stands (it must come from a verified source). They are absent today.
OPTIONAL_COLUMNS = {"KILLERS' AGE BAND": "kage", "KILLERS' GENDER": "kgen",
                    "LOCAL OR FOREIGN ATTACKER": "kloc", "APPEARANCE AND APPROACH": "kapp"}


def _unknown(t: str) -> bool:
    t = (t or "").strip().lower()
    return (not t) or t.startswith("unknown") or t in ("n/a", "none", "no", "-", "—")


def _by_rules(raw: str, rules) -> str:
    t = re.sub(r"\s+", " ", (raw or "").strip())
    if _unknown(t) or re.match(r"^the specific motive (for the attack )?(is|was) unknown\.?$", t, re.I):
        return NOT_RECORDED
    for label, pats in rules:
        if any(p.search(t) for p in pats):
            return label
    return "Other"


def bucket_group(raw: str) -> str:
    t = re.sub(r"\s*\((uncertain|possibly|alleged)[^)]*\)\s*$", "", (raw or "").strip(), flags=re.I)
    if _unknown(t):
        return NOT_RECORDED
    for label, pats in GROUP_RULES:
        if any(p.search(t) for p in pats):
            return label
    return re.sub(r"\s+", " ", t)


def perpetrator_type(group_raw: str, nk) -> str:
    """Kind of perpetrator. A recorded group is typed by fixed rules on its name;
    with no group recorded the type follows from the number of attackers."""
    g = re.sub(r"\s+", " ", (group_raw or "").strip())
    if not _unknown(g):
        for label, pats in GROUP_TYPE_RULES:
            if any(p.search(g) for p in pats):
                return label
        return "Other named group"
    if nk is None:
        return NOT_RECORDED
    return "Lone attacker (no group recorded)" if nk == 1 else "Two or more attackers, no group recorded"


def bucket_weapon(raw: str) -> str:
    """Category of weapon from the free text. Several categories give 'Multiple types'."""
    t = re.sub(r"\s+", " ", (raw or "").strip())
    if _unknown(t) or re.match(r"^unknown( (gun|explosive|weapon) type)?\.?$", t, re.I):
        return NOT_RECORDED
    hits = [(label, generic) for label, generic, pats in WEAPON_RULES if any(p.search(t) for p in pats)]
    if any(not g for _, g in hits):
        hits = [h for h in hits if not h[1]]
    if not hits:
        return "Other"
    return hits[0][0] if len(hits) == 1 else "Multiple types"


def perp_fields(v: dict) -> dict:
    """v holds the raw cell text for the perpetrator columns, keyed by column name."""
    n = re.search(r"\d+", v.get("No. KILLERS", ""))
    nk = int(n.group(0)) if n and not _unknown(v.get("No. KILLERS", "")) else None
    nkB = NOT_RECORDED if nk is None else ("One attacker" if nk == 1 else "Two attackers" if nk == 2 else "Three or more")

    sv = (v.get("No. KILLERS SURVIVED", "") or "").strip().lower()
    if _unknown(sv):
        survB = NOT_RECORDED
    elif re.match(r"^(no|none|died|dead|\d+ dead|killed|all )", sv) or "died" in sv:
        survB = "None survived"
    elif sv.startswith("yes") or re.match(r"^\d+", sv) or re.search(r"surviv|arrest|captur", sv):
        survB = "At least one survived"
    else:
        survB = NOT_RECORDED

    kr = (v.get("KILLERS' RELIGION", "") or "").strip()
    krel = NOT_RECORDED if _unknown(kr) else bucket_religion(re.sub(r"\s*\(uncertain\)", "", kr, flags=re.I))
    if krel == "Other/Unknown":
        krel = "Other"

    cl = (v.get("CLAIM", "") or "").strip().lower()
    if _unknown(cl):
        claimB = NOT_RECORDED
    elif re.search(r"no group claimed|not claimed|no claim|unclaimed", cl):
        claimB = "No group claimed it"
    elif cl.startswith("yes") or "claimed responsibility" in cl or "claimed" in cl:
        claimB = "A group claimed it"
    else:
        claimB = NOT_RECORDED

    cv, tr, se = (v.get("N.KILLERS CONVICTED", "") or "").strip().lower(), (v.get("TRIAL", "") or "").strip().lower(), (v.get("SENTENCE", "") or "").strip().lower()
    if cv.startswith("yes") or re.match(r"^[1-9]\d*\b", cv) or "convict" in cv + tr + se or se.startswith("yes") or re.search(r"imprison|sentenc|execut|death", se):
        justB = "Convicted or sentenced"
    elif re.search(r"arrest|ongoing|in absentia", cv + tr + se) or tr.startswith("yes"):
        justB = "Arrested or on trial"
    elif cv in ("no", "none") or (tr == "no" and se == "no"):
        justB = "No conviction recorded"
    else:
        justB = NOT_RECORDED

    return {"nk": nk, "nkB": nkB, "survB": survB, "krel": krel,
            "grpB": bucket_group(v.get("KILLERS' MILITANT GROUP", "")),
            "ideB": _by_rules(v.get("KILLERS' IDEOLOGY", ""), IDEOLOGY_RULES),
            "motB": _by_rules(v.get("MOTIVE", ""), MOTIVE_RULES),
            "claimB": claimB, "justB": justB}


def first_date(raw) -> str:
    """Rows sometimes carry alternatives such as '19830423/19830425 (Uncertain)'.
    The map needs one date, so the first is taken and the rest is preserved
    only in the workbook."""
    d = str(raw).strip()
    m = re.match(r"\s*(\d{4,8})", d)
    return m.group(1) if m else ""


def year_of(d: str):
    return int(d[:4]) if len(d) >= 4 and d[:4].isdigit() else None


def numeric_floor(raw) -> int:
    """'40-50' -> 40, 'at least 13 (Uncertain)' -> 13, 'Unknown' -> 0."""
    if raw is None:
        return 0
    if isinstance(raw, (int, float)):
        return int(raw)
    nums = re.findall(r"\d+", str(raw))
    return int(nums[0]) if nums else 0


def first_url(raw):
    if not raw:
        return None
    m = re.search(r"https?://\S+", str(raw))
    return m.group(0).rstrip(").,;") if m else None


def clean(v):
    if v is None:
        return ""
    return str(v).strip()


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------

COLUMNS = _SHARED["columns"] if _SHARED else {
    "DATE": 0, "EVENT": 2, "PLACE-BASED WORSHIP": 3, "CITY": 5, "COUNTRY": 6,
    "MODE": 7, "N. INJURED": 9, "PARTICULAR OCCASION": 1, "PLACE OF WORSHIP": 4, "N. VICTIMS": 10, "VICTIMS' RELIGION": 12,
    "KILLERS' MILITANT GROUP": 20, "SOURCES": 31,
}

FIELD_ORDER = ["d", "y", "ev", "city", "country", "cc", "rel", "relRaw", "worship",
               "mode", "vic", "inj", "grp", "src", "modeB", "vicN", "occ", "place", "occB",
               "nk", "nkB", "survB", "krel", "grpB", "ideB", "motB", "claimB", "justB",
               "wpn", "wpnB", "typB", "kage", "kgen", "kloc", "kapp"]


def read_rows(path: Path):
    """Yield the header row and then the data rows, from .xlsx or .csv.

    Excel on macOS exports CSV as semicolon-separated latin-1 by default, so
    both that and the usual comma-separated UTF-8 are accepted."""
    if path.suffix.lower() in (".csv", ".txt"):
        raw = path.read_bytes()
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        first = text.split("\n", 1)[0]
        delimiter = ";" if first.count(";") > first.count(",") else ","
        print(f"  reading CSV, encoding {encoding}, delimiter '{delimiter}'")
        if encoding in ("cp1252", "latin-1"):
            print("    note: this file is not UTF-8. Check that accented characters and")
            print("    non-Latin scripts survived the export before importing.")
        reader = csv.reader(text.splitlines(), delimiter=delimiter, quotechar='"')
        rows = list(reader)
        return rows[0], rows[1:]

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    return next(it), list(it)


def read_workbook(path: Path):
    header, rows = read_rows(path)

    for name, idx in COLUMNS.items():
        actual = clean(header[idx]) if idx < len(header) else ""
        if actual != name:
            print(f"  ! column {idx} is '{actual}', expected '{name}'")
            print("    The workbook layout has changed. Fix COLUMNS in this script "
                  "before trusting the output.")
            sys.exit(1)

    hdr = [clean(h) for h in header]
    optional_idx = {f: hdr.index(name) for name, f in OPTIONAL_COLUMNS.items() if name in hdr}
    if optional_idx:
        print("  optional perpetrator columns found: " + ", ".join(optional_idx))
    records, skipped, unknown_countries = [], [], set()
    for n, row in enumerate(rows, start=2):
        if not row or row[0] in (None, "") or not str(row[0]).strip():
            continue
        if len(row) <= max(COLUMNS.values()):
            row = list(row) + [None] * (max(COLUMNS.values()) + 1 - len(row))
        d = first_date(row[0])
        ev = clean(row[COLUMNS["EVENT"]])
        country = clean(row[COLUMNS["COUNTRY"]])
        if not d or not country:
            skipped.append((n, ev or "(no event)", "missing date or country"))
            continue

        cc = COUNTRY_TO_MAP_NAME.get(country, country)
        if "," in cc or cc.lower() in ("unknown", ""):
            unknown_countries.add(country)

        rel_raw = clean(row[COLUMNS["VICTIMS' RELIGION"]]) or "Unknown"
        mode_raw = clean(row[COLUMNS["MODE"]]) or "Unknown"
        vic_raw = row[COLUMNS["N. VICTIMS"]]
        inj_raw = row[COLUMNS["N. INJURED"]]

        records.append({
            "d": d,
            "y": year_of(d),
            "ev": ev or "Unknown",
            "city": clean(row[COLUMNS["CITY"]]) or "Unknown",
            "country": country,
            "cc": cc,
            "rel": bucket_religion(rel_raw),
            "relRaw": rel_raw,
            "worship": clean(row[COLUMNS["PLACE-BASED WORSHIP"]]) or "Unknown",
            "mode": mode_raw,
            "vic": clean(vic_raw) or "Unknown",
            "inj": clean(inj_raw) or "Unknown",
            "grp": clean(row[COLUMNS["KILLERS' MILITANT GROUP"]]) or "Unknown",
            "src": first_url(row[COLUMNS["SOURCES"]]),
            "modeB": bucket_mode(mode_raw),
            "vicN": numeric_floor(vic_raw),
            "occ": clean(row[COLUMNS["PARTICULAR OCCASION"]]) or "Unknown",
            "place": clean(row[COLUMNS["PLACE OF WORSHIP"]]) or "Unknown",
            "occB": bucket_occasion(clean(row[COLUMNS["PARTICULAR OCCASION"]])),
            "typB": perpetrator_type(clean(row[COLUMNS["KILLERS' MILITANT GROUP"]]), perp_fields({"No. KILLERS": clean(row[COLUMNS["No. KILLERS"]])})["nk"]),
            "wpn": clean(row[COLUMNS["WEAPONS"]]) or "Unknown",
            "wpnB": bucket_weapon(clean(row[COLUMNS["WEAPONS"]])),
            **perp_fields({k: clean(row[COLUMNS[k]]) for k in COLUMNS if k in (
                "No. KILLERS", "No. KILLERS SURVIVED", "CLAIM", "KILLERS' RELIGION", "KILLERS' IDEOLOGY",
                "MOTIVE", "N.KILLERS CONVICTED", "TRIAL", "SENTENCE", "KILLERS' MILITANT GROUP")}),
            **{f: clean(row[i]) for f, i in optional_idx.items() if clean(row[i]) and not _unknown(clean(row[i]))},
        })
    return records, skipped, unknown_countries


def read_existing(path: Path):
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8")
    return json.loads(text[text.index("["):text.rindex("]") + 1])


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------

HEADER = """/* ------------------------------------------------------------------------
   Plorabunt dataset. One incident per line.

   To add a case by hand: copy the last line, paste it above the closing "];",
   and edit the values. Keep the comma at the end of every line except the last.

   To import from the workbook: python3 tools/update_data.py <workbook.xlsx>

   Field notes:
     d       date as YYYYMMDD (string)
     y       year (number)
     cc      country name the map uses to place the dot; must match the
             spelling already used by other records for that country
     modeB   one of: Bombing/Explosion, Armed Assault, Assassination,
             Hostage Taking, Other
     vic     casualties as reported, ranges and uncertainty preserved
     vicN    numeric floor of vic, used for sizing and totals
------------------------------------------------------------------------ */
window.PLORABUNT_DATA = [
"""


def serialise(rec):
    ordered = {k: rec[k] for k in FIELD_ORDER if k in rec}
    ordered.update({k: v for k, v in rec.items() if k not in ordered})
    return json.dumps(ordered, ensure_ascii=False)


TYP_FIELDS = ("typB",)
TYP_DEFAULTS = {"typB": NOT_RECORDED}
WPN_FIELDS = ("wpn", "wpnB")
WPN_DEFAULTS = {"wpn": "Unknown", "wpnB": NOT_RECORDED}
OCC_FIELDS = ("occ", "place", "occB")
OCC_DEFAULTS = {"occ": "Unknown", "place": "Unknown", "occB": "Not recorded"}


def backfill(existing, incoming, fields=OCC_FIELDS, defaults=OCC_DEFAULTS):
    """Add the given fields to records that lack them, by matching each to its
    workbook row. Nothing else in any record is changed. Returns (records, report)."""
    marker = fields[-1] if fields[-1] != "kapp" else fields[0]
    by_key = {}
    by_dc = {}
    for r in incoming:
        by_key.setdefault((r["d"], r["ev"].strip().lower()), []).append(r)
        by_dc.setdefault((r["d"], r["cc"]), []).append(r)
    out, how = [], {"exact": 0, "close": 0, "single": 0, "already": 0}
    unmatched = []
    for rec in existing:
        if marker in rec:
            how["already"] += 1
            out.append(rec)
            continue
        hit = None
        c = by_key.get((rec.get("d"), (rec.get("ev") or "").strip().lower()))
        if c:
            hit, kind = c[0], "exact"
        else:
            sib = by_dc.get((rec.get("d"), rec.get("cc")), [])
            names = [x["ev"] for x in sib]
            m = difflib.get_close_matches(rec.get("ev") or "", names, n=1, cutoff=DUPLICATE_CUTOFF)
            if m:
                hit, kind = sib[names.index(m[0])], "close"
            elif len(sib) == 1:
                hit, kind = sib[0], "single"
        new = dict(rec)
        if hit:
            how[kind] += 1
            for f in fields:
                new[f] = hit.get(f, defaults.get(f))
            for f in OPTIONAL_COLUMNS.values():
                if f in hit and f not in new:
                    new[f] = hit[f]
        else:
            unmatched.append(rec)
            for f in fields:
                new[f] = defaults.get(f)
        out.append(new)
    return out, how, unmatched


def write_data_js(path: Path, records):
    body = ",\n".join(serialise(r) for r in records)
    path.write_text(HEADER + body + "\n];\n", encoding="utf-8")


# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="Update data.js from the Plorabunt workbook (.xlsx or .csv).")
    ap.add_argument("workbook", type=Path, help="path to the .xlsx or .csv export")
    ap.add_argument("--data", type=Path, default=Path(__file__).resolve().parent.parent / "data.js")
    ap.add_argument("--rebuild", action="store_true",
                    help="regenerate from scratch, discarding manual edits in data.js")
    ap.add_argument("--dry-run", action="store_true", help="report only, write nothing")
    ap.add_argument("--backfill", action="store_true",
                    help="add the occasion and place fields to existing records from the "
                         "workbook, changing nothing else")
    ap.add_argument("--force", action="store_true",
                    help="append rows flagged as probable duplicates as well")
    args = ap.parse_args()

    if not args.workbook.exists():
        sys.exit(f"Workbook not found: {args.workbook}")

    print(f"Reading {args.workbook.name}")
    incoming, skipped, unknown_countries = read_workbook(args.workbook)
    print(f"  {len(incoming)} usable rows")

    existing = read_existing(args.data)
    print(f"  {len(existing)} records currently in {args.data.name}")

    if args.backfill:
        final = existing
        for fields, defaults in ((OCC_FIELDS, OCC_DEFAULTS), (PERP_FIELDS, PERP_DEFAULTS), (WPN_FIELDS, WPN_DEFAULTS), (TYP_FIELDS, TYP_DEFAULTS)):
            final, how, unmatched = backfill(final, incoming, fields, defaults)
            print(f"  backfill of {', '.join(fields)}: matched exactly {how['exact']}, close {how['close']}, "
                  f"single {how['single']}, already present {how['already']}, no match {len(unmatched)}")
        print(f"\nMODE: backfill. Matched exactly: {how['exact']}, by close wording: {how['close']}, "
              f"by being the only record that day in that country: {how['single']}, "
              f"already had the fields: {how['already']}, no match: {len(unmatched)}.")
        for r in unmatched[:40]:
            print(f"  no match: {r.get('d')} {r.get('cc')} {str(r.get('ev'))[:70]}")
        if len(unmatched) > 40:
            print(f"  … and {len(unmatched) - 40} more")
        if args.dry_run:
            print("\nDry run: nothing written.")
            return
        if args.data.exists():
            backup = args.data.with_suffix(".js.backup")
            shutil.copy2(args.data, backup)
            print(f"\nPrevious version saved as {backup.name}")
        write_data_js(args.data, final)
        print(f"Wrote {args.data.name} with {len(final)} records.")
        return

    if args.rebuild:
        final = incoming
        added, near = incoming, []
        print("\nMODE: rebuild. Manual edits made in data.js will be lost.")
    else:
        seen = {(r.get("d"), (r.get("ev") or "").strip().lower()) for r in existing}
        candidates = [r for r in incoming
                      if (r["d"], r["ev"].strip().lower()) not in seen]
        added, near = [], []
        by_dc = {}
        for r in existing:
            by_dc.setdefault((r.get("d"), r.get("cc")), []).append(r.get("ev") or "")
        for r in candidates:
            siblings = by_dc.get((r["d"], r["cc"]), [])
            match = difflib.get_close_matches(r["ev"], siblings, n=1, cutoff=DUPLICATE_CUTOFF)
            if match and not args.force:
                near.append((r, match[0]))
            else:
                added.append(r)
        final = existing + added
        print(f"\nMODE: merge. {len(added)} new record(s) to append, "
              f"{len(near)} held back as probable duplicates.")

    if added:
        print("\nNew records:")
        for r in added[:40]:
            print(f"  {r['d']}  {r['cc']:<22} {r['ev'][:60]}")
        if len(added) > 40:
            print(f"  … and {len(added) - 40} more")

    if near:
        print("\n! Held back as probable duplicates of records already present.")
        print("  Same date and country, near-identical wording. Usually this means the")
        print("  workbook still carries a typo that was corrected in data.js.")
        print("  Fix the workbook, or re-run with --force to add them anyway.\n")
        for r, other in near[:40]:
            print(f"  {r['d']} {r['cc']}")
            print(f"      workbook: {r['ev'][:78]}")
            print(f"      data.js : {other[:78]}")
        if len(near) > 40:
            print(f"  … and {len(near) - 40} more")

    if skipped:
        print(f"\nSkipped {len(skipped)} row(s):")
        for n, ev, why in skipped[:20]:
            print(f"  row {n}: {why} — {ev[:60]}")

    if unknown_countries:
        print("\n! Country names the map may not be able to place:")
        for c in sorted(unknown_countries):
            print(f"    {c}")
        print("  Add them to COUNTRY_TO_MAP_NAME in this script, mapping each to the")
        print("  spelling the map already uses.")

    no_source = [r for r in added if not r["src"]]
    if no_source:
        print(f"\n! {len(no_source)} new record(s) carry no source URL. "
              "Consider filling SOURCES in the workbook before publishing.")

    if args.dry_run:
        print("\nDry run: nothing written.")
        return
    if not added and not args.rebuild:
        print("\nNothing to do. data.js is already in step with the workbook.")
        return

    if args.data.exists():
        backup = args.data.with_suffix(".js.backup")
        shutil.copy2(args.data, backup)
        print(f"\nPrevious version saved as {backup.name}")

    write_data_js(args.data, final)
    print(f"Wrote {args.data.name} with {len(final)} records.")
    print("\nNext: open the map locally to check it still renders, then commit "
          "data.js to GitHub.")


if __name__ == "__main__":
    main()
