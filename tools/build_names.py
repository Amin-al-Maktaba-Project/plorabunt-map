#!/usr/bin/env python3
"""Build names.js (victims' and attackers' names) from the workbook.

    python3 tools/build_names.py <workbook.xlsx>

Names are kept in their own file, apart from data.js, so that publishing them
is a separate, deliberate step: delete names.js (and its <script> line) to
publish the site without names. Each entry is keyed "YYYYMMDD|event title" in
lower case, the same key the page builds from a record, and is shown with the
record's own source link.

Names are free text in the workbook. They are split only on semicolons and line
breaks; anything else is kept as given, so no name is cut in the wrong place.
"""
import json, re, sys, difflib
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
import update_data as U

def split_names(text):
    t = U.clean(text)
    if not t or U._unknown(t) or re.match(r"^(never identified|not identified|unidentified|n/?a)\b", t, re.I):
        return []
    parts = [p.strip(" ,.") for p in re.split(r"\s*;\s*|\s*\n\s*", t) if p.strip(" ,.")]
    return parts

def main():
    wb = Path(sys.argv[1])
    out = Path(__file__).resolve().parent.parent / "names.js"
    header, rows = U.read_rows(wb)
    hdr = [U.clean(h) for h in header]
    iv, ip, isv = hdr.index("VICTIMS' NAMES"), hdr.index("KILLERS' NAMES"), hdr.index("No. KILLERS SURVIVED")
    ink = hdr.index("No. KILLERS")
    existing = U.read_existing(Path(__file__).resolve().parent.parent / "data.js")
    by_key = {(r["d"], (r["ev"] or "").strip().lower()): r for r in existing}
    by_dc = {}
    for r in existing:
        by_dc.setdefault((r["d"], r["cc"]), []).append(r)
    names, miss = {}, 0
    for row in rows:
        if not row or not row[0]:
            continue
        row = list(row) + [None] * (len(hdr) - len(row))
        d = U.first_date(row[0]); ev = U.clean(row[U.COLUMNS["EVENT"]]) or "Unknown"
        country = U.clean(row[U.COLUMNS["COUNTRY"]]); cc = U.COUNTRY_TO_MAP_NAME.get(country, country)
        v, p = split_names(row[iv]), split_names(row[ip])
        if not (v or p):
            continue
        rec = by_key.get((d, ev.strip().lower()))
        if not rec:
            sib = by_dc.get((d, cc), [])
            m = difflib.get_close_matches(ev, [x["ev"] for x in sib], n=1, cutoff=U.DUPLICATE_CUTOFF)
            rec = next((x for x in sib if x["ev"] == m[0]), None) if m else (sib[0] if len(sib) == 1 else None)
        if not rec:
            miss += 1
            continue
        surv = U.perp_fields({"No. KILLERS SURVIVED": U.clean(row[isv]) or "", "No. KILLERS": U.clean(row[ink]) or ""})["survB"]
        st = {"None survived": "died in the attack", "At least one survived": "at least one survived"}.get(surv, "")
        entry = {}
        if v: entry["v"] = v
        if p: entry["p"] = p; entry["pst"] = st
        names[f'{rec["d"]}|{(rec["ev"] or "").strip().lower()}'] = entry
    body = ",\n".join(f"{json.dumps(k, ensure_ascii=False)}:{json.dumps(e, ensure_ascii=False)}" for k, e in names.items())
    out.write_text("/* Names, kept apart from data.js. Delete this file and its <script> line to publish without names. */\nwindow.PLORABUNT_NAMES = {\n" + body + "\n};\n", encoding="utf-8")
    nv = sum(len(e.get("v", [])) for e in names.values()); npn = sum(len(e.get("p", [])) for e in names.values())
    print(f"{len(names)} incidents with names; {nv} victim entries, {npn} attacker entries; {miss} rows not matched to a record")

main()
