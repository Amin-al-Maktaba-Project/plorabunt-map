# Assistant for administrators

Drafts new Plorabunt cases from news sources. It never touches `data.js`.
A person checks every row before anything can be imported.

## One-off setup

    pip install openpyxl requests
    export ANTHROPIC_API_KEY=...        # cloud model; the key is never stored in a file
    # or, for a local model:  ollama pull qwen2.5:14b   and use  --backend ollama
    python3 tools/assistant.py check

Edit `tools/assistant_config.json` to change the model. Add `drafts/` to `.gitignore`.

## A round of work

1. Draft from sources you choose, or let it search for reports newer than the latest case:

       python3 tools/assistant.py ingest --url https://...
       python3 tools/assistant.py ingest --file urls.txt
       python3 tools/assistant.py ingest --text saved_article.txt --source-url https://...   # paywalled pages
       python3 tools/assistant.py search --keyword "Kano mosque"
       python3 tools/assistant.py search -k "Kano mosque" -k "Maiduguri church" --since 20250101
       python3 tools/assistant.py search --since 20260601 --max 30      # anything new, no keyword

   With a keyword, the assistant first proposes search phrasings (other languages,
   transliterations, synonyms) and shows them. It then lists the reports it found, marks those
   already seen, and asks which to read ("1,3-5", "all", or Enter for none). Add --list-only to
   browse without reading, --no-expand to use your keyword exactly as typed, --yes to read all
   new reports. Search provider: "web" (the cloud model's web search, any date, used
   automatically when ANTHROPIC_API_KEY is set) or "gdelt" (free, recent news only, last three months).

2. Open `drafts/draft_<time>.xlsx`. Read the FLAGS column first, then check each row against
   its source. Correct cells. Set REVIEW STATUS to APPROVED or REJECTED (add a note if rejected).
   The sheet "Not drafted" lists sources with no qualifying incident, and why.

3. Export what you approved:

       python3 tools/assistant.py approve drafts/draft_<time>.xlsx --reviewer "Your name"

4. Import as usual:

       python3 tools/update_data.py drafts/approved_<time>.xlsx --dry-run
       python3 tools/update_data.py drafts/approved_<time>.xlsx

## Safeguards

- Human gate: only rows marked APPROVED by a named reviewer are exported; PENDING rows never are.
- Grounding: each value must come with a verbatim quote found in the article. Unsupported values
  become "Unknown" and are flagged (DATE and COUNTRY are kept but highlighted). Counts must
  match the numbers in their quote.
- Abstention: opinion pieces, foiled plots and unclear cases produce no record, with a reason.
- Duplicates: checked against `data.js` by source URL, then date, country, city and title.
- Copyright: only the URL and short quotes (at most 25 words) are stored, never article text.
- Untrusted text: article text is treated as data; the model's output is validated, and the
  script cannot write to `data.js`.
- Provenance: approved rows carry "[Drafted with AI assistance (model); checked by NAME, date]"
  in ADDITIONAL NOTES (switch off with `"mark_ai_assisted": false`).
- Log: `drafts/assistant_log.jsonl` records every draft, abstention, approval and rejection.
