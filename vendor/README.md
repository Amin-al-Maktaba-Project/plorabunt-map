# vendor/

Local copies of everything the map would otherwise fetch from someone else's
server. With these files present the page has no external dependencies at all:
it runs offline, it can be archived, and it cannot be broken by a CDN changing
or disappearing.

## Files expected here

| File | Download from |
| --- | --- |
| `d3.min.js` | https://unpkg.com/d3@7.9.0/dist/d3.min.js |
| `topojson-client.min.js` | https://unpkg.com/topojson-client@3.1.0/dist/topojson-client.min.js |
| `countries-110m.json` | https://cdn.jsdelivr.net/npm/world-atlas@2.0.2/countries-110m.json |
| `xlsx.full.min.js` | https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js |
| `PTSerif-Regular.ttf`, `PTSerif-Bold.ttf`, `PTSerif-Italic.ttf`, `PTSerif-BoldItalic.ttf` | https://fonts.google.com/specimen/PT+Serif (Download family) |
| `OFL.txt` | ships inside the same download |

`.woff2` versions are used in preference to `.ttf` if present.

PT Serif is published under the SIL Open Font Licence, which requires the
licence text to travel with the font wherever it is redistributed. Keeping
`OFL.txt` here is not optional tidiness.

## What happens if a file is missing

| Missing | Effect |
| --- | --- |
| `d3.min.js` | the map does not render at all |
| `topojson-client.min.js` | the map does not render at all |
| `countries-110m.json` | falls back to the CDN, with a console warning |
| `xlsx.full.min.js` | falls back to the CDN when the curator panel reads an .xlsx |
| the fonts | headings fall back to Georgia |

D3 and TopoJSON have no fallback on purpose: silently reaching out to a third
party is exactly what vendoring is meant to prevent. Upload the vendor files
before uploading an `index.html` that points at them.

## Versions

Pinned deliberately. Upgrading means downloading the new file and testing, not
editing a version number in a URL. That is the trade-off vendoring asks for, and
it is the right one for something meant to stay citable.
