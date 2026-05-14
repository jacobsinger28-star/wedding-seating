"""
Wedding Seating Chart Generator
Run: python generate.py
Output: index.html

Edit your seating chart in "Seating Chart.xlsx", then re-run this script
and push index.html to GitHub Pages.
"""

import pandas as pd
import json
import re
import sys
from pathlib import Path

EXCEL_FILE = "Seating Chart.xlsx"
OUTPUT_FILE = "index.html"

# Mediterranean tile SVG — encodes # as %23 so it can be embedded in a CSS url()
TILE_SVG = (
    "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100' viewBox='0 0 100 100'>"
    "<rect width='100' height='100' fill='%23ede8d4'/>"
    # Corner triangles
    "<polygon points='0,0 28,0 0,28' fill='%232a4a8a'/>"
    "<polygon points='100,0 72,0 100,28' fill='%232a4a8a'/>"
    "<polygon points='0,100 28,100 0,72' fill='%232a4a8a'/>"
    "<polygon points='100,100 72,100 100,72' fill='%232a4a8a'/>"
    # White circles + dot in corners
    "<circle cx='10' cy='10' r='7' fill='%23ede8d4'/>"
    "<circle cx='90' cy='10' r='7' fill='%23ede8d4'/>"
    "<circle cx='10' cy='90' r='7' fill='%23ede8d4'/>"
    "<circle cx='90' cy='90' r='7' fill='%23ede8d4'/>"
    "<circle cx='10' cy='10' r='3' fill='%232a4a8a'/>"
    "<circle cx='90' cy='10' r='3' fill='%232a4a8a'/>"
    "<circle cx='10' cy='90' r='3' fill='%232a4a8a'/>"
    "<circle cx='90' cy='90' r='3' fill='%232a4a8a'/>"
    # Outer octagonal frame
    "<polygon points='28,1 72,1 99,28 99,72 72,99 28,99 1,72 1,28' "
    "fill='none' stroke='%232a4a8a' stroke-width='2.5'/>"
    # Inner octagonal frame
    "<polygon points='34,10 66,10 90,34 90,66 66,90 34,90 10,66 10,34' "
    "fill='none' stroke='%232a4a8a' stroke-width='1'/>"
    # Four petals (cross/flower)
    "<ellipse cx='50' cy='30' rx='7' ry='13' fill='%232a4a8a'/>"
    "<ellipse cx='50' cy='70' rx='7' ry='13' fill='%232a4a8a'/>"
    "<ellipse cx='30' cy='50' rx='13' ry='7' fill='%232a4a8a'/>"
    "<ellipse cx='70' cy='50' rx='13' ry='7' fill='%232a4a8a'/>"
    # Center medallion
    "<circle cx='50' cy='50' r='16' fill='%232a4a8a'/>"
    "<circle cx='50' cy='50' r='10' fill='%23ede8d4'/>"
    "<circle cx='50' cy='50' r='5' fill='%232a4a8a'/>"
    # Diagonal accent lines in corners
    "<line x1='34' y1='34' x2='42' y2='42' stroke='%232a4a8a' stroke-width='1.5'/>"
    "<line x1='66' y1='34' x2='58' y2='42' stroke='%232a4a8a' stroke-width='1.5'/>"
    "<line x1='34' y1='66' x2='42' y2='58' stroke='%232a4a8a' stroke-width='1.5'/>"
    "<line x1='66' y1='66' x2='58' y2='58' stroke='%232a4a8a' stroke-width='1.5'/>"
    "</svg>"
)

TILE_URL = f"data:image/svg+xml,{TILE_SVG}"


def parse_seating_chart(path: str) -> dict[str, int]:
    df = pd.read_excel(path, sheet_name="Seating Chart", header=None)

    header_row = None
    for i, row in df.iterrows():
        if any(str(v).strip().upper() == "TABLE 1" for v in row if pd.notna(v)):
            header_row = i
            break
    if header_row is None:
        sys.exit("Could not find table header row in Excel file.")

    col_to_table: dict[int, int] = {}
    for col_idx, val in enumerate(df.iloc[header_row]):
        if pd.isna(val):
            continue
        m = re.match(r"TABLE\s*(\d+)", str(val).strip(), re.IGNORECASE)
        if m:
            col_to_table[col_idx] = int(m.group(1))

    guest_to_table: dict[str, int] = {}
    row = header_row + 1
    while row < len(df):
        for col_idx, table_num in col_to_table.items():
            val = df.iat[row, col_idx]
            if pd.isna(val):
                continue
            s = str(val).strip()
            if re.match(r"^\d+$", s) or s.startswith("✓") or s.startswith("guests") or s.startswith("seats"):
                continue
            guest_to_table[s] = table_num
        row_vals = [str(df.iat[row, c]).strip() for c in col_to_table if pd.notna(df.iat[row, c])]
        if row_vals and all(re.match(r"^\d+$", v) or v.startswith("✓") or "seats" in v or v == "guests" for v in row_vals):
            break
        row += 1

    return guest_to_table


def build_table_map(guest_to_table: dict[str, int]) -> dict[int, list[str]]:
    tables: dict[int, list[str]] = {}
    for guest, table in guest_to_table.items():
        tables.setdefault(table, []).append(guest)
    return {t: sorted(names) for t, names in sorted(tables.items())}


HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Jake &amp; Shir's Wedding</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,500;0,600;1,400;1,500&family=Jost:wght@300;400&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    background: #f0ebe0;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    color: #2a4a8a;
  }}

  /* ── Tile strips ── */
  .tile-strip {{
    width: 100%;
    background: #ede8d4 url('{tile_url}') repeat-x center;
    background-size: 100px 100px;
    height: 100px;
    flex-shrink: 0;
  }}

  /* ── Main content ── */
  main {{
    flex: 1;
    width: 100%;
    max-width: 480px;
    padding: 36px 28px 44px;
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
  }}

  .bh {{
    font-family: 'Cormorant Garamond', serif;
    font-size: 0.9rem;
    color: #2a4a8a;
    opacity: 0.5;
    margin-bottom: 18px;
    letter-spacing: 0.08em;
  }}

  h1 {{
    font-family: 'Cormorant Garamond', serif;
    font-weight: 400;
    font-size: clamp(3rem, 11vw, 4.4rem);
    line-height: 1.05;
    color: #2a4a8a;
    margin-bottom: 6px;
  }}

  h1 em {{ font-style: italic; }}

  .ornament {{
    color: #2a4a8a;
    opacity: 0.35;
    font-size: 0.7rem;
    letter-spacing: 0.5em;
    margin: 10px 0 12px;
  }}

  .wedding-date {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.68rem;
    letter-spacing: 0.26em;
    text-transform: uppercase;
    color: #2a4a8a;
    margin-bottom: 5px;
  }}

  .wedding-location {{
    font-family: 'Cormorant Garamond', serif;
    font-style: italic;
    font-size: 1rem;
    color: #2a4a8a;
    opacity: 0.75;
    margin-bottom: 32px;
  }}

  .section-rule {{
    width: 100%;
    border: none;
    border-top: 1px solid rgba(42,74,138,0.2);
    margin: 0 0 28px;
  }}

  /* ── Search ── */
  .search-label {{
    font-family: 'Cormorant Garamond', serif;
    font-style: italic;
    font-size: 1.05rem;
    color: #2a4a8a;
    opacity: 0.8;
    margin-bottom: 14px;
    display: block;
  }}

  .search-wrap {{
    width: 100%;
    position: relative;
    margin-bottom: 24px;
  }}

  #search {{
    width: 100%;
    padding: 14px 22px;
    font-size: 0.97rem;
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    border: 1.5px solid rgba(42,74,138,0.35);
    border-radius: 40px;
    background: white;
    color: #2a4a8a;
    outline: none;
    letter-spacing: 0.03em;
    transition: border-color 0.2s, box-shadow 0.2s;
  }}

  #search:focus {{
    border-color: #2a4a8a;
    box-shadow: 0 0 0 3px rgba(42,74,138,0.1);
  }}

  #search.has-suggestions {{
    border-radius: 22px 22px 0 0;
    border-bottom-color: rgba(42,74,138,0.12);
  }}

  #search::placeholder {{ color: rgba(42,74,138,0.3); }}

  #suggestions {{
    display: none;
    position: absolute;
    top: 100%;
    left: 0; right: 0;
    background: white;
    border: 1.5px solid #2a4a8a;
    border-top: none;
    border-radius: 0 0 22px 22px;
    overflow: hidden;
    z-index: 10;
    box-shadow: 0 8px 24px rgba(42,74,138,0.12);
  }}

  #suggestions.open {{ display: block; }}

  .suggestion {{
    padding: 13px 22px;
    cursor: pointer;
    font-family: 'Cormorant Garamond', serif;
    font-size: 1.08rem;
    color: #2a4a8a;
    transition: background 0.1s;
    text-align: left;
  }}

  .suggestion:hover,
  .suggestion.active {{
    background: #eee8d5;
  }}

  .suggestion:last-child {{ border-radius: 0 0 20px 20px; }}

  /* ── Result card ── */
  #result {{ width: 100%; }}

  .card {{
    background: white;
    border: 1.5px solid rgba(42,74,138,0.22);
    border-radius: 4px;
    padding: 36px 28px;
    text-align: center;
    position: relative;
    overflow: hidden;
    animation: cardBloom 0.55s cubic-bezier(0.22, 1, 0.36, 1) both;
  }}

  /* Tile corner accents on the result card */
  .card::before,
  .card::after {{
    content: '';
    position: absolute;
    width: 50px;
    height: 50px;
    background: url('{tile_url}') no-repeat;
    background-size: 50px 50px;
    opacity: 0.55;
  }}
  .card::before {{ top: 0; left: 0; }}
  .card::after  {{ bottom: 0; right: 0; transform: rotate(180deg); }}

  @keyframes cardBloom {{
    0%   {{ opacity: 0; transform: scale(0.91) translateY(14px); }}
    65%  {{ opacity: 1; transform: scale(1.015) translateY(-2px); }}
    100% {{ opacity: 1; transform: scale(1) translateY(0); }}
  }}

  .card-welcome {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.63rem;
    letter-spacing: 0.26em;
    text-transform: uppercase;
    color: #2a4a8a;
    opacity: 0.5;
    margin-bottom: 10px;
  }}

  .guest-name {{
    font-family: 'Cormorant Garamond', serif;
    font-weight: 500;
    font-size: 1.9rem;
    color: #2a4a8a;
    line-height: 1.2;
    margin-bottom: 24px;
  }}

  .table-label {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.63rem;
    letter-spacing: 0.26em;
    text-transform: uppercase;
    color: #2a4a8a;
    opacity: 0.5;
    margin-bottom: 2px;
  }}

  .table-number {{
    font-family: 'Cormorant Garamond', serif;
    font-size: 4.5rem;
    font-weight: 400;
    color: #2a4a8a;
    line-height: 1;
  }}

  .card-divider {{
    border: none;
    border-top: 1px solid rgba(42,74,138,0.15);
    margin: 22px 0;
  }}

  .also-seated-label {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.63rem;
    letter-spacing: 0.26em;
    text-transform: uppercase;
    color: #2a4a8a;
    opacity: 0.5;
    margin-bottom: 12px;
  }}

  .also-seated {{
    font-family: 'Cormorant Garamond', serif;
    font-size: 1rem;
    color: #2a4a8a;
    opacity: 0.85;
    line-height: 1.9;
  }}

  .no-result {{
    font-family: 'Cormorant Garamond', serif;
    font-style: italic;
    font-size: 1rem;
    color: #2a4a8a;
    opacity: 0.55;
    padding: 20px 0;
  }}
</style>
</head>
<body>

<div class="tile-strip"></div>

<main>
  <div class="bh">&#1489;&#x22;&#1492;</div>
  <h1>Jake &amp; <em>Shir</em></h1>
  <div class="ornament">&#10022; &nbsp; &#10022; &nbsp; &#10022;</div>
  <div class="wedding-date">June 28, 2026</div>
  <div class="wedding-location">Katerini, Olympus Riviera &middot; Greece</div>
  <hr class="section-rule">
  <span class="search-label">Find your seat for the evening</span>
  <div class="search-wrap">
    <input id="search" type="text" placeholder="Begin typing your name&hellip;" autocomplete="off" autocorrect="off" spellcheck="false">
    <div id="suggestions"></div>
  </div>
  <div id="result"></div>
</main>

<div class="tile-strip"></div>

<script>
const GUESTS = {guests_json};
const TABLES = {tables_json};

const searchEl = document.getElementById('search');
const suggestionsEl = document.getElementById('suggestions');
const resultEl = document.getElementById('result');

let activeIdx = -1;
let currentHits = [];

function normalize(s) {{
  return s.toLowerCase().replace(/[^a-z0-9]/g, ' ').replace(/\\s+/g, ' ').trim();
}}

function score(name, query) {{
  const n = normalize(name);
  const q = normalize(query);
  if (n === q) return 100;
  if (n.startsWith(q)) return 80;
  if (n.includes(q)) return 60;
  const qw = q.split(' '), nw = n.split(' ');
  if (qw.every(w => nw.some(x => x.startsWith(w)))) return 70;
  if (qw.some(w => nw.some(x => x.startsWith(w)))) return 40;
  return 0;
}}

function openSuggestions(hits) {{
  currentHits = hits;
  activeIdx = -1;
  suggestionsEl.innerHTML = hits.map((h, i) =>
    `<div class="suggestion" data-idx="${{i}}">${{h.name}}</div>`
  ).join('');
  suggestionsEl.classList.add('open');
  searchEl.classList.add('has-suggestions');
}}

function closeSuggestions() {{
  suggestionsEl.classList.remove('open');
  searchEl.classList.remove('has-suggestions');
  activeIdx = -1;
}}

function selectGuest(name) {{
  closeSuggestions();
  searchEl.value = name;
  const tableNum = GUESTS[name];
  const tablemates = (TABLES[tableNum] || []).filter(n => n !== name);
  const matesHTML = tablemates.length
    ? `<hr class="card-divider">
       <div class="also-seated-label">Joining you at your table</div>
       <div class="also-seated">${{tablemates.join('<br>')}}</div>`
    : '';
  resultEl.innerHTML = `<div class="card">
    <div class="card-welcome">Welcome</div>
    <div class="guest-name">${{name}}</div>
    <div class="table-label">Your table</div>
    <div class="table-number" id="tNum">—</div>
    ${{matesHTML}}
  </div>`;

  // Count-up animation
  const el = document.getElementById('tNum');
  const start = performance.now();
  const duration = 700;
  function tick(now) {{
    const t = Math.min((now - start) / duration, 1);
    const eased = 1 - Math.pow(1 - t, 3);
    el.textContent = t < 1 ? (Math.round(eased * Math.max(tableNum, 12)) || 1) : tableNum;
    if (t < 1) requestAnimationFrame(tick);
  }}
  setTimeout(() => requestAnimationFrame(tick), 180);
}}

function setActive(idx) {{
  const items = suggestionsEl.querySelectorAll('.suggestion');
  items.forEach(el => el.classList.remove('active'));
  activeIdx = Math.max(-1, Math.min(idx, items.length - 1));
  if (activeIdx >= 0) items[activeIdx].classList.add('active');
}}

searchEl.addEventListener('input', () => {{
  const q = searchEl.value.trim();
  resultEl.innerHTML = '';
  if (q.length < 2) {{ closeSuggestions(); return; }}
  const hits = Object.entries(GUESTS)
    .map(([name, table]) => ({{ name, table, s: score(name, q) }}))
    .filter(x => x.s > 0)
    .sort((a, b) => b.s - a.s)
    .slice(0, 8);
  if (!hits.length) {{
    closeSuggestions();
    resultEl.innerHTML = '<p class="no-result">We couldn\'t find that name &mdash; please try a different spelling.</p>';
  }} else {{
    openSuggestions(hits);
  }}
}});

searchEl.addEventListener('keydown', e => {{
  if (!suggestionsEl.classList.contains('open')) return;
  if (e.key === 'ArrowDown') {{ e.preventDefault(); setActive(activeIdx + 1); }}
  else if (e.key === 'ArrowUp') {{ e.preventDefault(); setActive(activeIdx - 1); }}
  else if (e.key === 'Enter') {{
    e.preventDefault();
    if (activeIdx >= 0) selectGuest(currentHits[activeIdx].name);
    else if (currentHits.length === 1) selectGuest(currentHits[0].name);
  }}
  else if (e.key === 'Escape') closeSuggestions();
}});

suggestionsEl.addEventListener('mousedown', e => {{
  const item = e.target.closest('.suggestion');
  if (item) {{ e.preventDefault(); selectGuest(currentHits[+item.dataset.idx].name); }}
}});

document.addEventListener('click', e => {{
  if (!e.target.closest('.search-wrap')) closeSuggestions();
}});

searchEl.focus();
</script>
</body>
</html>
"""


def main():
    path = Path(EXCEL_FILE)
    if not path.exists():
        sys.exit(f"File not found: {EXCEL_FILE}\nMake sure '{EXCEL_FILE}' is in the same folder as this script.")

    print(f"Reading {EXCEL_FILE}...")
    guest_to_table = parse_seating_chart(str(path))
    table_map = build_table_map(guest_to_table)

    print(f"Found {len(guest_to_table)} guests across {len(table_map)} tables.")

    guests_json = json.dumps(guest_to_table, ensure_ascii=False, indent=2)
    tables_json = json.dumps({str(k): v for k, v in table_map.items()}, ensure_ascii=False, indent=2)

    html = HTML_TEMPLATE.format(
        guests_json=guests_json,
        tables_json=tables_json,
        tile_url=TILE_URL,
    )
    Path(OUTPUT_FILE).write_text(html, encoding="utf-8")
    print(f"Generated {OUTPUT_FILE} — {len(guest_to_table)} guests ready.")


if __name__ == "__main__":
    main()
