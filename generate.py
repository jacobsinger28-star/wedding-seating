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


def parse_seating_chart(path: str) -> dict[str, int]:
    df = pd.read_excel(path, sheet_name="Seating Chart", header=None)

    # Find the row that contains "TABLE 1"
    header_row = None
    for i, row in df.iterrows():
        if any(str(v).strip().upper() == "TABLE 1" for v in row if pd.notna(v)):
            header_row = i
            break
    if header_row is None:
        sys.exit("Could not find table header row in Excel file.")

    # Map column index → table number
    col_to_table: dict[int, int] = {}
    for col_idx, val in enumerate(df.iloc[header_row]):
        if pd.isna(val):
            continue
        m = re.match(r"TABLE\s*(\d+)", str(val).strip(), re.IGNORECASE)
        if m:
            col_to_table[col_idx] = int(m.group(1))

    # Read guest rows below the header until we hit a non-name row
    guest_to_table: dict[str, int] = {}
    row = header_row + 1
    while row < len(df):
        all_non_name = True
        for col_idx, table_num in col_to_table.items():
            val = df.iat[row, col_idx]
            if pd.isna(val):
                continue
            s = str(val).strip()
            # Stop when we hit summary rows (pure numbers, status symbols, etc.)
            if re.match(r"^\d+$", s) or s.startswith("✓") or s.startswith("guests") or s.startswith("seats"):
                continue
            # It's a guest name
            all_non_name = False
            guest_to_table[s] = table_num
        # Stop if the entire row looks like a summary row
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
<title>Jake & Shir's Wedding</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;1,400&family=Jost:wght@300;400&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    background: #f0f4f8;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 0 20px 60px;
    color: #1a2e44;
  }}

  /* ── Hero banner ── */
  .hero {{
    width: 100vw;
    background: linear-gradient(160deg, #0d3b6e 0%, #1a5fa8 55%, #2e86c1 100%);
    text-align: center;
    padding: 52px 24px 44px;
    margin-bottom: 40px;
    position: relative;
    overflow: hidden;
  }}

  .hero::before {{
    content: '';
    position: absolute;
    inset: 0;
    background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='60' height='60'%3E%3Ccircle cx='30' cy='30' r='28' fill='none' stroke='rgba(255,255,255,0.06)' stroke-width='1'/%3E%3C/svg%3E") repeat;
    pointer-events: none;
  }}

  .hero-eyebrow {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.7rem;
    letter-spacing: 0.28em;
    text-transform: uppercase;
    color: #a8d4f5;
    margin-bottom: 14px;
  }}

  .hero h1 {{
    font-family: 'Cormorant Garamond', serif;
    font-weight: 400;
    font-size: clamp(2.4rem, 7vw, 3.6rem);
    color: #ffffff;
    line-height: 1.15;
    letter-spacing: 0.02em;
    margin-bottom: 10px;
  }}

  .hero h1 em {{
    font-style: italic;
    color: #d4eaf7;
  }}

  .hero-sub {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.88rem;
    color: #a8d4f5;
    letter-spacing: 0.12em;
    margin-top: 6px;
  }}

  .hero-divider {{
    width: 48px;
    height: 1px;
    background: rgba(168, 212, 245, 0.5);
    margin: 18px auto 0;
  }}

  /* ── Search area ── */
  .search-section {{
    width: 100%;
    max-width: 500px;
    text-align: center;
    margin-bottom: 10px;
  }}

  .search-label {{
    font-family: 'Cormorant Garamond', serif;
    font-size: 1.15rem;
    font-style: italic;
    color: #2e6da4;
    margin-bottom: 16px;
    display: block;
  }}

  .search-wrap {{
    width: 100%;
    position: relative;
    margin-bottom: 32px;
  }}

  #search {{
    width: 100%;
    padding: 15px 22px;
    font-size: 1rem;
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    border: 1.5px solid #8ab8d8;
    border-radius: 40px;
    background: #fff;
    color: #1a2e44;
    outline: none;
    letter-spacing: 0.03em;
    transition: border-color 0.2s, box-shadow 0.2s;
  }}

  #search:focus {{
    border-color: #1a5fa8;
    box-shadow: 0 0 0 3px rgba(26, 95, 168, 0.12);
  }}

  #search.has-suggestions {{
    border-radius: 22px 22px 0 0;
    border-bottom-color: #dce9f4;
  }}

  #search::placeholder {{ color: #8aabbf; }}

  #suggestions {{
    display: none;
    position: absolute;
    top: 100%;
    left: 0; right: 0;
    background: #fff;
    border: 1.5px solid #1a5fa8;
    border-top: none;
    border-radius: 0 0 22px 22px;
    overflow: hidden;
    z-index: 10;
    box-shadow: 0 8px 24px rgba(13, 59, 110, 0.12);
  }}

  #suggestions.open {{ display: block; }}

  .suggestion {{
    padding: 13px 22px;
    cursor: pointer;
    font-size: 0.97rem;
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    color: #1a2e44;
    transition: background 0.12s, color 0.12s;
    text-align: left;
    letter-spacing: 0.02em;
  }}

  .suggestion:hover,
  .suggestion.active {{
    background: #e8f3fb;
    color: #0d3b6e;
  }}

  .suggestion:last-child {{ border-radius: 0 0 20px 20px; }}

  /* ── Result card ── */
  #result {{
    width: 100%;
    max-width: 500px;
    text-align: center;
  }}

  .card {{
    background: #fff;
    border: 1px solid #c8dff0;
    border-radius: 20px;
    padding: 36px 32px;
    box-shadow: 0 4px 24px rgba(13, 59, 110, 0.09);
    animation: cardBloom 0.5s cubic-bezier(0.22, 1, 0.36, 1) both;
    position: relative;
    overflow: hidden;
  }}

  .card::before {{
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: linear-gradient(90deg, #0d3b6e, #2e86c1, #0d3b6e);
    animation: shimmer 1s ease 0.4s both;
  }}

  .card::after {{
    content: '';
    position: absolute;
    top: 0; left: -100%; right: auto;
    width: 60%;
    height: 4px;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.7), transparent);
    animation: shimmerGloss 1s ease 0.4s both;
  }}

  @keyframes cardBloom {{
    0%   {{ opacity: 0; transform: scale(0.88) translateY(16px); box-shadow: none; }}
    60%  {{ opacity: 1; transform: scale(1.02) translateY(-2px); }}
    100% {{ opacity: 1; transform: scale(1) translateY(0);       box-shadow: 0 4px 24px rgba(13,59,110,0.09); }}
  }}

  @keyframes shimmerGloss {{
    0%   {{ left: -60%; }}
    100% {{ left: 140%; }}
  }}

  .table-number.counting {{ opacity: 0.4; }}
  .table-number {{ transition: opacity 0.1s; }}

  .card-welcome {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.72rem;
    letter-spacing: 0.22em;
    text-transform: uppercase;
    color: #2e86c1;
    margin-bottom: 10px;
  }}

  .card .guest-name {{
    font-family: 'Cormorant Garamond', serif;
    font-size: 1.7rem;
    font-weight: 600;
    color: #0d3b6e;
    margin-bottom: 20px;
    line-height: 1.2;
  }}

  .card .table-label {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.72rem;
    letter-spacing: 0.22em;
    text-transform: uppercase;
    color: #2e86c1;
    margin-bottom: 4px;
  }}

  .card .table-number {{
    font-family: 'Cormorant Garamond', serif;
    font-size: 3.2rem;
    font-weight: 400;
    color: #1a5fa8;
    line-height: 1;
    margin-bottom: 6px;
  }}

  .divider {{
    border: none;
    border-top: 1px solid #dce9f4;
    margin: 22px 0;
  }}

  .card .also-seated-label {{
    font-family: 'Jost', sans-serif;
    font-weight: 300;
    font-size: 0.68rem;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: #6a9dbf;
    margin-bottom: 10px;
  }}

  .card .also-seated {{
    font-family: 'Cormorant Garamond', serif;
    font-size: 1rem;
    color: #2a4d6e;
    line-height: 1.9;
  }}

  .no-result {{
    font-family: 'Cormorant Garamond', serif;
    font-style: italic;
    font-size: 1.05rem;
    color: #6a9dbf;
    padding: 20px 0;
  }}
</style>
</head>
<body>

<div class="hero">
  <div class="hero-eyebrow">You are warmly welcomed to</div>
  <h1>Jake &amp; <em>Shir</em></h1>
  <div class="hero-sub">Katerini &nbsp;&middot;&nbsp; Greece</div>
  <div class="hero-divider"></div>
</div>

<div class="search-section">
  <span class="search-label">Find your seat for the evening</span>
  <div class="search-wrap">
    <input id="search" type="text" placeholder="Begin typing your name&hellip;" autocomplete="off" autocorrect="off" spellcheck="false">
    <div id="suggestions"></div>
  </div>
</div>

<div id="result"></div>

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
  const queryWords = q.split(' ');
  const nameWords = n.split(' ');
  if (queryWords.every(qw => nameWords.some(nw => nw.startsWith(qw)))) return 70;
  if (queryWords.some(qw => nameWords.some(nw => nw.startsWith(qw)))) return 40;
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
  const tablemateStr = tablemates.length
    ? `<hr class="divider">
       <div class="also-seated-label">Joining you at your table</div>
       <div class="also-seated">${{tablemates.join('<br>')}}</div>`
    : '';
  resultEl.innerHTML = `<div class="card">
    <div class="card-welcome">Welcome</div>
    <div class="guest-name">${{name}}</div>
    <div class="table-label">Your table</div>
    <div class="table-number" id="tableNum">—</div>
    ${{tablemateStr}}
  </div>`;

  // Count-up animation to the real table number
  const el = document.getElementById('tableNum');
  const duration = 600;
  const start = performance.now();
  const maxNum = Math.max(12, tableNum);
  function tick(now) {{
    const t = Math.min((now - start) / duration, 1);
    const eased = 1 - Math.pow(1 - t, 3);
    const display = t < 1 ? Math.round(eased * maxNum) || 1 : tableNum;
    el.textContent = display;
    if (t < 1) requestAnimationFrame(tick);
  }}
  setTimeout(() => requestAnimationFrame(tick), 150);
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

    html = HTML_TEMPLATE.format(guests_json=guests_json, tables_json=tables_json)
    Path(OUTPUT_FILE).write_text(html, encoding="utf-8")
    print(f"Generated {OUTPUT_FILE} — {len(guest_to_table)} guests ready.")


if __name__ == "__main__":
    main()
