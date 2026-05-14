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
<title>Wedding Seating</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: 'Georgia', serif;
    background: #fdf8f2;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
    color: #3a2c1e;
  }}

  header {{
    text-align: center;
    margin-bottom: 36px;
  }}

  header h1 {{
    font-size: 2rem;
    font-weight: normal;
    letter-spacing: 0.06em;
    color: #6b4c2a;
  }}

  header p {{
    margin-top: 8px;
    font-size: 0.95rem;
    color: #8a6d50;
    font-style: italic;
  }}

  .search-wrap {{
    width: 100%;
    max-width: 480px;
    position: relative;
    margin-bottom: 32px;
  }}

  #search {{
    width: 100%;
    padding: 14px 20px;
    font-size: 1.05rem;
    font-family: inherit;
    border: 1.5px solid #c9a97a;
    border-radius: 40px;
    background: #fff;
    color: #3a2c1e;
    outline: none;
    transition: border-color 0.2s, box-shadow 0.2s;
  }}

  #search:focus {{
    border-color: #9b6f3b;
    box-shadow: 0 0 0 3px rgba(155, 111, 59, 0.15);
  }}

  #search::placeholder {{ color: #b89e82; }}

  #result {{
    width: 100%;
    max-width: 480px;
    text-align: center;
  }}

  .card {{
    background: #fff;
    border: 1px solid #e2cdb0;
    border-radius: 16px;
    padding: 32px 28px;
    margin-bottom: 16px;
    box-shadow: 0 2px 12px rgba(100,70,30,0.07);
  }}

  .card .guest-name {{
    font-size: 1.25rem;
    font-weight: bold;
    color: #6b4c2a;
    margin-bottom: 8px;
  }}

  .card .table-label {{
    font-size: 2.2rem;
    font-weight: bold;
    color: #9b6f3b;
    margin: 8px 0 16px;
    letter-spacing: 0.04em;
  }}

  .card .also-seated {{
    font-size: 0.82rem;
    color: #9a8068;
    font-style: italic;
    line-height: 1.6;
  }}

  .divider {{
    border: none;
    border-top: 1px solid #e2cdb0;
    margin: 14px 0;
  }}

  .multi-hint {{
    font-size: 0.85rem;
    color: #9a8068;
    margin-bottom: 10px;
    font-style: italic;
  }}

  .no-result {{
    color: #9a8068;
    font-style: italic;
    padding: 20px 0;
  }}
</style>
</head>
<body>
<header>
  <h1>&#10022; Wedding Seating &#10022;</h1>
  <p>Type your name to find your table</p>
</header>

<div class="search-wrap">
  <input id="search" type="text" placeholder="Your name&hellip;" autocomplete="off" autocorrect="off" spellcheck="false">
</div>

<div id="result"></div>

<script>
const GUESTS = {guests_json};
const TABLES = {tables_json};

const searchEl = document.getElementById('search');
const resultEl = document.getElementById('result');

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
  const allMatch = queryWords.every(qw => nameWords.some(nw => nw.startsWith(qw)));
  if (allMatch) return 70;
  const anyMatch = queryWords.some(qw => nameWords.some(nw => nw.startsWith(qw)));
  if (anyMatch) return 40;
  return 0;
}}

function render(query) {{
  const q = query.trim();
  if (q.length < 2) {{ resultEl.innerHTML = ''; return; }}

  const hits = Object.entries(GUESTS)
    .map(([name, table]) => ({{ name, table, s: score(name, q) }}))
    .filter(x => x.s > 0)
    .sort((a, b) => b.s - a.s)
    .slice(0, 5);

  if (!hits.length) {{
    resultEl.innerHTML = '<p class="no-result">No guest found &mdash; try a different spelling.</p>';
    return;
  }}

  if (hits.length > 1) {{
    resultEl.innerHTML = '<p class="multi-hint">Multiple matches &mdash; keep typing to narrow down.</p>' +
      hits.map(h => cardHTML(h.name, h.table, false)).join('');
  }} else {{
    resultEl.innerHTML = cardHTML(hits[0].name, hits[0].table, true);
  }}
}}

function cardHTML(name, tableNum, showTablemates) {{
  const tablemates = (TABLES[tableNum] || []).filter(n => n !== name);
  const tablemateStr = showTablemates && tablemates.length
    ? '<hr class="divider"><div class="also-seated">Also at your table:<br>' + tablemates.join(' &bull; ') + '</div>'
    : '';
  return `<div class="card">
    <div class="guest-name">${{name}}</div>
    <div class="table-label">Table ${{tableNum}}</div>
    ${{tablemateStr}}
  </div>`;
}}

searchEl.addEventListener('input', () => render(searchEl.value));
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
