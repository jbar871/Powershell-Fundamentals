#!/usr/bin/env python3
"""
Generate an HTML computer inventory report from the Freshservice export Excel file.
Usage: python3 computer-inventory-report.py <input.xlsx> [output.html]
"""

import sys
import warnings
import html
from datetime import datetime, timezone
from pathlib import Path

warnings.filterwarnings("ignore")

try:
    import openpyxl
except ImportError:
    print("Missing dependency: pip install openpyxl")
    sys.exit(1)

# ── Column indices (0-based) ────────────────────────────────────────────────
COL = {
    "target":          0,   # A
    "asset_name":      4,   # E
    "location":        7,   # H  Used By Location
    "last_login_by":   8,   # I  Last login by
    "email":           9,   # J  Used By Primary Email
    "name":           10,   # K  Used By Name
    "job_title":      11,   # L  Used By Job Title
    "user_active":    12,   # M  Used By Is Active (bool)
    "device_active":  13,   # N  Device Active in AD
    "last_audit":     29,   # AD Last Audit Date
    "last_updated":   30,   # AE Last Updated Date
    "ad_changed":     31,   # AF AD whenChanged
    "ad_logon":       32,   # AG AD LastLogonDate
}

STALE_DAYS = 90  # flag dates older than this


def fmt_date(val):
    if val is None or val == "#N/A":
        return ""
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    return str(val)


def days_ago(val, now):
    if not isinstance(val, datetime):
        return None
    dt = val.replace(tzinfo=None)
    return (now - dt).days


def active_label(val):
    if val is True:
        return "Yes"
    if val is False:
        return "No"
    if val == "#N/A" or val is None:
        return "N/A"
    return str(val)


def row_class(user_active, device_active, last_audit_days, logon_days):
    if user_active is False:
        return "inactive-user"
    if user_active == "#N/A" or user_active is None:
        return "unknown-user"
    if device_active is False:
        return "inactive-device"
    if (last_audit_days is not None and last_audit_days > STALE_DAYS) or \
       (logon_days is not None and logon_days > STALE_DAYS):
        return "stale"
    return ""


def load_rows(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        rows.append(r)
    return rows


def build_html(rows, source_file):
    now = datetime.now()
    generated = now.strftime("%Y-%m-%d %H:%M")

    total = len(rows)
    inactive_user = sum(1 for r in rows if r[COL["user_active"]] is False)
    na_user = sum(1 for r in rows if r[COL["user_active"]] in (None, "#N/A"))
    inactive_device = sum(1 for r in rows if r[COL["device_active"]] is False)
    stale = sum(
        1 for r in rows
        if r[COL["user_active"]] not in (False, None, "#N/A")
        and r[COL["device_active"]] is not False
        and (
            (days_ago(r[COL["last_audit"]], now) or 0) > STALE_DAYS
            or (days_ago(r[COL["ad_logon"]], now) or 0) > STALE_DAYS
        )
    )

    table_rows = []
    for r in rows:
        ua = r[COL["user_active"]]
        da = r[COL["device_active"]]
        audit_days = days_ago(r[COL["last_audit"]], now)
        logon_days = days_ago(r[COL["ad_logon"]], now)
        cls = row_class(ua, da, audit_days, logon_days)

        def d(col):
            return html.escape(fmt_date(r[COL[col]]))

        def s(col):
            v = r[COL[col]]
            if v is None or v == "#N/A":
                return ""
            return html.escape(str(v))

        def stale_cls(days):
            if days is None:
                return ""
            return ' class="stale-cell"' if days > STALE_DAYS else ""

        table_rows.append(f"""
        <tr class="{cls}">
          <td>{s("asset_name")}</td>
          <td>{s("name")}</td>
          <td>{s("email")}</td>
          <td>{s("job_title")}</td>
          <td>{s("location")}</td>
          <td class="bool {'yes' if ua is True else 'no' if ua is False else 'na'}">{active_label(ua)}</td>
          <td class="bool {'yes' if da is True else 'no' if da is False else 'na'}">{active_label(da)}</td>
          <td{stale_cls(audit_days)}>{d("last_audit")}</td>
          <td>{d("last_updated")}</td>
          <td>{d("ad_changed")}</td>
          <td{stale_cls(logon_days)}>{d("ad_logon")}</td>
        </tr>""")

    rows_html = "\n".join(table_rows)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Computer Inventory Report</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
         font-size: 13px; background: #f0f2f5; color: #1a1a2e; }}
  header {{ background: #1a1a2e; color: #fff; padding: 16px 24px;
            display: flex; align-items: center; justify-content: space-between; }}
  header h1 {{ font-size: 18px; font-weight: 600; }}
  header span {{ font-size: 11px; opacity: 0.6; }}

  .summary {{ display: flex; gap: 12px; padding: 16px 24px; flex-wrap: wrap; }}
  .card {{ background: #fff; border-radius: 8px; padding: 12px 18px;
           min-width: 130px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  .card .num {{ font-size: 26px; font-weight: 700; }}
  .card .lbl {{ font-size: 11px; color: #666; margin-top: 2px; }}
  .card.red .num {{ color: #d32f2f; }}
  .card.orange .num {{ color: #e65100; }}
  .card.yellow .num {{ color: #f9a825; }}
  .card.gray .num {{ color: #555; }}

  .toolbar {{ padding: 0 24px 12px; display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }}
  .toolbar input {{ padding: 7px 12px; border: 1px solid #ccc; border-radius: 6px;
                    font-size: 13px; width: 280px; }}
  .toolbar select {{ padding: 7px 10px; border: 1px solid #ccc; border-radius: 6px;
                     font-size: 13px; }}
  .toolbar label {{ font-size: 12px; color: #555; }}

  .wrap {{ padding: 0 24px 32px; overflow-x: auto; }}
  table {{ width: 100%; border-collapse: collapse; background: #fff;
           border-radius: 8px; overflow: hidden;
           box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  th {{ background: #1a1a2e; color: #fff; padding: 10px 10px; text-align: left;
        font-weight: 600; font-size: 12px; white-space: nowrap;
        cursor: pointer; user-select: none; position: sticky; top: 0; z-index: 2; }}
  th:hover {{ background: #2d2d4e; }}
  th .sort-icon {{ margin-left: 4px; opacity: 0.5; }}
  th.asc .sort-icon::after {{ content: " ▲"; }}
  th.desc .sort-icon::after {{ content: " ▼"; }}
  th:not(.asc):not(.desc) .sort-icon::after {{ content: " ⇅"; }}

  td {{ padding: 8px 10px; border-bottom: 1px solid #f0f0f0; vertical-align: middle;
        white-space: nowrap; }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: rgba(0,0,0,.02); }}

  tr.inactive-user td {{ background: #fdecea; }}
  tr.inactive-user:hover td {{ background: #fcd8d5; }}
  tr.unknown-user td {{ background: #fff8e1; }}
  tr.unknown-user:hover td {{ background: #fff0c0; }}
  tr.inactive-device td {{ background: #fff3e0; }}
  tr.stale td {{ background: #f3f8ff; }}

  td.bool {{ text-align: center; font-weight: 600; font-size: 12px; }}
  td.bool.yes {{ color: #2e7d32; }}
  td.bool.no {{ color: #c62828; }}
  td.bool.na {{ color: #888; }}
  td.stale-cell {{ color: #e65100; font-weight: 600; }}

  .legend {{ display: flex; gap: 12px; padding: 0 24px 16px; flex-wrap: wrap; font-size: 11px; }}
  .legend span {{ display: flex; align-items: center; gap: 5px; }}
  .swatch {{ width: 12px; height: 12px; border-radius: 2px; display: inline-block; }}

  #row-count {{ font-size: 12px; color: #555; padding: 0 24px 8px; }}
</style>
</head>
<body>
<header>
  <h1>Computer Inventory Report</h1>
  <span>Source: {html.escape(source_file)} &nbsp;|&nbsp; Generated: {generated}</span>
</header>

<div class="summary">
  <div class="card gray"><div class="num">{total}</div><div class="lbl">Total Devices</div></div>
  <div class="card red"><div class="num">{inactive_user}</div><div class="lbl">Inactive User</div></div>
  <div class="card yellow"><div class="num">{na_user}</div><div class="lbl">User Unknown / N/A</div></div>
  <div class="card orange"><div class="num">{stale}</div><div class="lbl">Stale (&gt;{STALE_DAYS}d)</div></div>
</div>

<div class="toolbar">
  <input type="text" id="search" placeholder="Search asset, user, email, location...">
  <label>User Active:
    <select id="filter-user">
      <option value="">All</option>
      <option value="Yes">Yes</option>
      <option value="No">No</option>
      <option value="N/A">N/A</option>
    </select>
  </label>
  <label>Device Active in AD:
    <select id="filter-device">
      <option value="">All</option>
      <option value="Yes">Yes</option>
      <option value="No">No</option>
    </select>
  </label>
  <label>Show only stale (&gt;{STALE_DAYS}d):
    <input type="checkbox" id="filter-stale">
  </label>
</div>
<div id="row-count"></div>

<div class="legend">
  <span><span class="swatch" style="background:#fdecea"></span> Inactive User</span>
  <span><span class="swatch" style="background:#fff8e1"></span> User Unknown/N/A</span>
  <span><span class="swatch" style="background:#fff3e0"></span> Inactive Device in AD</span>
  <span><span class="swatch" style="background:#f3f8ff"></span> Stale (&gt;{STALE_DAYS}d no audit/logon)</span>
  <span style="color:#e65100; font-weight:600">Orange date = stale</span>
</div>

<div class="wrap">
<table id="tbl">
  <thead>
    <tr>
      <th data-col="0">Asset Name<span class="sort-icon"></span></th>
      <th data-col="1">Used By Name<span class="sort-icon"></span></th>
      <th data-col="2">Email<span class="sort-icon"></span></th>
      <th data-col="3">Job Title<span class="sort-icon"></span></th>
      <th data-col="4">Location<span class="sort-icon"></span></th>
      <th data-col="5">User Active<span class="sort-icon"></span></th>
      <th data-col="6">Device Active in AD<span class="sort-icon"></span></th>
      <th data-col="7">Last Audit Date<span class="sort-icon"></span></th>
      <th data-col="8">Last Updated Date<span class="sort-icon"></span></th>
      <th data-col="9">AD whenChanged<span class="sort-icon"></span></th>
      <th data-col="10">AD LastLogonDate<span class="sort-icon"></span></th>
    </tr>
  </thead>
  <tbody id="tbody">
{rows_html}
  </tbody>
</table>
</div>

<script>
const tbody = document.getElementById("tbody");
const allRows = Array.from(tbody.querySelectorAll("tr"));
const searchBox = document.getElementById("search");
const filterUser = document.getElementById("filter-user");
const filterDevice = document.getElementById("filter-device");
const filterStale = document.getElementById("filter-stale");
const rowCount = document.getElementById("row-count");

function applyFilters() {{
  const q = searchBox.value.toLowerCase();
  const ua = filterUser.value;
  const da = filterDevice.value;
  const staleOnly = filterStale.checked;
  let visible = 0;

  allRows.forEach(row => {{
    const cells = row.querySelectorAll("td");
    const text = row.textContent.toLowerCase();
    const userActiveCell = cells[5]?.textContent.trim();
    const deviceActiveCell = cells[6]?.textContent.trim();
    const rowCls = row.className;

    const matchQ = !q || text.includes(q);
    const matchUA = !ua || userActiveCell === ua;
    const matchDA = !da || deviceActiveCell === da;
    const matchStale = !staleOnly || rowCls === "stale";

    const show = matchQ && matchUA && matchDA && matchStale;
    row.style.display = show ? "" : "none";
    if (show) visible++;
  }});

  rowCount.textContent = `Showing ${{visible}} of ${{allRows.length}} devices`;
}}

searchBox.addEventListener("input", applyFilters);
filterUser.addEventListener("change", applyFilters);
filterDevice.addEventListener("change", applyFilters);
filterStale.addEventListener("change", applyFilters);
applyFilters();

// Sorting
let sortCol = -1, sortDir = 1;
document.querySelectorAll("th[data-col]").forEach(th => {{
  th.addEventListener("click", () => {{
    const col = parseInt(th.dataset.col);
    if (sortCol === col) sortDir *= -1;
    else {{ sortCol = col; sortDir = 1; }}
    document.querySelectorAll("th").forEach(h => h.classList.remove("asc","desc"));
    th.classList.add(sortDir === 1 ? "asc" : "desc");

    const sorted = allRows.slice().sort((a, b) => {{
      const av = a.querySelectorAll("td")[col]?.textContent.trim() || "";
      const bv = b.querySelectorAll("td")[col]?.textContent.trim() || "";
      return av.localeCompare(bv, undefined, {{numeric: true}}) * sortDir;
    }});
    sorted.forEach(r => tbody.appendChild(r));
    applyFilters();
  }});
}});
</script>
</body>
</html>"""


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 computer-inventory-report.py <input.xlsx> [output.html]")
        sys.exit(1)

    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else Path(src).stem + "-report.html"

    print(f"Reading {src}...")
    rows = load_rows(src)
    print(f"  {len(rows)} devices loaded.")

    print(f"Building report...")
    html_content = build_html(rows, Path(src).name)

    with open(out, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"  Report saved to: {out}")


if __name__ == "__main__":
    main()
