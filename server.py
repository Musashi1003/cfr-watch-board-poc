from __future__ import annotations

import cgi
from html import escape
from http import cookies
import mimetypes
import os
from pathlib import Path
import re
from secrets import token_urlsafe
from socketserver import ThreadingMixIn
from tempfile import NamedTemporaryFile
from urllib.parse import parse_qs
from uuid import uuid4

from wsgiref.simple_server import WSGIServer, make_server

from cfr_watch_analyzer import (
  FILTER_FIELDS,
  WorkbookUpload,
  analyze_dataset,
  dataset_options,
  parse_workbooks,
)
from mail_summary_analyzer import analyze_mail_summary


BASE_DIR = Path(__file__).resolve().parent
INDEX_FILE = BASE_DIR / "index.html"
SCRIPT_FILE = BASE_DIR / "script.js"
STYLE_FILE = BASE_DIR / "styles.css"

DEFAULT_ALLOWED_EMPLOYEE_IDS = {"IEC950458"}
ALLOWED_EMPLOYEE_IDS_ENV = os.environ.get("SITES_ALLOWED_EMPLOYEE_IDS", "")
ALLOWED_WORKBOOK_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
SESSION_COOKIE_NAME = "sites_session"
SESSIONS: dict[str, str] = {}
CFR_DATASETS: dict[str, dict] = {}


class ThreadingWSGIServer(ThreadingMixIn, WSGIServer):
  daemon_threads = True


def load_allowed_employee_ids() -> set[str] | None:
  raw_value = ALLOWED_EMPLOYEE_IDS_ENV.strip()
  if not raw_value:
    return DEFAULT_ALLOWED_EMPLOYEE_IDS

  employee_ids = {
    item.strip().upper()
    for item in re.split(r"[,;\s]+", raw_value)
    if item.strip()
  }
  if "*" in employee_ids or "ALL" in employee_ids:
    return None
  return employee_ids or DEFAULT_ALLOWED_EMPLOYEE_IDS


ALLOWED_EMPLOYEE_IDS = load_allowed_employee_ids()


def is_allowed_employee_id(employee_id: str) -> bool:
  if not employee_id:
    return False
  if ALLOWED_EMPLOYEE_IDS is None:
    return True
  return employee_id in ALLOWED_EMPLOYEE_IDS


LOGIN_PAGE = """<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sites Access</title>
  <style>
    :root {
      --bg: #f4efe7;
      --panel: rgba(255,255,255,0.92);
      --text: #1f2a2c;
      --muted: #5d6b6d;
      --teal: #0d7c86;
      --line: rgba(31,42,44,0.12);
      --shadow: 0 20px 60px rgba(49,55,57,0.12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
      font-family: "Microsoft JhengHei", "PingFang TC", sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at top left, rgba(13,124,134,0.16), transparent 28%),
        radial-gradient(circle at top right, rgba(221,108,79,0.12), transparent 26%),
        linear-gradient(180deg, #fbf6ef 0%, var(--bg) 100%);
    }
    .card {
      width: min(560px, 100%);
      padding: 34px;
      border-radius: 30px;
      border: 1px solid var(--line);
      background: var(--panel);
      box-shadow: var(--shadow);
    }
    .eyebrow {
      margin: 0 0 12px;
      color: var(--teal);
      font-size: 0.78rem;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      font-family: "Segoe UI Variable Display", "Trebuchet MS", sans-serif;
    }
    h1 {
      margin: 0 0 14px;
      font-size: clamp(2rem, 5vw, 3.4rem);
      line-height: 1.05;
      max-width: 10ch;
      font-family: "Segoe UI Variable Display", "Trebuchet MS", sans-serif;
    }
    p { color: var(--muted); line-height: 1.8; }
    form { display: grid; gap: 12px; margin: 24px 0 14px; }
    label {
      font-size: 0.92rem;
      font-family: "Segoe UI Variable Display", "Trebuchet MS", sans-serif;
    }
    input {
      width: 100%;
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid var(--line);
      background: rgba(255,250,242,0.96);
      font: inherit;
      color: var(--text);
    }
    button {
      width: fit-content;
      padding: 14px 20px;
      border: 0;
      border-radius: 999px;
      background: var(--text);
      color: #fff;
      font: inherit;
      cursor: pointer;
    }
    .error { color: #ba4e32; }
  </style>
</head>
<body>
  <div class="card">
    <p class="eyebrow">Employee Access</p>
    <h1>先輸入工號，再進入網站入口</h1>
    <p>使用者需先輸入工號，驗證通過後由 server 建立 session 才能進站。內網共享版本可由啟動檔設定允許名單。</p>
    <form method="post" action="/login">
      <label for="employee-id">工號</label>
      <input id="employee-id" name="employee_id" type="text" placeholder="例如 IEC950458" autocomplete="off">
      <button type="submit">進入網站</button>
    </form>
    <p>Prototype demo 目前只開放示範工號：<strong>IEC950458</strong></p>
    __ERROR_BLOCK__
  </div>
</body>
</html>
"""


def render_login_page(error: str = "") -> bytes:
  error_block = f'<p class="error">{error}</p>' if error else ""
  return LOGIN_PAGE.replace("__ERROR_BLOCK__", error_block).encode("utf-8")


def get_employee_id_from_cookie(environ: dict) -> str | None:
  raw_cookie = environ.get("HTTP_COOKIE", "")
  if not raw_cookie:
    return None
  jar = cookies.SimpleCookie()
  jar.load(raw_cookie)
  morsel = jar.get(SESSION_COOKIE_NAME)
  if morsel is None:
    return None
  return SESSIONS.get(morsel.value)


def parse_form_body(environ: dict) -> dict[str, str]:
  length = int(environ.get("CONTENT_LENGTH") or 0)
  body = environ["wsgi.input"].read(length).decode("utf-8")
  parsed = parse_qs(body, keep_blank_values=True)
  return {key: values[0] for key, values in parsed.items()}


def redirect(start_response, location: str, headers: list[tuple[str, str]] | None = None):
  response_headers = [("Location", location)]
  if headers:
    response_headers.extend(headers)
  start_response("302 Found", response_headers)
  return [b""]


def set_session_headers(employee_id: str) -> list[tuple[str, str]]:
  session_id = token_urlsafe(24)
  SESSIONS[session_id] = employee_id
  cookie = cookies.SimpleCookie()
  cookie[SESSION_COOKIE_NAME] = session_id
  cookie[SESSION_COOKIE_NAME]["path"] = "/"
  cookie[SESSION_COOKIE_NAME]["httponly"] = True
  return [("Set-Cookie", cookie.output(header="").strip())]


def clear_session_headers(environ: dict) -> list[tuple[str, str]]:
  raw_cookie = environ.get("HTTP_COOKIE", "")
  if raw_cookie:
    jar = cookies.SimpleCookie()
    jar.load(raw_cookie)
    morsel = jar.get(SESSION_COOKIE_NAME)
    if morsel is not None:
      SESSIONS.pop(morsel.value, None)

  cookie = cookies.SimpleCookie()
  cookie[SESSION_COOKIE_NAME] = ""
  cookie[SESSION_COOKIE_NAME]["path"] = "/"
  cookie[SESSION_COOKIE_NAME]["max-age"] = 0
  cookie[SESSION_COOKIE_NAME]["httponly"] = True
  return [("Set-Cookie", cookie.output(header="").strip())]


def render_dashboard(employee_id: str) -> bytes:
  html = INDEX_FILE.read_text(encoding="utf-8")
  html = html.replace(
    '<div class="login-shell" id="login-shell">',
    '<div class="login-shell" id="login-shell" hidden>',
    1,
  )
  html = re.sub(
    r'<button class="ghost-btn logout-btn" type="button" id="logout-btn">.*?</button>',
    '<form method="post" action="/logout"><button class="ghost-btn logout-btn" type="submit">登出</button></form>',
    html,
    count=1,
    flags=re.S,
  )
  html = re.sub(
    r'(<strong id="active-employee-id">)(.*?)(</strong>)',
    rf"\1{employee_id}\3",
    html,
    count=1,
    flags=re.S,
  )
  injection = (
    '<section style="margin:0 auto 18px;max-width:1180px;">'
    '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px;">'
    '<div style="padding:18px 22px;border:1px solid rgba(31,42,44,0.12);border-radius:16px;'
    'background:rgba(255,255,255,0.78);box-shadow:0 20px 60px rgba(49,55,57,0.08);">'
    '<p style="margin:0 0 6px;color:#0d7c86;font:600 12px/1.2 Segoe UI Variable Display, Trebuchet MS, sans-serif;letter-spacing:.08em;text-transform:uppercase;">Upload Flow</p>'
    '<a href="/mail-summary" style="color:#1f2a2c;font:700 18px/1.2 Segoe UI Variable Display, Trebuchet MS, sans-serif;text-decoration:none;">Open Excel Mail Summary Generator</a>'
    '<p style="margin:8px 0 0;color:#5d6b6d;">Upload one workbook and generate mail-style summary content from HP-AIO and ASUS related data.</p>'
    '</div>'
    '<div style="padding:18px 22px;border:1px solid rgba(31,42,44,0.12);border-radius:16px;'
    'background:rgba(255,255,255,0.78);box-shadow:0 20px 60px rgba(49,55,57,0.08);">'
    '<p style="margin:0 0 6px;color:#0d7c86;font:600 12px/1.2 Segoe UI Variable Display, Trebuchet MS, sans-serif;letter-spacing:.08em;text-transform:uppercase;">CFR Watch</p>'
    '<a href="/cfr-watch" style="color:#1f2a2c;font:700 18px/1.2 Segoe UI Variable Display, Trebuchet MS, sans-serif;text-decoration:none;">Open Weekly CFR Watch Board</a>'
    '<p style="margin:8px 0 0;color:#5d6b6d;">Upload Gaming / PC weekly raw data and review model, segment, ODM/OEM, module, and problem trends.</p>'
    '</div>'
    '</div></section>\n'
    f'<script>localStorage.setItem("sites_prototype_employee_id", "{employee_id}");</script>\n'
    '<script src="script.js"></script>'
  )
  html = html.replace('<script src="script.js"></script>', injection, 1)
  return html.encode("utf-8")


def render_mail_summary_page(employee_id: str, result: dict | None = None, error: str = "") -> bytes:
  current_rows = ""
  ytd_rows = ""
  summary_block = ""

  if result:
    current_rows = "".join(
      f"""
      <tr>
        <td>{row['label']}</td>
        <td>{row['previous']:,}</td>
        <td>{row['current']:,}</td>
        <td style="color:{'#0e7a63' if row['delta'] >= 0 else '#ba4e32'};">{row['delta']:+,}</td>
        <td style="color:{'#0e7a63' if (row['ratio'] or 0) >= 0 else '#ba4e32'};">{format_pct(row['ratio'])}</td>
      </tr>
      """
      for row in result["current_table"]
    )
    ytd_rows = "".join(
      f"""
      <tr>
        <td>{row['label']}</td>
        <td>{row['previous']:,}</td>
        <td>{row['current']:,}</td>
        <td style="color:{'#0e7a63' if row['delta'] >= 0 else '#ba4e32'};">{row['delta']:+,}</td>
        <td style="color:{'#0e7a63' if (row['ratio'] or 0) >= 0 else '#ba4e32'};">{format_pct(row['ratio'])}</td>
      </tr>
      """
      for row in result["ytd_table"]
    )
    summary_block = f"""
      <section class="panel">
        <p class="eyebrow">Generated Summary</p>
        <h2>{result['workbook_name']}</h2>
        <div class="note-box">
          <strong>Mail-style headline</strong>
          <ol>
            {''.join(f'<li>{item}</li>' for item in result['headline_bullets'])}
          </ol>
        </div>
        <div class="grid two">
          <article class="card">
            <h3>Current Month Compare</h3>
            <p>{result['current_month_label']} vs {result['previous_month_label']}</p>
            <table>
              <thead>
                <tr><th>Customers/Month</th><th>{result['previous_month_label']}</th><th>{result['current_month_label']}</th><th>差異</th><th>Ratio</th></tr>
              </thead>
              <tbody>{current_rows}</tbody>
              <tfoot>
                <tr><td>Total</td><td>{result['current_total']['previous']:,}</td><td>{result['current_total']['current']:,}</td><td>{result['current_total']['delta']:+,}</td><td>{format_pct(result['current_total']['ratio'])}</td></tr>
              </tfoot>
            </table>
          </article>
          <article class="card">
            <h3>Accumulated Compare</h3>
            <p>{result['ytd_previous_label']} vs {result['ytd_current_label']}</p>
            <table>
              <thead>
                <tr><th>Customers/Month</th><th>{result['ytd_previous_label']}</th><th>{result['ytd_current_label']}</th><th>差異</th><th>Ratio</th></tr>
              </thead>
              <tbody>{ytd_rows}</tbody>
              <tfoot>
                <tr><td>Total</td><td>{result['ytd_total']['previous']:,}</td><td>{result['ytd_total']['current']:,}</td><td>{result['ytd_total']['delta']:+,}</td><td>{format_pct(result['ytd_total']['ratio'])}</td></tr>
              </tfoot>
            </table>
          </article>
        </div>
        <div class="grid two">
          <article class="card">
            <h3>Supporting Notes</h3>
            <ul>
              {''.join(f'<li>{item}</li>' for item in result['detail_bullets'])}
            </ul>
          </article>
          <article class="card">
            <h3>Source Notes</h3>
            <ul>
              {''.join(f'<li>{item}</li>' for item in result['source_notes'])}
            </ul>
          </article>
        </div>
      </section>
    """

  error_html = f'<p class="error">{error}</p>' if error else ""
  page = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Mail Summary Generator</title>
  <style>
    :root {{
      --bg: #f4efe7;
      --panel: rgba(255,255,255,0.92);
      --text: #1f2a2c;
      --muted: #5d6b6d;
      --teal: #0d7c86;
      --line: rgba(31,42,44,0.12);
      --shadow: 0 20px 60px rgba(49,55,57,0.12);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      padding: 28px 16px 56px;
      font-family: "Microsoft JhengHei", "PingFang TC", sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at top left, rgba(13,124,134,0.16), transparent 28%),
        radial-gradient(circle at top right, rgba(221,108,79,0.12), transparent 26%),
        linear-gradient(180deg, #fbf6ef 0%, var(--bg) 100%);
    }}
    .shell {{ max-width: 1180px; margin: 0 auto; }}
    .bar, .panel {{
      border: 1px solid var(--line);
      border-radius: 28px;
      background: var(--panel);
      box-shadow: var(--shadow);
    }}
    .bar {{
      display: flex; justify-content: space-between; align-items: center; gap: 16px;
      padding: 18px 22px; margin-bottom: 18px;
    }}
    .panel {{ padding: 26px; margin-bottom: 18px; }}
    .eyebrow {{
      margin: 0 0 10px; color: var(--teal); font: 600 12px/1.2 Segoe UI Variable Display, Trebuchet MS, sans-serif;
      letter-spacing: .08em; text-transform: uppercase;
    }}
    h1, h2, h3 {{ margin: 0; font-family: "Segoe UI Variable Display", "Trebuchet MS", sans-serif; }}
    h1 {{ font-size: clamp(2.2rem, 5vw, 3.8rem); max-width: 12ch; }}
    p, li {{ color: var(--muted); line-height: 1.8; }}
    .actions {{ display: flex; gap: 12px; flex-wrap: wrap; margin-top: 20px; }}
    .btn, button {{
      display: inline-flex; align-items: center; justify-content: center;
      padding: 12px 18px; border-radius: 999px; border: 0; background: #1f2a2c; color: #fff;
      font: inherit; cursor: pointer; text-decoration: none;
    }}
    .btn.alt {{ background: transparent; color: #1f2a2c; border: 1px solid var(--line); }}
    .input-row {{ display: grid; gap: 12px; margin-top: 18px; }}
    input[type=file] {{
      padding: 14px 16px; border-radius: 16px; border: 1px solid var(--line); background: rgba(255,250,242,0.96);
    }}
    .grid.two {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 18px; }}
    .card {{ padding: 18px; border-radius: 22px; border: 1px solid rgba(31,42,44,0.08); background: rgba(255,255,255,0.72); }}
    .note-box {{ padding: 18px; border-radius: 22px; background: linear-gradient(180deg, rgba(218,243,241,0.5), rgba(255,255,255,0.76)); margin: 18px 0; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ padding: 10px 8px; border-bottom: 1px solid rgba(31,42,44,0.08); text-align: left; font-size: 0.92rem; }}
    th {{ font-family: "Segoe UI Variable Display", "Trebuchet MS", sans-serif; }}
    tfoot td {{ font-weight: 700; }}
    .error {{ color: #ba4e32; }}
    @media (max-width: 920px) {{
      .grid.two {{ grid-template-columns: 1fr; }}
      .bar {{ flex-direction: column; align-items: flex-start; }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <div class="bar">
      <div>
        <p class="eyebrow">Signed In</p>
        <strong>{employee_id}</strong>
      </div>
      <div class="actions">
        <a class="btn alt" href="/">Back To Dashboard</a>
        <form method="post" action="/logout" style="margin:0;"><button type="submit">登出</button></form>
      </div>
    </div>
    <section class="panel">
      <p class="eyebrow">Upload Workbook</p>
      <h1>上傳 Excel，產出郵件式 summary</h1>
      <p>目前這版會優先對齊你們郵件中的 customer summary 表，並用 HP-AIO / ASUS 相關資料產出可讀的 mail-style summary。</p>
      <form method="post" action="/mail-summary" enctype="multipart/form-data" class="input-row">
        <input type="file" name="workbook" accept=".xlsx,.xlsm,.xls" required>
        <div class="actions">
          <button type="submit">Generate Summary</button>
        </div>
      </form>
      {error_html}
    </section>
    {summary_block}
  </div>
</body>
</html>"""
  return page.encode("utf-8")


def html_escape(value) -> str:
  return escape(str(value), quote=True)


def format_number(value: int | float) -> str:
  return f"{value:,}"


def format_plain_pct(value: float | None) -> str:
  if value is None:
    return "N/A"
  return f"{value:.1f}%"


def format_cfr_pct(value: float | None) -> str:
  if value is None:
    return "N/A"
  return f"{value * 100:.2f}%"


def selected_values(value) -> set[str]:
  if value is None:
    return set()
  if isinstance(value, str):
    values = [value]
  else:
    values = list(value)
  return {
    str(item).strip()
    for item in values
    if str(item).strip()
  }


def checkbox_filter(name: str, title: str, options: list[str], selected) -> str:
  selected_set = selected_values(selected)
  if selected_set:
    selected_preview = ", ".join(list(selected_set)[:2])
    if len(selected_set) > 2:
      selected_preview += f" +{len(selected_set) - 2}"
    status = f"{len(selected_set)} selected"
  else:
    selected_preview = "All"
    status = "All"

  items = "".join(
    f"""
    <label class="check-option">
      <input class="filter-value" type="checkbox" name="{html_escape(name)}" value="{html_escape(value)}"{" checked" if value in selected_set else ""}>
      <span>{html_escape(value)}</span>
    </label>
    """
    for value in options
  )
  all_option = f"""
    <label class="check-option all-option">
      <input class="filter-all" type="checkbox"{" checked" if not selected_set else ""}>
      <span>All</span>
    </label>
  """

  return f"""
    <details class="filter-box">
      <summary>
        <span>
          <strong>{html_escape(title)}</strong>
          <em>{html_escape(selected_preview)}</em>
        </span>
        <b>{html_escape(status)}</b>
      </summary>
      <div class="filter-list">
        {all_option}
        {items or '<div class="empty-state">No filter values.</div>'}
      </div>
    </details>
  """


def render_trend_svg(trend: list[dict]) -> str:
  if not trend:
    return '<div class="empty-state">No weekly trend available for the current filters.</div>'

  width = 820
  height = 230
  padding_x = 36
  padding_y = 28
  values = [point["count"] for point in trend]
  labels = [point["week"] for point in trend]
  max_value = max(values) or 1
  step = (width - padding_x * 2) / max(len(values) - 1, 1)
  points = []
  for index, value in enumerate(values):
    x = padding_x + step * index
    y = height - padding_y - (value / max_value) * (height - padding_y * 2)
    points.append((x, y, value, labels[index]))
  path = " ".join(
    f"{'M' if index == 0 else 'L'} {x:.1f} {y:.1f}"
    for index, (x, y, _, _) in enumerate(points)
  )
  dots = "".join(
    f'<circle cx="{x:.1f}" cy="{y:.1f}" r="5"><title>{html_escape(label)}: {value}</title></circle>'
    for x, y, value, label in points
  )
  axis = "".join(
    f'<span>{html_escape(label)}</span>'
    for label in labels
  )
  return f"""
    <div class="trend-chart">
      <svg viewBox="0 0 {width} {height}" preserveAspectRatio="none" aria-label="Weekly trend chart">
        <line x1="{padding_x}" y1="{padding_y}" x2="{width - padding_x}" y2="{padding_y}"></line>
        <line x1="{padding_x}" y1="{height / 2}" x2="{width - padding_x}" y2="{height / 2}"></line>
        <line x1="{padding_x}" y1="{height - padding_y}" x2="{width - padding_x}" y2="{height - padding_y}"></line>
        <path d="{path}"></path>
        {dots}
      </svg>
      <div class="trend-axis">{axis}</div>
    </div>
  """


def render_bar_rows(rows: list[dict], label_key: str = "label", max_rows: int | None = None) -> str:
  selected_rows = rows[:max_rows] if max_rows else rows
  max_count = max((row["count"] for row in selected_rows), default=1)
  return "".join(
    f"""
    <article class="bar-row">
      <div class="row-head">
        <strong>{html_escape(row[label_key])}</strong>
        <span>{format_number(row["count"])} / {format_plain_pct(row.get("share"))}</span>
      </div>
      <div class="bar-track"><div class="bar-fill" style="--bar-width:{(row["count"] / max_count) * 100:.2f}%;"></div></div>
    </article>
    """
    for row in selected_rows
  )


def render_pareto_rows(rows: list[dict]) -> str:
  max_count = max((row["count"] for row in rows), default=1)
  return "".join(
    f"""
    <tr>
      <td>{html_escape(row["label"])}</td>
      <td>{format_number(row["count"])}</td>
      <td>{format_plain_pct(row["share"])}</td>
      <td>
        <div class="pareto-track">
          <span style="--bar-width:{(row["count"] / max_count) * 100:.2f}%;"></span>
        </div>
      </td>
      <td>{format_plain_pct(row["cumulative"])}</td>
    </tr>
    """
    for row in rows
  )


def render_heatmap(result: dict) -> str:
  columns = result["heatmap_columns"]
  rows = result["heatmap"]
  if not columns or not rows:
    return '<div class="empty-state">No heatmap data available for the current filters.</div>'

  max_value = max((value["count"] for row in rows for value in row["values"]), default=1)
  max_cfr = max((value["cfr"] or 0 for row in rows for value in row["values"]), default=0)
  header = "".join(f"<th>{html_escape(column)}</th>" for column in columns)
  body = ""
  for row in rows:
    cells = ""
    for value in row["values"]:
      count = value["count"]
      cfr = value.get("cfr")
      intensity = (cfr / max_cfr) if max_cfr and cfr is not None else (count / max_value if max_value else 0)
      cfr_line = f'<small>CFR {format_cfr_pct(cfr)}</small>' if count and cfr is not None else ""
      title = f'{row["label"]} / {count} failures'
      if cfr is not None:
        title += f' / CFR {format_cfr_pct(cfr)}'
      cells += (
        f'<td style="--heat:{intensity:.3f};" title="{html_escape(title)}">'
        f'<strong>{count}</strong>{cfr_line}</td>'
      )
    body += f"<tr><th>{html_escape(row['label'])}</th>{cells}</tr>"
  return f"""
    <div class="table-scroll">
      <table class="heatmap">
        <thead><tr><th>{html_escape(result["primary_label"])}</th>{header}</tr></thead>
        <tbody>{body}</tbody>
      </table>
    </div>
  """


def render_cfr_watch_page(
  employee_id: str,
  result: dict | None = None,
  dataset_id: str = "",
  options: dict[str, list[str]] | None = None,
  files: list[dict] | None = None,
  missing_columns: dict[str, list[str]] | None = None,
  selections: dict[str, str] | None = None,
  error: str = "",
) -> bytes:
  options = options or {"model": [], "segment": [], "odm_oem": [], "muc_module": [], "problem_mapping": []}
  files = files or []
  missing_columns = missing_columns or {}
  selections = selections or {
    "model": [],
    "segment": [],
    "odm_oem": [],
    "muc_module": [],
    "problem_mapping": [],
  }

  error_html = f'<p class="error">{html_escape(error)}</p>' if error else ""
  upload_summary = ""
  control_panel = ""
  result_panel = ""

  if files:
    upload_summary = f"""
      <section class="panel compact-panel">
        <p class="eyebrow">Loaded Weekly Raw Data</p>
        <div class="file-list">
          {''.join(f'<article><strong>{html_escape(item["source_type"])}</strong><span>{html_escape(item["filename"])}</span><em>{format_number(item["rows"])} rows</em></article>' for item in files)}
        </div>
      </section>
    """

  if missing_columns:
    upload_summary += f"""
      <section class="panel warning-panel">
        <p class="eyebrow">Column Notice</p>
        {''.join(f'<p><strong>{html_escape(filename)}</strong>: missing {html_escape(", ".join(columns))}</p>' for filename, columns in missing_columns.items())}
      </section>
    """

  if result and dataset_id:
    control_panel = f"""
      <section class="panel">
        <p class="eyebrow">Analysis Controls</p>
        <form method="post" action="/cfr-watch" class="control-grid">
          <input type="hidden" name="dataset_id" value="{html_escape(dataset_id)}">
          {checkbox_filter("model", "Model filter", options.get("model", []), selections.get("model", []))}
          {checkbox_filter("segment", "Segment filter", options.get("segment", []), selections.get("segment", []))}
          {checkbox_filter("odm_oem", "ODM/OEM filter", options.get("odm_oem", []), selections.get("odm_oem", []))}
          {checkbox_filter("muc_module", "Module filter", options.get("muc_module", []), selections.get("muc_module", []))}
          {checkbox_filter("problem_mapping", "Problem filter", options.get("problem_mapping", []), selections.get("problem_mapping", []))}
          <button type="submit">Refresh Watch Board</button>
        </form>
      </section>
    """

    kpis = result["kpis"]
    source_mix = render_bar_rows(result["source_mix"])
    share_rows = render_bar_rows(result["share"])
    module_rows = render_bar_rows(result["module_ratio"])
    pareto_rows = render_pareto_rows(result["pareto"])
    heatmap = render_heatmap(result)
    combinations = "".join(
      f"""
      <tr>
        <td>{html_escape(row["primary"])}</td>
        <td>{html_escape(row["breakdown"])}</td>
        <td>{format_number(row["count"])}</td>
      </tr>
      """
      for row in result["top_combinations"]
    )

    result_panel = f"""
      <section class="kpi-grid">
        <article><span>Filtered CFR</span><strong>{format_cfr_pct(kpis["filtered_cfr"])}</strong><p>Failure Qty / Derived ACT</p></article>
        <article><span>Failure Qty</span><strong>{format_number(kpis["filtered_failure_qty"])}</strong><p>Filtered raw-data failures</p></article>
        <article><span>Derived ACT</span><strong>{format_number(round(kpis["derived_act"])) if kpis["derived_act"] else "N/A"}</strong><p>From SUMMARY_IEC CFR rule</p></article>
        <article><span>ACT Model Coverage</span><strong>{format_number(kpis["act_model_count"])}</strong><p>Target CFR: {format_cfr_pct(kpis["target_cfr"])}</p></article>
        <article><span>Latest WoW</span><strong>{kpis["week_delta"]:+,}</strong><p>{html_escape(kpis["previous_week"])} to {html_escape(kpis["latest_week"])}</p></article>
      </section>

      <section class="grid two">
        <article class="panel wide">
          <div class="panel-head"><div><p class="eyebrow">Trend</p><h2>Weekly Failure Trend</h2></div><span>{html_escape(result["primary_label"])} x {html_escape(result["breakdown_label"])}</span></div>
          {render_trend_svg(result["trend"])}
        </article>
        <article class="panel">
          <div class="panel-head"><div><p class="eyebrow">Source Mix</p><h2>Gaming / PC Mix</h2></div></div>
          <div class="stack">{source_mix}</div>
        </article>
      </section>

      <section class="grid two">
        <article class="panel">
          <div class="panel-head"><div><p class="eyebrow">Share</p><h2>{html_escape(result["primary_label"])} Share</h2></div></div>
          <div class="stack">{share_rows or '<div class="empty-state">No share data.</div>'}</div>
        </article>
        <article class="panel">
          <div class="panel-head"><div><p class="eyebrow">Ratio</p><h2>MUC_MODULE Ratio</h2></div></div>
          <div class="stack">{module_rows or '<div class="empty-state">No module data.</div>'}</div>
        </article>
      </section>

      <section class="panel">
          <div class="panel-head"><div><p class="eyebrow">Pareto</p><h2>{html_escape(result["breakdown_label"])} Pareto</h2></div></div>
          <div class="table-scroll">
            <table>
              <thead><tr><th>Item</th><th>Count</th><th>Share</th><th>Bar</th><th>Cumulative</th></tr></thead>
              <tbody>{pareto_rows}</tbody>
            </table>
          </div>
      </section>

      <section class="panel">
        <div class="panel-head"><div><p class="eyebrow">Watch Matrix</p><h2>{html_escape(result["primary_label"])} by {html_escape(result["breakdown_label"])}</h2></div></div>
        {heatmap}
      </section>

      <section class="panel">
        <div class="panel-head"><div><p class="eyebrow">Top Watch Items</p><h2>Highest Combination Counts</h2></div></div>
        <div class="table-scroll">
          <table>
            <thead><tr><th>{html_escape(result["primary_label"])}</th><th>{html_escape(result["breakdown_label"])}</th><th>Count</th></tr></thead>
            <tbody>{combinations}</tbody>
          </table>
        </div>
      </section>
    """

  page = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CFR Watch Board</title>
  <style>
    :root {{
      --bg:#f6f3ed; --panel:#fffdf8; --text:#1e292b; --muted:#607074; --teal:#087984;
      --green:#178061; --red:#b94d35; --amber:#c98518; --line:rgba(30,41,43,.12);
      --shadow:0 18px 45px rgba(30,41,43,.09);
    }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; padding:24px 16px 60px; background:var(--bg); color:var(--text); font-family:"Microsoft JhengHei","Segoe UI",sans-serif; }}
    .shell {{ max-width:1280px; margin:0 auto; }}
    .topbar, .panel, .kpi-grid article {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; box-shadow:var(--shadow); }}
    .topbar {{ display:flex; justify-content:space-between; gap:16px; align-items:center; padding:16px 18px; margin-bottom:16px; }}
    .topbar div:last-child {{ display:flex; gap:10px; flex-wrap:wrap; align-items:center; }}
    .panel {{ padding:20px; margin-bottom:16px; }}
    .compact-panel {{ padding:16px 20px; }}
    .hero {{ padding:24px; margin-bottom:16px; background:#fffdf8; border:1px solid var(--line); border-radius:8px; box-shadow:var(--shadow); }}
    .eyebrow {{ margin:0 0 8px; color:var(--teal); font:700 12px/1.2 "Segoe UI",sans-serif; letter-spacing:.08em; text-transform:uppercase; }}
    h1,h2 {{ margin:0; line-height:1.08; font-family:"Segoe UI","Microsoft JhengHei",sans-serif; }}
    h1 {{ font-size:clamp(2rem,4vw,3.6rem); }}
    h2 {{ font-size:1.35rem; }}
    p {{ color:var(--muted); line-height:1.7; }}
    .btn, button {{ border:0; border-radius:6px; background:var(--text); color:#fff; padding:11px 14px; font:inherit; cursor:pointer; text-decoration:none; }}
    .btn.alt {{ background:transparent; color:var(--text); border:1px solid var(--line); }}
    .upload-form {{ display:grid; grid-template-columns:1fr auto; gap:12px; align-items:end; margin-top:18px; }}
    input[type=file] {{ width:100%; padding:11px 12px; border:1px solid var(--line); border-radius:6px; background:#fff; color:var(--text); font:inherit; }}
    label {{ display:grid; gap:7px; color:var(--muted); font-size:.9rem; }}
    .control-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:12px; align-items:start; }}
    .filter-box {{ position:relative; min-width:0; }}
    .filter-box summary {{
      display:flex; justify-content:space-between; gap:12px; align-items:center;
      min-height:58px; padding:10px 12px; border:1px solid var(--line); border-radius:6px;
      background:#fff; color:var(--text); cursor:pointer; list-style:none;
    }}
    .filter-box summary::-webkit-details-marker {{ display:none; }}
    .filter-box summary::after {{ content:""; width:8px; height:8px; border-right:2px solid var(--muted); border-bottom:2px solid var(--muted); transform:rotate(45deg); flex:0 0 auto; margin-left:auto; }}
    .filter-box[open] summary::after {{ transform:rotate(225deg); margin-top:6px; }}
    .filter-box summary span {{ display:grid; gap:3px; min-width:0; }}
    .filter-box summary strong {{ font-size:.92rem; font-weight:600; color:var(--text); }}
    .filter-box summary em {{ color:var(--muted); font-style:normal; font-size:.82rem; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:18ch; }}
    .filter-box summary b {{ padding:4px 8px; border-radius:999px; background:rgba(8,121,132,.1); color:var(--teal); font-size:.75rem; font-weight:700; white-space:nowrap; }}
    .filter-list {{
      position:absolute; z-index:20; inset:auto 0 auto 0; max-height:260px; overflow:auto;
      margin-top:6px; padding:10px; border:1px solid var(--line); border-radius:8px;
      background:#fff; box-shadow:0 18px 45px rgba(30,41,43,.18);
    }}
    .check-option {{ display:flex; grid-template-columns:none; flex-direction:row; align-items:flex-start; gap:8px; padding:8px; border-radius:6px; color:var(--text); cursor:pointer; }}
    .check-option:hover {{ background:rgba(8,121,132,.08); }}
    .check-option input {{ width:16px; height:16px; margin-top:2px; flex:0 0 auto; accent-color:var(--teal); }}
    .check-option span {{ line-height:1.35; overflow-wrap:anywhere; }}
    .all-option {{ margin-bottom:6px; border-bottom:1px solid var(--line); border-radius:6px 6px 0 0; font-weight:700; }}
    .file-list {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(260px,1fr)); gap:10px; }}
    .file-list article {{ border:1px solid var(--line); border-radius:8px; padding:12px; display:grid; gap:4px; }}
    .file-list span,.file-list em {{ color:var(--muted); font-style:normal; overflow-wrap:anywhere; }}
    .warning-panel {{ border-color:rgba(201,133,24,.35); background:#fff9ec; }}
    .error {{ color:var(--red); }}
    .kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:12px; margin-bottom:16px; }}
    .kpi-grid article {{ padding:18px; }}
    .kpi-grid span,.panel-head span {{ color:var(--muted); font-size:.88rem; }}
    .kpi-grid strong {{ display:block; margin:8px 0; font-size:2rem; font-family:"Segoe UI",sans-serif; }}
    .grid.two {{ display:grid; grid-template-columns:1.2fr .8fr; gap:16px; }}
    .panel-head {{ display:flex; justify-content:space-between; gap:12px; align-items:flex-start; margin-bottom:16px; }}
    .trend-chart svg {{ display:block; width:100%; height:230px; }}
    .trend-chart line {{ stroke:rgba(30,41,43,.12); stroke-width:1; }}
    .trend-chart path {{ fill:none; stroke:var(--teal); stroke-width:4; stroke-linecap:round; stroke-linejoin:round; }}
    .trend-chart circle {{ fill:#fff; stroke:var(--teal); stroke-width:3; }}
    .trend-axis {{ display:flex; justify-content:space-between; gap:8px; color:var(--muted); font-size:.78rem; overflow-x:auto; }}
    .stack {{ display:grid; gap:10px; }}
    .bar-row {{ border:1px solid var(--line); border-radius:8px; padding:12px; }}
    .row-head {{ display:flex; justify-content:space-between; gap:12px; align-items:baseline; margin-bottom:10px; }}
    .row-head strong {{ overflow-wrap:anywhere; }}
    .row-head span {{ color:var(--muted); white-space:nowrap; }}
    .bar-track,.pareto-track {{ height:10px; border-radius:999px; background:rgba(30,41,43,.08); overflow:hidden; }}
    .bar-fill,.pareto-track span {{ display:block; height:100%; width:var(--bar-width,0%); background:linear-gradient(90deg,var(--teal),#58b6bd); }}
    .table-scroll {{ overflow:auto; }}
    table {{ width:100%; border-collapse:collapse; }}
    th,td {{ padding:10px; border-bottom:1px solid var(--line); text-align:left; vertical-align:middle; font-size:.92rem; }}
    th {{ color:var(--muted); font-weight:700; }}
    .heatmap td {{ background:rgba(8,121,132,calc(.06 + var(--heat) * .36)); min-width:86px; text-align:center; }}
    .heatmap td strong {{ display:block; line-height:1.1; }}
    .heatmap td small {{ display:block; margin-top:3px; color:rgba(30,41,43,.72); font-size:.72rem; white-space:nowrap; }}
    .empty-state {{ color:var(--muted); padding:18px; border:1px dashed var(--line); border-radius:8px; }}
    @media (max-width:980px) {{
      .topbar,.upload-form {{ grid-template-columns:1fr; flex-direction:column; align-items:flex-start; }}
      .control-grid,.kpi-grid,.grid.two {{ grid-template-columns:1fr; }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <div class="topbar">
      <div><p class="eyebrow">Signed In</p><strong>{html_escape(employee_id)}</strong></div>
      <div>
        <a class="btn alt" href="/">Back To Sites</a>
        <form method="post" action="/logout" style="margin:0;"><button type="submit">Logout</button></form>
      </div>
    </div>

    <section class="hero">
      <p class="eyebrow">Weekly CFR Raw Data</p>
      <h1>CFR Watch Board</h1>
      <p>Upload the weekly Gaming NB and PC NB workbooks. The board reads the <strong>raw data</strong> sheet, then lets the team switch between ORG_MODEL(PRODUCT_DESC), Segment, ODM_OEM, MUC_MODULE, and PROBLEM_Mapping views.</p>
      <form method="post" action="/cfr-watch" enctype="multipart/form-data" class="upload-form">
        <label>
          Weekly raw data files
          <input type="file" name="workbooks" accept=".xlsx,.xlsm,.xls" multiple required>
        </label>
        <button type="submit">Upload And Analyze</button>
      </form>
      {error_html}
    </section>

    {upload_summary}
    {control_panel}
    {result_panel}
  </div>
  <script>
    document.querySelectorAll(".filter-box").forEach((filterBox) => {{
      const allBox = filterBox.querySelector(".filter-all");
      const valueBoxes = Array.from(filterBox.querySelectorAll(".filter-value"));
      const syncAllState = () => {{
        allBox.checked = valueBoxes.every((box) => !box.checked);
      }};
      allBox?.addEventListener("change", () => {{
        if (!allBox.checked) {{
          syncAllState();
          return;
        }}
        valueBoxes.forEach((box) => {{
          box.checked = false;
        }});
      }});
      valueBoxes.forEach((box) => {{
        box.addEventListener("change", () => {{
          if (box.checked && allBox) {{
            allBox.checked = false;
          }}
          syncAllState();
        }});
      }});
      syncAllState();
    }});
  </script>
</body>
</html>"""
  return page.encode("utf-8")


def serve_file(path: Path, start_response):
  if not path.exists():
    start_response("404 Not Found", [("Content-Type", "text/plain; charset=utf-8")])
    return [b"Not found"]

  mime_type, _ = mimetypes.guess_type(path.name)
  content_type = mime_type or "application/octet-stream"
  if content_type.startswith("text/") or path.suffix in {".js", ".css", ".html"}:
    content_type = f"{content_type}; charset=utf-8"
  start_response("200 OK", [("Content-Type", content_type)])
  return [path.read_bytes()]


def parse_upload_form(environ):
  return cgi.FieldStorage(fp=environ["wsgi.input"], environ=environ, keep_blank_values=True)


def delete_temp_file(path: Path):
  try:
    path.unlink(missing_ok=True)
  except PermissionError:
    pass


def format_pct(value: float | None) -> str:
  if value is None:
    return "N/A"
  sign = "+" if value >= 0 else ""
  return f"{sign}{value:.2f}%"


def application(environ, start_response):
  method = environ.get("REQUEST_METHOD", "GET").upper()
  path = environ.get("PATH_INFO", "/")
  employee_id = get_employee_id_from_cookie(environ)

  if path == "/health":
    start_response("200 OK", [("Content-Type", "text/plain; charset=utf-8")])
    return [b"ok"]

  if path == "/styles.css":
    return serve_file(STYLE_FILE, start_response)
  if path == "/script.js":
    return serve_file(SCRIPT_FILE, start_response)

  if path == "/login":
    if method == "GET":
      if employee_id:
        return redirect(start_response, "/")
      start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
      return [render_login_page()]

    if method == "POST":
      form = parse_form_body(environ)
      submitted = form.get("employee_id", "").strip().upper()
      if is_allowed_employee_id(submitted):
        return redirect(start_response, "/", set_session_headers(submitted))
      start_response("401 Unauthorized", [("Content-Type", "text/html; charset=utf-8")])
      return [render_login_page("Employee ID is not in the current site allowlist.")]

  if path == "/logout" and method == "POST":
    return redirect(start_response, "/login", clear_session_headers(environ))

  if path == "/mail-summary":
    if not employee_id:
      return redirect(start_response, "/login")
    if method == "GET":
      start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
      return [render_mail_summary_page(employee_id)]
    if method == "POST":
      form = parse_upload_form(environ)
      upload = form["workbook"] if "workbook" in form else None
      if upload is None or not getattr(upload, "filename", ""):
        start_response("400 Bad Request", [("Content-Type", "text/html; charset=utf-8")])
        return [render_mail_summary_page(employee_id, error="Please choose an Excel workbook first.")]

      suffix = Path(upload.filename).suffix.lower()
      if suffix not in ALLOWED_WORKBOOK_SUFFIXES:
        start_response("400 Bad Request", [("Content-Type", "text/html; charset=utf-8")])
        return [render_mail_summary_page(employee_id, error="Only .xlsx, .xlsm, and .xls files are supported.")]

      with NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
        temp_path = Path(temp_file.name)
        temp_file.write(upload.file.read())

      try:
        result = analyze_mail_summary(temp_path)
      except Exception as exc:
        delete_temp_file(temp_path)
        start_response("500 Internal Server Error", [("Content-Type", "text/html; charset=utf-8")])
        return [render_mail_summary_page(employee_id, error=f"Summary generation failed: {exc}")]

      delete_temp_file(temp_path)
      start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
      return [render_mail_summary_page(employee_id, result=result)]

  if path == "/cfr-watch":
    if not employee_id:
      return redirect(start_response, "/login")
    if method == "GET":
      start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
      return [render_cfr_watch_page(employee_id)]

    if method == "POST":
      form = parse_upload_form(environ)
      dataset_id = (form.getvalue("dataset_id") or "").strip()
      parsed: dict | None = None
      temp_paths: list[Path] = []

      upload_fields = []
      if "workbooks" in form:
        raw_uploads = form["workbooks"]
        upload_fields = raw_uploads if isinstance(raw_uploads, list) else [raw_uploads]
        upload_fields = [item for item in upload_fields if getattr(item, "filename", "")]

      if upload_fields:
        uploads: list[WorkbookUpload] = []
        for upload in upload_fields:
          suffix = Path(upload.filename).suffix.lower()
          if suffix not in ALLOWED_WORKBOOK_SUFFIXES:
            start_response("400 Bad Request", [("Content-Type", "text/html; charset=utf-8")])
            return [render_cfr_watch_page(employee_id, error="Only .xlsx, .xlsm, and .xls files are supported.")]
          with NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            temp_path = Path(temp_file.name)
            temp_file.write(upload.file.read())
          temp_paths.append(temp_path)
          uploads.append(WorkbookUpload(path=temp_path, filename=Path(upload.filename).name))

        try:
          parsed = parse_workbooks(uploads)
        except Exception as exc:
          for temp_path in temp_paths:
            delete_temp_file(temp_path)
          start_response("500 Internal Server Error", [("Content-Type", "text/html; charset=utf-8")])
          return [render_cfr_watch_page(employee_id, error=f"CFR analysis failed: {exc}")]

        for temp_path in temp_paths:
          delete_temp_file(temp_path)

        dataset_id = uuid4().hex
        CFR_DATASETS[dataset_id] = parsed
      elif dataset_id in CFR_DATASETS:
        parsed = CFR_DATASETS[dataset_id]
      else:
        start_response("400 Bad Request", [("Content-Type", "text/html; charset=utf-8")])
        return [render_cfr_watch_page(employee_id, error="Please upload weekly raw data files first.")]

      records = parsed.get("records", [])
      selections = {
        "model": form.getlist("model"),
        "segment": form.getlist("segment"),
        "odm_oem": form.getlist("odm_oem"),
        "muc_module": form.getlist("muc_module"),
        "problem_mapping": form.getlist("problem_mapping"),
      }
      result = analyze_dataset(
        records,
        summary_by_model=parsed.get("summary_by_model", {}),
        primary_dimension="org_model",
        breakdown_dimension="problem_mapping",
        filters={
          "model": selections["model"],
          "segment": selections["segment"],
          "odm_oem": selections["odm_oem"],
          "muc_module": selections["muc_module"],
          "problem_mapping": selections["problem_mapping"],
        },
      )
      start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
      return [
        render_cfr_watch_page(
          employee_id,
          result=result,
          dataset_id=dataset_id,
          options=dataset_options(records, selections),
          files=parsed.get("files", []),
          missing_columns=parsed.get("missing_columns", {}),
          selections=selections,
        )
      ]

  if path == "/":
    if not employee_id:
      return redirect(start_response, "/login")
    start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
    return [render_dashboard(employee_id)]

  start_response("404 Not Found", [("Content-Type", "text/plain; charset=utf-8")])
  return [b"Not found"]


def main():
  host = os.environ.get("SITES_HOST", "127.0.0.1")
  port = int(os.environ.get("SITES_PORT", "8088"))
  with make_server(host, port, application, server_class=ThreadingWSGIServer) as server:
    print(f"Serving on http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
  main()
