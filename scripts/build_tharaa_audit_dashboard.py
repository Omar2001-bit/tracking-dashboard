from __future__ import annotations

import json
import os
import re
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
SOURCE = Path(
    os.environ.get(
        "THARAA_AUDIT_WORKBOOK",
        r"C:\Users\Omar Maged\Downloads\tharaa_audit_command_center_with_behavioral_tool.xlsx",
    )
)
OUTPUT = ROOT / "index.html"


def clean(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def extract_issue_id(title: str) -> str:
    bracket = re.search(r"\|\s*([^|\]]+?)\s*(?:\||\])", title)
    if bracket:
        return bracket.group(1).strip()
    direct = re.search(r"\b([A-Z]{2,}(?:-[A-Z0-9]+)+-\d{3})\b", title)
    return direct.group(1).strip() if direct else ""


def extract_status(title: str, detail: str, fix: str) -> str:
    haystack = f"{title}\n{detail}\n{fix}".lower()
    if re.search(r"\|\s*correct\s*\]", title, re.I) or 'audit status is "correct"' in haystack:
        return "Correct"
    if re.search(r"\|\s*unable to verify\s*\]", title, re.I) or 'audit status is "unable to verify"' in haystack:
        return "Unable to Verify"
    if re.search(r"\|\s*critical\s*\]", title, re.I) or " critical" in haystack[:600]:
        return "Critical"
    if re.search(r"\|\s*high\s*\]", title, re.I) or " high" in haystack[:600]:
        return "High"
    if "no action required" in fix.lower()[:120]:
        return "Correct"
    return "Action Needed"


def requires_manual_review(title: str, detail: str, fix: str, verify: str) -> bool:
    title_patterns = [
        r"\]\s*(confirm|verify|check|review|inspect|test)\b",
        r"\]\s*(is|are|does|do|can|should|has|have)\b.*\?$",
        r"\b(manual|unable to verify|not confirmed|unverified|unknown|indeterminate|blocked)\b",
    ]
    if any(re.search(pattern, title, re.I) for pattern in title_patterns):
        return True

    haystack = f"{title}\n{detail}\n{fix}\n{verify}".lower()
    patterns = [
        r"\|\s*unable to verify\s*\]",
        r"\bunable to verify\b",
        r"\bmanual review\b",
        r"\bmanual check\b",
        r"\bmanual verification\b",
        r"\brequires manual\b",
        r"\brequired access\b",
        r"\bcannot confirm\b",
        r"\bcannot be confirmed\b",
        r"\bcould not be confirmed\b",
        r"\bcould not be proven\b",
        r"\bnot available via (?:mcp|api)\b",
        r"\badmin verification\b",
    ]
    return any(re.search(pattern, haystack) for pattern in patterns)


def summarize(text: str, limit: int = 220) -> str:
    collapsed = re.sub(r"\s+", " ", text).strip()
    if len(collapsed) <= limit:
        return collapsed
    return collapsed[: limit - 1].rstrip() + "..."


def load_rows() -> list[dict[str, str]]:
    workbook = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
    rows: list[dict[str, str]] = []
    for ws in workbook.worksheets:
        sheet_rows = ws.iter_rows(values_only=True)
        header = next(sheet_rows, None)
        if not header:
            continue
        header_map = {clean(name): index for index, name in enumerate(header) if clean(name)}
        required = [
            "Issue Name",
            "Detailed Explanation / Why It Matters",
            "How to Fix",
            "How to Verify",
            "Official Docs Used",
        ]
        if not all(name in header_map for name in required):
            continue
        for row_index, values in enumerate(sheet_rows, start=2):
            values = tuple(values or ())
            issue = clean(values[header_map["Issue Name"]] if header_map["Issue Name"] < len(values) else "")
            detail = clean(
                values[header_map["Detailed Explanation / Why It Matters"]]
                if header_map["Detailed Explanation / Why It Matters"] < len(values)
                else ""
            )
            fix = clean(values[header_map["How to Fix"]] if header_map["How to Fix"] < len(values) else "")
            verify = clean(values[header_map["How to Verify"]] if header_map["How to Verify"] < len(values) else "")
            docs = clean(values[header_map["Official Docs Used"]] if header_map["Official Docs Used"] < len(values) else "")
            if not any([issue, detail, fix, verify, docs]):
                continue
            issue_id = extract_issue_id(issue)
            rows.append(
                {
                    "id": f"{ws.title}-{row_index}-{issue_id or len(rows) + 1}",
                    "row": str(row_index),
                    "tool": ws.title,
                    "issueId": issue_id,
                    "status": extract_status(issue, detail, fix),
                    "manualReview": requires_manual_review(issue, detail, fix, verify),
                    "issue": issue,
                    "summary": summarize(detail or issue),
                    "detail": detail,
                    "fix": fix,
                    "verify": verify,
                    "docs": docs,
                }
            )
    workbook.close()
    return rows


def build_html(data: list[dict[str, str]]) -> str:
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    counts = {}
    for row in data:
        counts[row["tool"]] = counts.get(row["tool"], 0) + 1
    tool_counts = json.dumps(counts, ensure_ascii=False)

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Tharaa Audit Fixing Dashboard</title>
  <style>
    :root {{
      --c-primary: #6AE499;
      --c-dark: #0E1C26;
      --c-white: #FFFFFF;
      --c-muted: #8899AA;
      --c-light: #F4F7F6;
      --c-mid: #D9E5DF;
      --c-text: #1A2E22;
      --c-body: #2C3E35;
      --c-accent2: #EAF6EF;
      --c-accentDk: #4A6556;
      --bg: var(--c-light);
      --surface: var(--c-white);
      --surface-2: var(--c-accent2);
      --ink: var(--c-text);
      --muted: var(--c-muted);
      --line: var(--c-mid);
      --blue: var(--c-accentDk);
      --blue-soft: var(--c-accent2);
      --green: var(--c-accentDk);
      --green-soft: var(--c-accent2);
      --red: #C0392B;
      --red-soft: #FDECEA;
      --amber: #B7950B;
      --amber-soft: #FEFBD0;
      --soft-hover: #C7D7D0;
      --focus-ring: rgba(106, 228, 153, .28);
      --shadow: 0 14px 30px rgba(14, 28, 38, .08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      color: var(--c-body);
      background: var(--bg);
      line-height: 1.6;
    }}
    button, input, select, textarea {{ font: inherit; }}
    button {{ cursor: pointer; }}
    .app {{
      min-height: calc(100vh - 86px);
      display: grid;
      grid-template-columns: minmax(360px, 42%) 1fr;
    }}
    .doc-header {{
      padding: 12px 24px 10px;
      border-bottom: 2px solid var(--c-primary);
      display: flex;
      align-items: center;
      background: var(--c-white);
    }}
    .header-agency {{
      font-size: 13px;
      font-weight: 700;
      color: var(--c-primary);
      letter-spacing: 0.08em;
    }}
    .header-sep {{
      font-size: 13px;
      color: var(--c-mid);
      margin: 0 6px;
    }}
    .header-title {{
      font-size: 13px;
      color: var(--c-muted);
    }}
    .doc-footer {{
      padding: 12px 24px;
      border-top: 2px solid var(--c-primary);
      text-align: right;
      font-size: 11px;
      color: var(--c-muted);
      background: var(--c-white);
    }}
    .left {{
      border-right: 1px solid var(--line);
      background: var(--surface);
      min-height: calc(100vh - 86px);
      display: flex;
      flex-direction: column;
    }}
    .right {{
      min-height: calc(100vh - 86px);
      display: flex;
      flex-direction: column;
      background: var(--c-white);
    }}
    .topbar {{
      padding: 18px 22px 14px;
      border-bottom: 1px solid var(--line);
      background: var(--surface);
    }}
    h1 {{
      margin: 0 0 14px;
      font-size: 28px;
      font-weight: 700;
      color: var(--c-text);
      line-height: 1.2;
      letter-spacing: 0;
    }}
    .brand-subtitle {{
      margin: -8px 0 14px;
      font-size: 12px;
      font-style: italic;
      color: var(--c-muted);
    }}
    .kpis {{
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 8px;
      margin-bottom: 14px;
    }}
    .kpi {{
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px;
      background: var(--c-light);
      min-width: 0;
    }}
    .kpi strong {{
      display: block;
      font-size: 21px;
      line-height: 1.05;
      color: var(--c-text);
    }}
    .kpi span {{
      color: var(--muted);
      font-size: 12px;
    }}
    .filters {{
      display: grid;
      grid-template-columns: 1fr 170px 200px;
      gap: 10px;
    }}
    .search, .select {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      color: var(--ink);
      min-height: 42px;
      padding: 0 12px;
      outline: none;
    }}
    .search:focus, .select:focus {{
      border-color: var(--c-primary);
      box-shadow: 0 0 0 3px var(--focus-ring);
    }}
    .toolbar {{
      display: flex;
      gap: 8px;
      padding: 12px 22px;
      overflow-x: auto;
      border-bottom: 1px solid var(--line);
      background: var(--c-light);
    }}
    .tool-btn {{
      flex: 0 0 auto;
      border: 1px solid var(--line);
      background: var(--surface);
      color: var(--ink);
      border-radius: 8px;
      padding: 8px 10px;
      line-height: 1.1;
      white-space: nowrap;
    }}
    .tool-btn.active {{
      border-color: var(--c-primary);
      background: var(--c-dark);
      color: var(--c-primary);
      font-weight: 700;
    }}
    .list-head {{
      padding: 12px 22px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      color: var(--muted);
      font-size: 13px;
      border-bottom: 1px solid var(--line);
    }}
    .list-actions {{
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }}
    .list {{
      overflow: auto;
      padding: 12px;
      flex: 1;
    }}
    .point {{
      width: 100%;
      display: block;
      text-align: left;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 13px;
      margin-bottom: 10px;
      background: var(--surface);
      color: var(--ink);
      box-shadow: none;
    }}
    .point:hover {{ border-color: var(--soft-hover); }}
    .point.active {{
      border-color: var(--c-primary);
      box-shadow: 0 0 0 3px var(--focus-ring);
    }}
    .point-title {{
      font-size: 14px;
      line-height: 1.35;
      font-weight: 700;
      margin-bottom: 8px;
    }}
    .meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      align-items: center;
    }}
    .chip {{
      display: inline-flex;
      align-items: center;
      min-height: 24px;
      border-radius: 999px;
      padding: 3px 8px;
      font-size: 12px;
      border: 1px solid var(--line);
      background: var(--c-light);
      color: var(--muted);
      white-space: nowrap;
    }}
    .chip.fixed {{ color: var(--green); background: var(--green-soft); border-color: var(--c-primary); }}
    .chip.notfixed {{ color: var(--red); background: var(--red-soft); border-color: #f2b8b5; }}
    .chip.unreviewed {{ color: var(--amber); background: var(--amber-soft); border-color: #f0d18c; }}
    .chip.tool {{ color: var(--c-primary); background: var(--c-dark); border-color: var(--c-primary); }}
    .chip.manual {{ color: var(--amber); background: var(--amber-soft); border-color: #f0d18c; font-weight: 700; }}
    .detail-top {{
      position: sticky;
      top: 0;
      z-index: 3;
      padding: 18px 24px;
      border-bottom: 1px solid var(--line);
      background: rgba(255, 255, 255, .96);
      backdrop-filter: blur(8px);
    }}
    .detail-title {{
      margin: 0 0 10px;
      font-size: 24px;
      color: var(--c-text);
      line-height: 1.25;
      letter-spacing: 0;
    }}
    .actions {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
      margin-top: 14px;
    }}
    .action-btn {{
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
      color: var(--ink);
      min-height: 40px;
      padding: 0 12px;
      font-weight: 700;
    }}
    .action-btn.fixed {{ border-color: #98d4b5; color: var(--green); background: var(--green-soft); }}
    .action-btn.notfixed {{ border-color: #f2b8b5; color: var(--red); background: var(--red-soft); }}
    .action-btn.ghost {{ color: var(--muted); }}
    .action-btn.manual-filter {{
      border-color: #f0d18c;
      color: var(--amber);
      background: var(--amber-soft);
    }}
    .action-btn.manual-filter.active {{
      border-color: var(--amber);
      box-shadow: 0 0 0 3px var(--focus-ring);
    }}
    .detail-body {{
      padding: 20px 24px 40px;
      overflow: auto;
    }}
    .section {{
      margin-bottom: 18px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      box-shadow: var(--shadow);
    }}
    .section h2 {{
      margin: 0;
      padding: 13px 15px;
      font-size: 13px;
      color: var(--c-primary);
      text-transform: uppercase;
      letter-spacing: 0.03em;
      border-bottom: 1px solid var(--line);
      background: var(--c-dark);
    }}
    .section .content {{
      padding: 15px;
      white-space: pre-wrap;
      line-height: 1.55;
      font-size: 13px;
      color: var(--c-body);
      overflow-wrap: anywhere;
    }}
    .docs a {{
      display: block;
      color: var(--blue);
      margin-bottom: 8px;
      overflow-wrap: anywhere;
    }}
    .empty {{
      margin: auto;
      max-width: 440px;
      text-align: center;
      color: var(--muted);
      padding: 30px;
    }}
    .note-box {{
      width: 100%;
      min-height: 90px;
      resize: vertical;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px;
      color: var(--ink);
      background: var(--c-white);
    }}
    .small {{
      font-size: 12px;
      color: var(--muted);
    }}
    @media (max-width: 980px) {{
      .app {{ grid-template-columns: 1fr; }}
      .left, .right {{ min-height: auto; }}
      .right {{ border-top: 1px solid var(--line); }}
      .kpis {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
    @media (max-width: 560px) {{
      .topbar, .detail-top, .detail-body {{ padding-left: 14px; padding-right: 14px; }}
      .toolbar, .list-head {{ padding-left: 14px; padding-right: 14px; }}
      .list-head {{ align-items: flex-start; gap: 10px; flex-direction: column; }}
      .list-actions {{ justify-content: flex-start; }}
      .filters {{ grid-template-columns: 1fr; }}
      .detail-title {{ font-size: 19px; }}
    }}
  </style>
</head>
<body>
  <header class="doc-header">
    <span class="header-agency">OPTIMIZERS</span>
    <span class="header-sep"> | </span>
    <span class="header-title">Tharaa Audit Fixing Dashboard</span>
  </header>
  <main class="app">
    <section class="left" aria-label="Audit points">
      <div class="topbar">
        <h1>Tharaa Audit Fixing Dashboard</h1>
        <div class="brand-subtitle">We make optimization Simple, Practical, and Profitable</div>
        <div class="kpis" aria-label="Audit totals">
          <div class="kpi"><strong id="totalCount">0</strong><span>Total audit points</span></div>
          <div class="kpi"><strong id="visibleCount">0</strong><span>Currently shown</span></div>
          <div class="kpi"><strong id="manualCount">0</strong><span>Manual review</span></div>
          <div class="kpi"><strong id="fixedCount">0</strong><span>Marked fixed</span></div>
          <div class="kpi"><strong id="notFixedCount">0</strong><span>Marked not fixed</span></div>
        </div>
        <div class="filters">
          <input id="search" class="search" type="search" placeholder="Search audit points, IDs, fixes, docs">
          <select id="statusFilter" class="select" aria-label="Status filter">
            <option value="all">All statuses</option>
            <option value="unreviewed">Not reviewed</option>
            <option value="fixed">Fixed</option>
            <option value="notfixed">Not fixed</option>
          </select>
          <select id="reviewFilter" class="select" aria-label="Review type filter">
            <option value="all">All review types</option>
            <option value="manual">Manual review needed</option>
            <option value="nonmanual">No manual review flag</option>
          </select>
        </div>
      </div>
      <nav class="toolbar" id="toolbar" aria-label="Tool filter"></nav>
      <div class="list-head">
        <span id="resultLabel">All audit points</span>
        <div class="list-actions">
          <button class="action-btn manual-filter" id="manualOnly" type="button">Manual review only</button>
          <button class="action-btn ghost" id="clearFilters" type="button">Clear filters</button>
        </div>
      </div>
      <div class="list" id="list"></div>
    </section>

    <section class="right" aria-label="Selected audit point">
      <div id="detail"></div>
    </section>
  </main>
  <footer class="doc-footer">
    <span>© Optimizers Agency &nbsp;|&nbsp; Confidential Client Dashboard</span>
  </footer>

  <script type="application/json" id="audit-data">{payload}</script>
  <script>
    const auditData = JSON.parse(document.getElementById("audit-data").textContent);
    const toolCounts = {tool_counts};
    const storeKey = "tharaa-audit-command-center-progress-v1";
    const state = {{
      tool: "All",
      status: "all",
      review: "all",
      query: "",
      selectedId: auditData[0]?.id || null,
      progress: loadProgress(),
    }};

    const els = {{
      toolbar: document.getElementById("toolbar"),
      list: document.getElementById("list"),
      detail: document.getElementById("detail"),
      search: document.getElementById("search"),
      statusFilter: document.getElementById("statusFilter"),
      totalCount: document.getElementById("totalCount"),
      visibleCount: document.getElementById("visibleCount"),
      manualCount: document.getElementById("manualCount"),
      fixedCount: document.getElementById("fixedCount"),
      notFixedCount: document.getElementById("notFixedCount"),
      resultLabel: document.getElementById("resultLabel"),
      clearFilters: document.getElementById("clearFilters"),
      reviewFilter: document.getElementById("reviewFilter"),
      manualOnly: document.getElementById("manualOnly"),
    }};

    function loadProgress() {{
      try {{
        return JSON.parse(localStorage.getItem(storeKey) || "{{}}");
      }} catch {{
        return {{}};
      }}
    }}

    function saveProgress() {{
      localStorage.setItem(storeKey, JSON.stringify(state.progress));
    }}

    function progressFor(id) {{
      return state.progress[id] || {{ status: "unreviewed", note: "" }};
    }}

    function setProgress(id, status) {{
      state.progress[id] = {{ ...progressFor(id), status, updatedAt: new Date().toISOString() }};
      saveProgress();
      render();
    }}

    function setNote(id, note) {{
      state.progress[id] = {{ ...progressFor(id), note, updatedAt: new Date().toISOString() }};
      saveProgress();
      renderCounts(filteredRows().length);
    }}

    function escapeHtml(value) {{
      return String(value || "").replace(/[&<>"']/g, char => ({{
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }}[char]));
    }}

    function statusLabel(status) {{
      if (status === "fixed") return "Fixed";
      if (status === "notfixed") return "Not fixed";
      return "Not reviewed";
    }}

    function docsHtml(text) {{
      const urls = String(text || "").split(/\\s+/).filter(Boolean);
      if (!urls.length) return '<span class="small">No official docs listed in the workbook.</span>';
      return urls.map(url => {{
        const safe = escapeHtml(url);
        if (/^https?:\\/\\//i.test(url)) return `<a href="${{safe}}" target="_blank" rel="noopener noreferrer">${{safe}}</a>`;
        return `<div>${{safe}}</div>`;
      }}).join("");
    }}

    function filteredRows() {{
      const q = state.query.trim().toLowerCase();
      return auditData.filter(row => {{
        const progress = progressFor(row.id);
        if (state.tool !== "All" && row.tool !== state.tool) return false;
        if (state.status !== "all" && progress.status !== state.status) return false;
        if (state.review === "manual" && !row.manualReview) return false;
        if (state.review === "nonmanual" && row.manualReview) return false;
        if (!q) return true;
        return [
          row.issue,
          row.issueId,
          row.tool,
          row.status,
          row.summary,
          row.fix,
          row.verify,
          row.docs,
        ].join("\\n").toLowerCase().includes(q);
      }});
    }}

    function renderToolbar() {{
      const tools = ["All", ...Object.keys(toolCounts)];
      els.toolbar.innerHTML = "";
      for (const tool of tools) {{
        const button = document.createElement("button");
        button.type = "button";
        button.className = `tool-btn${{state.tool === tool ? " active" : ""}}`;
        button.textContent = tool === "All" ? `All (${{auditData.length.toLocaleString()}})` : `${{tool}} (${{toolCounts[tool].toLocaleString()}})`;
        button.addEventListener("click", () => {{
          state.tool = tool;
          render();
        }});
        els.toolbar.appendChild(button);
      }}
    }}

    function renderCounts(visibleTotal) {{
      let fixed = 0;
      let notFixed = 0;
      let manual = 0;
      for (const row of auditData) {{
        const status = progressFor(row.id).status;
        if (status === "fixed") fixed += 1;
        if (status === "notfixed") notFixed += 1;
        if (row.manualReview) manual += 1;
      }}
      els.totalCount.textContent = auditData.length.toLocaleString();
      els.visibleCount.textContent = visibleTotal.toLocaleString();
      els.manualCount.textContent = manual.toLocaleString();
      els.fixedCount.textContent = fixed.toLocaleString();
      els.notFixedCount.textContent = notFixed.toLocaleString();
    }}

    function renderList(rows) {{
      els.list.innerHTML = "";
      if (!rows.length) {{
        els.list.innerHTML = '<div class="empty">No audit points match the current filters.</div>';
        return;
      }}
      const fragment = document.createDocumentFragment();
      for (const row of rows) {{
        const progress = progressFor(row.id);
        const button = document.createElement("button");
        button.type = "button";
        button.className = `point${{row.id === state.selectedId ? " active" : ""}}`;
        button.innerHTML = `
          <div class="point-title">${{escapeHtml(row.issue)}}</div>
          <div class="meta">
            <span class="chip tool">${{escapeHtml(row.tool)}}</span>
            ${{row.issueId ? `<span class="chip">${{escapeHtml(row.issueId)}}</span>` : ""}}
            ${{row.manualReview ? '<span class="chip manual">Manual review</span>' : ""}}
            <span class="chip ${{progress.status}}">${{statusLabel(progress.status)}}</span>
            <span class="chip">${{escapeHtml(row.status)}}</span>
          </div>`;
        button.addEventListener("click", () => {{
          state.selectedId = row.id;
          render();
        }});
        fragment.appendChild(button);
      }}
      els.list.appendChild(fragment);
    }}

    function selectedRow(rows) {{
      let row = auditData.find(item => item.id === state.selectedId);
      if (!row || (rows.length && !rows.some(item => item.id === row.id))) {{
        row = rows[0] || auditData[0];
        state.selectedId = row?.id || null;
      }}
      return row;
    }}

    function renderDetail(row) {{
      if (!row) {{
        els.detail.innerHTML = '<div class="empty">Select an audit point to see the fix.</div>';
        return;
      }}
      const progress = progressFor(row.id);
      els.detail.innerHTML = `
        <div class="detail-top">
          <h2 class="detail-title">${{escapeHtml(row.issue)}}</h2>
          <div class="meta">
            <span class="chip tool">${{escapeHtml(row.tool)}}</span>
            ${{row.issueId ? `<span class="chip">${{escapeHtml(row.issueId)}}</span>` : ""}}
            ${{row.manualReview ? '<span class="chip manual">Manual review required</span>' : ""}}
            <span class="chip">${{escapeHtml(row.status)}}</span>
            <span class="chip ${{progress.status}}">${{statusLabel(progress.status)}}</span>
          </div>
          <div class="actions">
            <button class="action-btn fixed" type="button" data-action="fixed">Mark fixed</button>
            <button class="action-btn notfixed" type="button" data-action="notfixed">Mark not fixed</button>
            <button class="action-btn ghost" type="button" data-action="unreviewed">Reset status</button>
          </div>
        </div>
        <div class="detail-body">
          <div class="section">
            <h2>How To Fix</h2>
            <div class="content">${{escapeHtml(row.fix || "No fix steps were listed in the workbook.")}}</div>
          </div>
          <div class="section">
            <h2>How To Verify</h2>
            <div class="content">${{escapeHtml(row.verify || "No verification steps were listed in the workbook.")}}</div>
          </div>
          <div class="section">
            <h2>Why It Matters</h2>
            <div class="content">${{escapeHtml(row.detail || row.summary)}}</div>
          </div>
          <div class="section">
            <h2>Official Docs</h2>
            <div class="content docs">${{docsHtml(row.docs)}}</div>
          </div>
          <div class="section">
            <h2>My Notes</h2>
            <div class="content">
              <textarea class="note-box" id="noteBox" placeholder="Add evidence links, screenshot names, or what changed...">${{escapeHtml(progress.note || "")}}</textarea>
              <div class="small">Saved automatically in this browser.</div>
            </div>
          </div>
        </div>`;

      els.detail.querySelectorAll("[data-action]").forEach(button => {{
        button.addEventListener("click", () => setProgress(row.id, button.dataset.action));
      }});
      const noteBox = document.getElementById("noteBox");
      noteBox.addEventListener("input", event => setNote(row.id, event.target.value));
    }}

    function render() {{
      renderToolbar();
      const rows = filteredRows();
      renderCounts(rows.length);
      els.resultLabel.textContent = `${{rows.length.toLocaleString()}} audit point${{rows.length === 1 ? "" : "s"}} shown`;
      els.manualOnly.classList.toggle("active", state.review === "manual");
      els.manualOnly.textContent = state.review === "manual" ? "Showing manual review" : "Manual review only";
      renderList(rows);
      renderDetail(selectedRow(rows));
    }}

    els.search.addEventListener("input", event => {{
      state.query = event.target.value;
      render();
    }});
    els.statusFilter.addEventListener("change", event => {{
      state.status = event.target.value;
      render();
    }});
    els.reviewFilter.addEventListener("change", event => {{
      state.review = event.target.value;
      render();
    }});
    els.manualOnly.addEventListener("click", () => {{
      state.review = state.review === "manual" ? "all" : "manual";
      els.reviewFilter.value = state.review;
      render();
    }});
    els.clearFilters.addEventListener("click", () => {{
      state.tool = "All";
      state.status = "all";
      state.review = "all";
      state.query = "";
      els.search.value = "";
      els.statusFilter.value = "all";
      els.reviewFilter.value = "all";
      render();
    }});

    render();
  </script>
</body>
</html>
"""


def main() -> None:
    data = load_rows()
    OUTPUT.write_text(build_html(data), encoding="utf-8")
    print(json.dumps({"output": str(OUTPUT), "rows": len(data)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
