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
        r"C:\Users\Omar Maged\Downloads\Dashboard data sheet.xlsx",
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


def slug(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-")
    return cleaned or "row"


def find_header(rows: list[tuple[int, tuple[object, ...]]]) -> tuple[int, dict[str, int]] | None:
    for position, (_excel_row, values) in enumerate(rows[:30]):
        names = [clean(value) for value in values]
        if "Issue Name" in names and ("Tab" in names or "Original Row" in names):
            return position, {name: index for index, name in enumerate(names) if name}
    return None


def value_for(values: tuple[object, ...], header_map: dict[str, int], *names: str) -> str:
    for name in names:
        index = header_map.get(name)
        if index is not None and index < len(values):
            value = clean(values[index])
            if value:
                return value
    return ""


BUSINESS_GLOSSARY = [
    (r"\bGA4\b", "analytics"),
    (r"\bGoogle Analytics 4\b", "analytics"),
    (r"\bGTM\b", "tracking manager"),
    (r"\bGoogle Tag Manager\b", "tracking manager"),
    (r"\bGoogle Search Console\b", "Google search performance tool"),
    (r"\bGSC\b", "Google search performance tool"),
    (r"\bMicrosoft Clarity\b", "customer behavior tool"),
    (r"\bClarity\b", "customer behavior tool"),
    (r"\bdataLayer\b", "tracking data layer"),
    (r"\bkey event[s]?\b", "important business action"),
    (r"\bconversion[s]?\b", "important business result"),
    (r"\becommerce\b", "online store"),
    (r"\battribution\b", "which marketing source gets credit"),
    (r"\bSmart Bidding\b", "automated ad bidding"),
    (r"\bparameter[s]?\b", "extra tracking details"),
    (r"\bdimension[s]?\b", "reporting field"),
    (r"\bmetric[s]?\b", "reporting number"),
    (r"\bevents\b", "tracked actions"),
    (r"\bevent\b", "tracked action"),
    (r"\btag[s]?\b", "tracking code"),
    (r"\btrigger[s]?\b", "rule that decides when tracking runs"),
    (r"\bcontainer\b", "tracking workspace"),
    (r"\bDebugView\b", "live testing screen"),
    (r"\bTag Assistant\b", "Google tracking checker"),
    (r"\bURL Inspection\b", "Google page check"),
    (r"\bcanonical\b", "preferred page version"),
    (r"\bindexing\b", "Google saving pages for search results"),
    (r"\bcrawl(?:ing)?\b", "Google reading the site"),
    (r"\bstructured data\b", "search-friendly page information"),
    (r"\bconsent mode\b", "privacy consent setup"),
    (r"\bCMP\b", "cookie/privacy banner"),
    (r"\bPII\b", "personal customer information"),
    (r"\bBigQuery\b", "raw data warehouse"),
    (r"\bAPI\b", "system connection"),
    (r"\bMCP\b", "system access tool"),
    (r"\bSKU\b", "product code"),
    (r"\bROAS\b", "ad return"),
    (r"\bCPA\b", "cost per result"),
    (r"\bCTR\b", "click rate"),
    (r"\bCWV\b", "page speed and experience checks"),
    (r"\bVAT\b", "sales tax"),
    (r"\bPDPL\b", "privacy law"),
]


def strip_issue_prefix(text: str) -> str:
    text = re.sub(r"^\[[^\]]+\]\s*", "", text).strip()
    text = re.sub(r"^[A-Z0-9-]+\s*[—-]\s*", "", text).strip()
    return text


def plain_language(text: str) -> str:
    text = clean(text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = text.replace("page_view", "page view").replace("session_start", "session start")
    text = text.replace("itemCategory2", "product subcategory")
    text = text.replace("itemCategory3", "product subcategory")
    text = text.replace("itemCategory", "product category")
    text = text.replace("itemListName", "product list name")
    text = text.replace("itemListId", "product list ID")
    text = text.replace("contentGroup", "page group")
    text = text.replace("→", " to ").replace("≥", "at least ").replace("×", " times ")
    text = text.replace("100%", "all")
    for pattern, replacement in BUSINESS_GLOSSARY:
        text = re.sub(pattern, replacement, text, flags=re.I)
    text = re.sub(r"\ba analytics\b", "an analytics", text)
    text = re.sub(r"\ba important\b", "an important", text)
    text = re.sub(r"\ban analytics important business action\b", "an important business action in analytics", text)
    text = re.sub(r"\bmarked as an important business action in analytics\b", "being counted as a business success action", text)
    text = re.sub(r"\bis an important business action in analytics\b", "is counted as a business success action", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def business_explanation(text: str) -> str:
    text = plain_language(text)
    text = re.sub(
        r"(?is)\n\s*(Recommended action|Verification standard|How to verify|How to fix):.*$",
        "",
        text,
    )
    text = re.sub(r"(?im)^\s*(Explanation|Why it matters|Business impact|Impact):\s*", "", text)
    text = re.sub(
        r"(?is)This row describes a real configuration or data-quality problem that needs to be reviewed before it is marked fixed\.?",
        "",
        text,
    )
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def short_business_issue(issue: str, tool: str) -> str:
    text = plain_language(strip_issue_prefix(issue))
    text = re.sub(r"\s+", " ", text)
    if len(text) > 180:
        text = text[:177].rstrip() + "..."
    return text


def business_detail(issue: str, detail: str, tool: str, review_status: str) -> str:
    issue_text = plain_language(strip_issue_prefix(issue))
    detail_text = business_explanation(detail or issue)
    if detail_text and detail_text.lower() != issue_text.lower():
        return f"{issue_text}\n\n{detail_text}"
    return issue_text


def business_fix(fix: str, tool: str) -> str:
    fix_text = plain_language(fix)
    if not fix_text:
        fix_text = "Confirm the current setup, decide the correct business outcome, make the change in the relevant platform, and save before/after proof."
    return (
        "What to do in business terms:\n"
        f"{fix_text}\n\n"
        "Keep it simple: confirm the problem is still true, fix the setting or tracking source, then record proof so the team can trust the result later."
    )


def business_verify(verify: str) -> str:
    verify_text = plain_language(verify)
    if not verify_text:
        verify_text = "Open the relevant platform, check the final setting or report, and save evidence."
    return (
        "How to know it is fixed:\n"
        f"{verify_text}\n\n"
        "Do not mark this complete from memory. Mark it complete only after you see proof and save the evidence."
    )


def load_rows() -> list[dict[str, str]]:
    workbook = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
    rows: list[dict[str, str]] = []
    for ws in workbook.worksheets:
        if ws.title in {"Cleanup Log", "Official Docs"} or ws.title.startswith("_"):
            continue

        raw_rows = [
            (index, tuple(values or ()))
            for index, values in enumerate(ws.iter_rows(values_only=True), start=1)
            if any(clean(value) for value in (values or ()))
        ]
        header_result = find_header(raw_rows)
        if not header_result:
            continue

        header_position, header_map = header_result
        for row_index, values in raw_rows[header_position + 1 :]:
            issue = value_for(values, header_map, "Issue Name")
            tool = value_for(values, header_map, "Tab") or ws.title
            original_row = value_for(values, header_map, "Original Row") or str(row_index)
            review_status = value_for(values, header_map, "Review Status", "Backlog Category")
            routing_note = value_for(values, header_map, "Reason / Routing Note")
            owner = value_for(values, header_map, "Owner")
            qa_outcome = value_for(values, header_map, "QA Outcome", "Decision")
            evidence_link = value_for(values, header_map, "Evidence Link")
            detail = value_for(
                values,
                header_map,
                "Detailed Explanation / Why It Matters",
                "QA Task / Verification Step",
                "Recommended Next Step",
            )
            fix = value_for(
                values,
                header_map,
                "Rewritten / Routed Fix",
                "Recommended Next Step",
                "QA Task / Verification Step",
                "Original How to Fix",
                "How to Fix",
            )
            verify = value_for(values, header_map, "How to Verify")
            docs = value_for(values, header_map, "Official Docs Used", "Official Basis")
            if not any([issue, detail, fix, verify, docs]):
                continue
            issue_id = extract_issue_id(issue)
            row_key = slug(f"{ws.title}-{tool}-{original_row}-{issue_id or issue[:80]}")
            rows.append(
                {
                    "id": row_key,
                    "row": original_row,
                    "sourceRow": str(row_index),
                    "sourceTab": ws.title,
                    "tool": tool,
                    "issueId": issue_id,
                    "status": extract_status(issue, detail, fix),
                    "manualReview": requires_manual_review(issue, detail, fix, verify),
                    "issue": issue,
                    "summary": summarize(detail or issue),
                    "detail": detail,
                    "fix": fix,
                    "verify": verify,
                    "docs": docs,
                    "reviewStatus": review_status,
                    "routingNote": routing_note,
                    "owner": owner,
                    "qaOutcome": qa_outcome,
                    "evidenceLink": evidence_link,
                    "businessIssue": short_business_issue(issue, tool),
                    "businessDetail": business_detail(issue, detail, tool, review_status),
                    "businessFix": business_fix(fix, tool),
                    "businessVerify": business_verify(verify),
                }
            )
    workbook.close()
    return rows


def build_html(data: list[dict[str, str]]) -> str:
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    counts = {}
    tab_counts = {}
    for row in data:
        counts[row["tool"]] = counts.get(row["tool"], 0) + 1
        tab_counts[row["sourceTab"]] = tab_counts.get(row["sourceTab"], 0) + 1
    tool_counts = json.dumps(counts, ensure_ascii=False)
    source_counts = json.dumps(tab_counts, ensure_ascii=False)

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
    html {{
      width: 100%;
      -webkit-text-size-adjust: 100%;
      text-size-adjust: 100%;
    }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      color: var(--c-body);
      background: var(--bg);
      line-height: 1.6;
      overflow-x: hidden;
    }}
    button, input, select, textarea {{ font: inherit; }}
    button {{ cursor: pointer; }}
    .app {{
      width: 100%;
      max-width: 100vw;
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
      min-width: 0;
    }}
    .right {{
      min-height: calc(100vh - 86px);
      display: flex;
      flex-direction: column;
      background: var(--c-white);
      min-width: 0;
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
    .completion-panel {{
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      margin-bottom: 14px;
      background: var(--c-white);
    }}
    .completion-head {{
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 8px;
    }}
    .completion-title {{
      color: var(--c-text);
      font-size: 13px;
      font-weight: 700;
    }}
    .completion-percent {{
      color: var(--c-dark);
      font-size: 20px;
      line-height: 1;
      font-weight: 700;
    }}
    .completion-track {{
      height: 12px;
      border-radius: 999px;
      overflow: hidden;
      background: var(--c-light);
      border: 1px solid var(--line);
    }}
    .completion-fill {{
      width: 0%;
      height: 100%;
      border-radius: inherit;
      background: linear-gradient(90deg, var(--c-primary), var(--c-accentDk));
      transition: width .25s ease;
    }}
    .completion-meta {{
      margin-top: 7px;
      color: var(--muted);
      font-size: 12px;
    }}
    .mode-switch {{
      display: inline-grid;
      grid-template-columns: 1fr 1fr;
      gap: 4px;
      padding: 4px;
      margin-bottom: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--c-light);
    }}
    .mode-btn {{
      border: 0;
      border-radius: 6px;
      min-height: 34px;
      padding: 0 12px;
      background: transparent;
      color: var(--muted);
      font-weight: 700;
    }}
    .mode-btn.active {{
      background: var(--c-dark);
      color: var(--c-primary);
    }}
    .filters {{
      display: grid;
      grid-template-columns: 1fr 170px 200px 220px;
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
    .action-btn.danger {{ border-color: #f2b8b5; color: var(--red); background: #fff7f6; }}
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
      .app {{
        display: block;
        min-height: auto;
      }}
      .left, .right {{
        width: 100%;
        min-height: auto;
      }}
      .left {{ border-right: 0; }}
      .right {{
        border-top: 1px solid var(--line);
        scroll-margin-top: 0;
      }}
      .kpis {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .list {{
        max-height: 56vh;
        overflow: auto;
        overscroll-behavior: contain;
      }}
    }}
    @media (max-width: 560px) {{
      .doc-header {{
        padding: 10px 14px;
        flex-wrap: wrap;
        row-gap: 2px;
      }}
      .header-title {{ font-size: 12px; }}
      .doc-footer {{
        padding: 10px 14px;
        text-align: left;
      }}
      .topbar, .detail-top, .detail-body {{ padding-left: 14px; padding-right: 14px; }}
      .topbar {{ padding-top: 14px; padding-bottom: 12px; }}
      h1 {{
        font-size: 22px;
        margin-bottom: 10px;
      }}
      .brand-subtitle {{
        margin-top: -4px;
        font-size: 11px;
      }}
      .kpis {{
        gap: 7px;
        margin-bottom: 12px;
      }}
      .kpi {{ padding: 8px; }}
      .kpi strong {{ font-size: 18px; }}
      .kpi span {{
        display: block;
        font-size: 11px;
        line-height: 1.25;
      }}
      .completion-panel {{
        padding: 10px;
        margin-bottom: 12px;
      }}
      .completion-title {{ font-size: 12px; }}
      .completion-percent {{ font-size: 18px; }}
      .mode-switch {{
        display: grid;
        width: 100%;
      }}
      .mode-btn {{ min-height: 40px; }}
      .toolbar, .list-head {{ padding-left: 14px; padding-right: 14px; }}
      .list-head {{ align-items: flex-start; gap: 10px; flex-direction: column; }}
      .list-actions {{ justify-content: flex-start; }}
      .filters {{ grid-template-columns: 1fr; }}
      .search, .select {{
        min-height: 44px;
        font-size: 16px;
      }}
      .toolbar {{
        gap: 7px;
        padding-top: 10px;
        padding-bottom: 10px;
      }}
      .tool-btn {{
        padding: 8px 9px;
        font-size: 13px;
      }}
      .list {{
        max-height: 48vh;
        padding: 10px 12px 14px;
      }}
      .point {{
        padding: 11px;
        margin-bottom: 8px;
      }}
      .point-title {{ font-size: 13px; }}
      .chip {{
        max-width: 100%;
        white-space: normal;
        font-size: 11px;
        line-height: 1.2;
      }}
      .detail-top {{
        position: static;
        padding-top: 14px;
        padding-bottom: 14px;
      }}
      .detail-title {{ font-size: 18px; }}
      .actions {{
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }}
      .action-btn {{
        width: 100%;
        min-height: 44px;
        padding: 0 10px;
        font-size: 13px;
      }}
      .detail-body {{
        padding-top: 14px;
        padding-bottom: 28px;
      }}
      .section {{ margin-bottom: 14px; }}
      .section h2 {{
        padding: 11px 12px;
        font-size: 11px;
      }}
      .section .content {{
        padding: 12px;
        font-size: 14px;
        line-height: 1.5;
      }}
    }}
    @media (max-width: 380px) {{
      .actions {{ grid-template-columns: 1fr; }}
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
          <div class="kpi"><strong id="cloudStatus">Local</strong><span>Firestore sync</span></div>
        </div>
        <div class="completion-panel" aria-label="Overall completion">
          <div class="completion-head">
            <span class="completion-title">Overall completion</span>
            <strong class="completion-percent" id="completionPercent">0%</strong>
          </div>
          <div class="completion-track" role="progressbar" aria-label="Fixed audit point completion" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
            <div class="completion-fill" id="completionFill"></div>
          </div>
          <div class="completion-meta" id="completionMeta">0 of 0 tasks fixed</div>
        </div>
        <div class="mode-switch" aria-label="Language mode">
          <button class="mode-btn active" id="technicalMode" type="button">Technical</button>
          <button class="mode-btn" id="businessMode" type="button">Business</button>
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
          <select id="sourceFilter" class="select" aria-label="Source tab filter">
            <option value="All">All workbook tabs</option>
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
  <script type="module">
    import {{ initializeApp }} from "https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js";
    import {{
      collection,
      doc,
      getDocs,
      getFirestore,
      serverTimestamp,
      setDoc,
    }} from "https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js";

    const firebaseConfig = {{
      apiKey: "AIzaSyCFF5gMTyp3ibY1ufEQfVbW1AUiDTH7ZQY",
      authDomain: "dashboard-fb7e9.firebaseapp.com",
      projectId: "dashboard-fb7e9",
      storageBucket: "dashboard-fb7e9.firebasestorage.app",
      messagingSenderId: "577559586975",
      appId: "1:577559586975:web:616d8e3699a91210d89aaa",
    }};

    const auditData = JSON.parse(document.getElementById("audit-data").textContent);
    const toolCounts = {tool_counts};
    const sourceCounts = {source_counts};
    const storeKey = "tharaa-audit-command-center-progress-v1";
    const deletedStoreKey = "tharaa-audit-command-center-deleted-v1";
    const firebaseApp = initializeApp(firebaseConfig);
    const db = getFirestore(firebaseApp);
    const progressCollection = collection(db, "dashboards", "tharaa-audit-fixing-dashboard", "progress");
    const cloudSaveTimers = new Map();
    const state = {{
      tool: "All",
      status: "all",
      review: "all",
      sourceTab: "All",
      businessMode: false,
      query: "",
      selectedId: auditData[0]?.id || null,
      progress: loadProgress(),
      deletedIds: new Set(loadDeletedIds()),
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
      cloudStatus: document.getElementById("cloudStatus"),
      completionPercent: document.getElementById("completionPercent"),
      completionFill: document.getElementById("completionFill"),
      completionMeta: document.getElementById("completionMeta"),
      resultLabel: document.getElementById("resultLabel"),
      clearFilters: document.getElementById("clearFilters"),
      reviewFilter: document.getElementById("reviewFilter"),
      sourceFilter: document.getElementById("sourceFilter"),
      technicalMode: document.getElementById("technicalMode"),
      businessMode: document.getElementById("businessMode"),
      manualOnly: document.getElementById("manualOnly"),
    }};

    function populateSourceFilter() {{
      const fragment = document.createDocumentFragment();
      for (const [sourceTab, count] of Object.entries(sourceCounts)) {{
        const option = document.createElement("option");
        option.value = sourceTab;
        option.textContent = `${{sourceTab}} (${{count.toLocaleString()}})`;
        fragment.appendChild(option);
      }}
      els.sourceFilter.appendChild(fragment);
    }}

    function loadProgress() {{
      try {{
        return JSON.parse(localStorage.getItem(storeKey) || "{{}}");
      }} catch {{
        return {{}};
      }}
    }}

    function loadDeletedIds() {{
      try {{
        const saved = JSON.parse(localStorage.getItem(deletedStoreKey) || "[]");
        if (!Array.isArray(saved)) return [];
        const knownIds = new Set(auditData.map(row => row.id));
        const validIds = [...new Set(saved)].filter(id => knownIds.has(id));
        if (validIds.length >= auditData.length) {{
          console.warn("Deleted point list hid every audit point; resetting it.");
          localStorage.removeItem(deletedStoreKey);
          return [];
        }}
        return validIds;
      }} catch {{
        return [];
      }}
    }}

    function saveDeletedIds() {{
      localStorage.setItem(deletedStoreKey, JSON.stringify([...state.deletedIds]));
    }}

    function activeRows() {{
      return auditData.filter(row => !state.deletedIds.has(row.id));
    }}

    function setCloudStatus(message) {{
      els.cloudStatus.textContent = message;
    }}

    function saveLocalProgress() {{
      localStorage.setItem(storeKey, JSON.stringify(state.progress));
    }}

    function saveProgress(id) {{
      saveLocalProgress();
      if (id) queueCloudSave(id);
    }}

    function progressFor(id) {{
      return state.progress[id] || {{ status: "unreviewed", note: "" }};
    }}

    function setProgress(id, status) {{
      state.progress[id] = {{ ...progressFor(id), status, updatedAt: new Date().toISOString() }};
      saveProgress(id);
      render();
    }}

    function setNote(id, note) {{
      state.progress[id] = {{ ...progressFor(id), note, updatedAt: new Date().toISOString() }};
      saveProgress(id);
      renderCounts(filteredRows().length);
    }}

    function deletePoint(id) {{
      const row = auditData.find(item => item.id === id);
      if (!row) return;
      const confirmed = window.confirm("Delete this audit point from your dashboard? It will be removed from lists and totals for everyone using the synced dashboard.");
      if (!confirmed) return;
      state.deletedIds.add(id);
      state.progress[id] = {{
        ...progressFor(id),
        deleted: true,
        deletedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      }};
      saveDeletedIds();
      saveProgress(id);
      state.selectedId = activeRows()[0]?.id || null;
      render();
    }}

    function progressPayload(id) {{
      const row = auditData.find(item => item.id === id) || {{}};
      const progress = progressFor(id);
      return {{
        pointId: id,
        row: row.row || "",
        tool: row.tool || "",
        issueId: row.issueId || "",
        auditStatus: row.status || "",
        sourceTab: row.sourceTab || "",
        sourceRow: row.sourceRow || "",
        issue: row.issue || "",
        manualReview: Boolean(row.manualReview),
        status: progress.status || "unreviewed",
        fixed: progress.status === "fixed",
        note: progress.note || "",
        deleted: state.deletedIds.has(id) || progress.deleted === true,
        deletedAt: progress.deletedAt || "",
        updatedAtIso: progress.updatedAt || new Date().toISOString(),
        updatedAt: serverTimestamp(),
      }};
    }}

    function queueCloudSave(id) {{
      clearTimeout(cloudSaveTimers.get(id));
      cloudSaveTimers.set(id, setTimeout(() => pushProgressToFirestore(id), 450));
    }}

    async function pushProgressToFirestore(id) {{
      try {{
        await setDoc(doc(progressCollection, id), progressPayload(id), {{ merge: true }});
        setCloudStatus("Saved");
      }} catch (error) {{
        console.error("Firestore save failed", error);
        setCloudStatus("Blocked");
      }}
    }}

    async function loadCloudProgress() {{
      setCloudStatus("Loading");
      try {{
        const snapshot = await getDocs(progressCollection);
        snapshot.forEach(item => {{
          const data = item.data();
          state.progress[item.id] = {{
            status: data.status || "unreviewed",
            note: data.note || "",
            deleted: data.deleted === true,
            deletedAt: data.deletedAt || "",
            updatedAt: data.updatedAtIso || "",
          }};
          if (data.deleted === true) state.deletedIds.add(item.id);
        }});
        if (state.deletedIds.size >= auditData.length) {{
          console.warn("Cloud deleted flags hid every audit point; resetting local deleted state.");
          state.deletedIds.clear();
          for (const progress of Object.values(state.progress)) {{
            if (progress.deleted === true) progress.deleted = false;
          }}
          localStorage.removeItem(deletedStoreKey);
        }} else {{
          saveDeletedIds();
        }}
        saveLocalProgress();
        setCloudStatus(`${{snapshot.size}} loaded`);
        render();
      }} catch (error) {{
        console.error("Firestore load failed", error);
        setCloudStatus("Blocked");
      }}
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

    function modeText(row, technicalField, businessField) {{
      if (state.businessMode && row[businessField]) return row[businessField];
      return row[technicalField] || "";
    }}

    function issueText(row) {{
      return modeText(row, "issue", "businessIssue");
    }}

    function detailText(row) {{
      return modeText(row, "detail", "businessDetail") || row.summary || "";
    }}

    function fixText(row) {{
      return modeText(row, "fix", "businessFix") || "No fix steps were listed in the workbook.";
    }}

    function verifyText(row) {{
      return modeText(row, "verify", "businessVerify") || "No verification steps were listed in the workbook.";
    }}

    function isStackedLayout() {{
      return window.matchMedia("(max-width: 980px)").matches;
    }}

    function jumpToDetailOnMobile() {{
      if (!isStackedLayout()) return;
      requestAnimationFrame(() => {{
        document.querySelector(".right")?.scrollIntoView({{ behavior: "smooth", block: "start" }});
      }});
    }}

    function filteredRows() {{
      const q = state.query.trim().toLowerCase();
      return activeRows().filter(row => {{
        const progress = progressFor(row.id);
        if (state.tool !== "All" && row.tool !== state.tool) return false;
        if (state.sourceTab !== "All" && row.sourceTab !== state.sourceTab) return false;
        if (state.status !== "all" && progress.status !== state.status) return false;
        if (state.review === "manual" && !row.manualReview) return false;
        if (state.review === "nonmanual" && row.manualReview) return false;
        if (!q) return true;
        return [
          row.issue,
          row.businessIssue,
          row.issueId,
          row.tool,
          row.sourceTab,
          row.status,
          row.reviewStatus,
          row.summary,
          row.fix,
          row.businessFix,
          row.verify,
          row.businessVerify,
          row.businessDetail,
          row.docs,
        ].join("\\n").toLowerCase().includes(q);
      }});
    }}

    function renderToolbar() {{
      const activeToolCounts = activeRows().reduce((counts, row) => {{
        counts[row.tool] = (counts[row.tool] || 0) + 1;
        return counts;
      }}, {{}});
      const tools = ["All", ...Object.keys(toolCounts)];
      els.toolbar.innerHTML = "";
      for (const tool of tools) {{
        const button = document.createElement("button");
        button.type = "button";
        button.className = `tool-btn${{state.tool === tool ? " active" : ""}}`;
        const count = tool === "All" ? activeRows().length : (activeToolCounts[tool] || 0);
        button.textContent = tool === "All" ? `All (${{count.toLocaleString()}})` : `${{tool}} (${{count.toLocaleString()}})`;
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
      const rows = activeRows();
      for (const row of rows) {{
        const status = progressFor(row.id).status;
        if (status === "fixed") fixed += 1;
        if (status === "notfixed") notFixed += 1;
        if (row.manualReview) manual += 1;
      }}
      els.totalCount.textContent = rows.length.toLocaleString();
      els.visibleCount.textContent = visibleTotal.toLocaleString();
      els.manualCount.textContent = manual.toLocaleString();
      els.fixedCount.textContent = fixed.toLocaleString();
      els.notFixedCount.textContent = notFixed.toLocaleString();
      const completion = rows.length ? (fixed / rows.length) * 100 : 0;
      const completionText = `${{completion.toFixed(1).replace(".0", "")}}%`;
      els.completionPercent.textContent = completionText;
      els.completionFill.style.width = `${{Math.min(100, completion).toFixed(2)}}%`;
      els.completionFill.parentElement.setAttribute("aria-valuenow", completion.toFixed(1));
      els.completionMeta.textContent = `${{fixed.toLocaleString()}} of ${{rows.length.toLocaleString()}} tasks fixed`;
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
          <div class="point-title">${{escapeHtml(issueText(row))}}</div>
          <div class="meta">
            <span class="chip tool">${{escapeHtml(row.tool)}}</span>
            ${{row.issueId ? `<span class="chip">${{escapeHtml(row.issueId)}}</span>` : ""}}
            <span class="chip">${{escapeHtml(row.sourceTab)}}</span>
            ${{row.manualReview ? '<span class="chip manual">Manual review</span>' : ""}}
            <span class="chip ${{progress.status}}">${{statusLabel(progress.status)}}</span>
            <span class="chip">${{escapeHtml(row.status)}}</span>
          </div>`;
        button.addEventListener("click", () => {{
          state.selectedId = row.id;
          render();
          jumpToDetailOnMobile();
        }});
        fragment.appendChild(button);
      }}
      els.list.appendChild(fragment);
    }}

    function selectedRow(rows) {{
      let row = auditData.find(item => item.id === state.selectedId);
      if (!row || state.deletedIds.has(row.id) || (rows.length && !rows.some(item => item.id === row.id))) {{
        row = rows[0] || activeRows()[0];
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
      const businessBody = `
        <div class="detail-body">
          <div class="section">
            <h2>Plain Business Explanation</h2>
            <div class="content">${{escapeHtml(detailText(row))}}</div>
          </div>
        </div>`;
      const technicalBody = `
        <div class="detail-body">
          <div class="section">
            <h2>How To Fix</h2>
            <div class="content">${{escapeHtml(fixText(row))}}</div>
          </div>
          <div class="section">
            <h2>How To Verify</h2>
            <div class="content">${{escapeHtml(verifyText(row))}}</div>
          </div>
          <div class="section">
            <h2>Why It Matters</h2>
            <div class="content">${{escapeHtml(detailText(row))}}</div>
          </div>
          <div class="section">
            <h2>Official Docs</h2>
            <div class="content docs">${{docsHtml(row.docs)}}</div>
          </div>
          <div class="section">
            <h2>Routing Details</h2>
            <div class="content">${{escapeHtml([
              `Source tab: ${{row.sourceTab || "Not listed"}}`,
              `Original row: ${{row.row || "Not listed"}}`,
              row.reviewStatus ? `Review status: ${{row.reviewStatus}}` : "",
              row.routingNote ? `Routing note: ${{row.routingNote}}` : "",
              row.owner ? `Owner: ${{row.owner}}` : "",
              row.qaOutcome ? `QA outcome: ${{row.qaOutcome}}` : "",
              row.evidenceLink ? `Evidence link: ${{row.evidenceLink}}` : "",
            ].filter(Boolean).join("\\n"))}}</div>
          </div>
          <div class="section">
            <h2>My Notes</h2>
            <div class="content">
              <textarea class="note-box" id="noteBox" placeholder="Add evidence links, screenshot names, or what changed...">${{escapeHtml(progress.note || "")}}</textarea>
              <div class="small">Saved automatically to Firestore, with browser backup.</div>
            </div>
          </div>
        </div>`;
      els.detail.innerHTML = `
        <div class="detail-top">
          <h2 class="detail-title">${{escapeHtml(issueText(row))}}</h2>
          <div class="meta">
            <span class="chip tool">${{escapeHtml(row.tool)}}</span>
            ${{row.issueId ? `<span class="chip">${{escapeHtml(row.issueId)}}</span>` : ""}}
            <span class="chip">${{escapeHtml(row.sourceTab)}}</span>
            ${{row.manualReview ? '<span class="chip manual">Manual review required</span>' : ""}}
            <span class="chip">${{escapeHtml(row.status)}}</span>
            <span class="chip ${{progress.status}}">${{statusLabel(progress.status)}}</span>
          </div>
          <div class="actions">
            <button class="action-btn fixed" type="button" data-action="fixed">Mark fixed</button>
            <button class="action-btn notfixed" type="button" data-action="notfixed">Mark not fixed</button>
            <button class="action-btn ghost" type="button" data-action="unreviewed">Reset status</button>
            <button class="action-btn danger" type="button" data-delete-point>Delete point</button>
          </div>
        </div>
        ${{state.businessMode ? businessBody : technicalBody}}`;

      els.detail.querySelectorAll("[data-action]").forEach(button => {{
        button.addEventListener("click", () => setProgress(row.id, button.dataset.action));
      }});
      const deleteButton = els.detail.querySelector("[data-delete-point]");
      if (deleteButton) deleteButton.addEventListener("click", () => deletePoint(row.id));
      const noteBox = document.getElementById("noteBox");
      if (noteBox) noteBox.addEventListener("input", event => setNote(row.id, event.target.value));
    }}

    function render() {{
      renderToolbar();
      const rows = filteredRows();
      renderCounts(rows.length);
      els.resultLabel.textContent = `${{rows.length.toLocaleString()}} audit point${{rows.length === 1 ? "" : "s"}} shown`;
      els.technicalMode.classList.toggle("active", !state.businessMode);
      els.businessMode.classList.toggle("active", state.businessMode);
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
    els.technicalMode.addEventListener("click", () => {{
      state.businessMode = false;
      render();
    }});
    els.businessMode.addEventListener("click", () => {{
      state.businessMode = true;
      render();
    }});
    els.sourceFilter.addEventListener("change", event => {{
      state.sourceTab = event.target.value;
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
      state.sourceTab = "All";
      state.query = "";
      els.search.value = "";
      els.statusFilter.value = "all";
      els.reviewFilter.value = "all";
      els.sourceFilter.value = "All";
      render();
    }});

    populateSourceFilter();
    render();
    loadCloudProgress();
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
