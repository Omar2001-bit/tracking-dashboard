/**
 * rewrite_business_lang.mjs
 * Rewrites businessIssue / businessDetail / businessFix / businessVerify
 * for every audit row using the local Claude CLI (authenticated session).
 *
 * Usage:  node scripts/rewrite_business_lang.mjs
 *
 * Progress is saved after every batch so the script is safely resumable.
 * Reads  : scripts/_audit_raw.json
 * Writes : scripts/_audit_improved.json   ← inspect before applying
 * Apply  : node scripts/apply_improved.mjs
 */

import { execSync } from "child_process";
import fs   from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dir   = path.dirname(fileURLToPath(import.meta.url));
const RAW     = path.join(__dir, "_audit_raw.json");
const OUT     = path.join(__dir, "_audit_improved.json");
const PROG    = path.join(__dir, "_progress.json");
const LOG     = path.join(__dir, "_rewrite.log");

const BATCH   = 20;   // rows per claude call (keep ≤25 for reliable JSON output)
const RETRIES = 3;

// ── Helpers ──────────────────────────────────────────────────────────────────

function log(msg) {
  const line = `[${new Date().toISOString()}] ${msg}`;
  console.log(line);
  fs.appendFileSync(LOG, line + "\n");
}

function buildPrompt(rows) {
  const slim = rows.map(r => ({
    id:      r.id,
    status:  r.status,
    issue:   (r.issue   || "").slice(0, 300),
    summary: (r.summary || "").slice(0, 400),
    detail:  (r.detail  || "").slice(0, 800),
    fix:     (r.fix     || "").slice(0, 400),
    verify:  (r.verify  || "").slice(0, 300),
  }));

  return [
    "Rewrite businessIssue, businessDetail, businessFix, businessVerify for each row.",
    "The client is a non-technical business owner of Tharaa.shop (Shopify, Middle East e-commerce).",
    "",
    "Rules:",
    "- businessIssue : 1-2 plain-English sentences. What is wrong or what was found.",
    "- businessDetail: 3-5 sentences. Concrete business impact: lost revenue, wasted ad budget, wrong data, legal/privacy risk.",
    "- businessFix   : 2-5 steps starting with a verb. What the owner or their developer must do. Keep it simple.",
    "- businessVerify: 1-2 sentences. What the owner will see change after the fix. Plain outcome, not a technical test.",
    "- ZERO jargon: no GA4, GTM, dataLayer, schema, tag, trigger, variable, cardinality, ROAS, CPA, gtag, pagePathPlusQueryString, item_*, (not set), hostname, pangleglobal, AW-xxxxxxxxx.",
    "- Replace with: Google Analytics, Google Tag Manager, tracking code, data fields, page visit, conversion, ad budget, product category, product list, missing data, junk/bot traffic.",
    "- If a row status is a PASS or the issue is already correct/no-action-needed, set businessIssue to a single sentence saying it is set up correctly and leave businessFix and businessVerify as empty strings \"\".",
    "",
    "Return ONLY a valid JSON array — no markdown fences, no explanation:",
    '[{"id":"...","businessIssue":"...","businessDetail":"...","businessFix":"...","businessVerify":"..."}]',
    "",
    "Input rows:",
    JSON.stringify(slim, null, 2),
  ].join("\n");
}

function callClaude(prompt) {
  // Write prompt to temp file to avoid shell length limits
  const tmp = path.join(__dir, "_tmp_prompt.txt");
  fs.writeFileSync(tmp, prompt, "utf8");

  const out = execSync(
    `claude --print --no-session-persistence --model haiku --dangerously-skip-permissions --allowedTools "" < "${tmp}"`,
    {
      encoding:   "utf8",
      maxBuffer:  20 * 1024 * 1024,
      timeout:    4 * 60 * 1000,   // 4 min per batch
      shell:      true,
    }
  );

  try { fs.unlinkSync(tmp); } catch {}
  return out.trim();
}

function parseJSON(text, expectedLen) {
  // Strip markdown fences if present
  let clean = text.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "").trim();
  const s = clean.indexOf("[");
  const e = clean.lastIndexOf("]");
  if (s === -1 || e === -1) throw new Error("No JSON array in response");
  clean = clean.slice(s, e + 1);
  const arr = JSON.parse(clean);
  if (!Array.isArray(arr)) throw new Error("Parsed value is not an array");
  if (arr.length !== expectedLen)
    throw new Error(`Expected ${expectedLen} items, got ${arr.length}`);
  return arr;
}

// ── Main ─────────────────────────────────────────────────────────────────────

const allRows = JSON.parse(fs.readFileSync(RAW, "utf8"));
log(`Loaded ${allRows.length} rows.`);

// Load prior progress
let done = {};
if (fs.existsSync(PROG)) {
  done = JSON.parse(fs.readFileSync(PROG, "utf8"));
  log(`Resuming — ${Object.keys(done).length} rows already improved.`);
}

const pending = allRows.filter(r => !done[r.id]);
const numBatches = Math.ceil(pending.length / BATCH);
log(`${pending.length} pending rows → ${numBatches} batches of ${BATCH}.`);

for (let b = 0; b < numBatches; b++) {
  const batch  = pending.slice(b * BATCH, (b + 1) * BATCH);
  const label  = `Batch ${b + 1}/${numBatches}`;
  const prompt = buildPrompt(batch);

  let ok = false;
  for (let attempt = 1; attempt <= RETRIES; attempt++) {
    try {
      log(`${label} — attempt ${attempt}…`);
      const raw    = callClaude(prompt);
      const result = parseJSON(raw, batch.length);
      result.forEach(item => { done[item.id] = item; });
      fs.writeFileSync(PROG, JSON.stringify(done, null, 2));
      log(`${label} ✓ (${Object.keys(done).length}/${allRows.length} total done)`);
      ok = true;
      break;
    } catch (err) {
      log(`${label} attempt ${attempt} failed: ${err.message}`);
      if (attempt === RETRIES) log(`${label} — giving up after ${RETRIES} attempts.`);
    }
  }
  if (!ok) fs.writeFileSync(PROG, JSON.stringify(done, null, 2)); // save partial
}

// Merge improved fields back onto original rows
const final = allRows.map(row => {
  const upd = done[row.id];
  if (!upd) return row; // left untouched if batch failed
  return {
    ...row,
    businessIssue:  upd.businessIssue  ?? row.businessIssue,
    businessDetail: upd.businessDetail ?? row.businessDetail,
    businessFix:    upd.businessFix    ?? row.businessFix,
    businessVerify: upd.businessVerify ?? row.businessVerify,
  };
});

fs.writeFileSync(OUT, JSON.stringify(final, null, 2));
log(`\nDone! Improved data → ${OUT}`);
log(`Review it, then run: node scripts/apply_improved.mjs`);
