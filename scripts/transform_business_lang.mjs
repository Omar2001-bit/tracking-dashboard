/**
 * transform_business_lang.mjs
 * Instantly rewrites all business-language fields using smart extraction
 * from the existing technical fields + a jargon replacement dictionary.
 * No API calls. Runs in < 1 second.
 *
 * Usage: node scripts/transform_business_lang.mjs
 */

import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dir = path.dirname(fileURLToPath(import.meta.url));
const HTML  = path.join(__dir, "..", "index.html");

// ── Jargon dictionary ─────────────────────────────────────────────────────────
// Applied to every business field after rewriting.
const JARGON = [
  // Tools
  [/\bGA4\b/g,                       "Google Analytics"],
  [/\bGTM\b/g,                       "Google Tag Manager"],
  [/\bGSC\b/g,                       "Google Search Console"],
  [/\bClarity\b/g,                   "Microsoft Clarity (visitor recording tool)"],
  [/\bShopify Pixels?\b/gi,          "Shopify tracking"],

  // Technical tracking terms
  [/\bdata\s?layer\b/gi,             "behind-the-scenes tracking data"],
  [/\bgtag\s*\([^)]*\)/gi,           "tracking code"],
  [/\bgtag\b/gi,                     "tracking code"],
  [/\bkey\s*events\b/gi,             "conversions"],
  [/\bkey\s*event\b/gi,              "conversion"],
  [/\bconversion event\b/gi,         "conversion"],
  [/\bevent\s*parameter\b/gi,        "data detail"],
  [/\bcustom\s*dimension\b/gi,       "custom data field"],
  [/\bdimension\b/gi,                "data field"],
  [/\bmetric\b/gi,                   "measurement"],
  [/\bcardinality\b/gi,              "too many unique values"],
  [/\bschema\b/gi,                   "data structure"],
  [/\bpayload\b/gi,                  "tracking data"],
  [/\bcontainer\b/gi,                "tag collection"],
  [/\btag\b/gi,                      "tracking code"],
  [/\btrigger\b/gi,                  "activation rule"],
  [/\bvariable\b/gi,                 "data value"],
  [/\bfiring\b/gi,                   "running"],

  // Sessions / traffic
  [/\bsession_start\b/gi,            "start of a visit"],
  [/\bpage_view\b/gi,                "page visit"],
  [/\bpageview\b/gi,                 "page visit"],
  [/\bsession\b/gi,                  "visit"],
  [/\bimpression\b/gi,               "ad view"],
  [/\bhostname\b/gi,                 "website domain"],
  [/\breferral\b/gi,                 "referral source"],
  [/\bbot\s*traffic\b/gi,            "fake/automated traffic"],
  [/\bbot\b/gi,                      "automated bot"],
  [/pangleglobal\.com/gi,            "a bot traffic source"],
  [/pangle[\w.-]*/gi,                "a bot traffic source"],
  [/\bghost\s*session/gi,            "fake visits"],

  // Ecommerce fields
  [/\bitemCategory\d*\b/gi,          "product category"],
  [/\bitemListName\b/gi,             "product list name"],
  [/\bitemListId\b/gi,               "product list ID"],
  [/\bitem_category\d*\b/gi,         "product category"],
  [/\bitem_list_name\b/gi,           "product list name"],
  [/\bitem_list_id\b/gi,             "product list ID"],
  [/\bitem_id\b/gi,                  "product ID"],
  [/\bitem_name\b/gi,                "product name"],
  [/\bitem_brand\b/gi,               "product brand"],
  [/\bitem_variant\b/gi,             "product variant"],
  [/\becommerce\b/gi,                "online store"],
  [/\be-commerce\b/gi,               "online store"],
  [/\btransaction\b/gi,              "order"],
  [/\bcheckout\b/gi,                 "checkout"],

  // Ad / attribution terms
  [/\bROAS\b/g,                      "return on ad spend"],
  [/\bCPA\b/g,                       "cost per customer"],
  [/\bCTR\b/g,                       "click rate"],
  [/\bCPC\b/g,                       "cost per click"],
  [/\bbidding\b/gi,                  "ad budget allocation"],
  [/\bsmart\s*bidding\b/gi,          "automated ad bidding"],
  [/\bconversion\s*window\b/gi,      "conversion tracking window"],
  [/\battribution\b/gi,              "credit for sales"],
  [/AW-[\w/]+/g,                     "your Google Ads conversion tag"],

  // Privacy / consent
  [/\bconsent\s*mode\b/gi,           "privacy consent settings"],
  [/\bconsent\s*signal\b/gi,         "privacy setting"],
  [/\bad_storage\b/gi,               "ad tracking"],
  [/\banalytics_storage\b/gi,        "analytics tracking"],
  [/\bad_user_data\b/gi,             "user data sharing"],
  [/\bad_personalization\b/gi,       "ad personalization"],
  [/\bGDPR\b/g,                      "GDPR privacy law"],
  [/\bGDPR\b/g,                      "GDPR privacy law"],
  [/\bcookie\s*consent\b/gi,         "cookie permission"],
  [/\bconsent\s*banner\b/gi,         "cookie banner"],

  // Technical UI paths
  [/GA4\s+Admin\s+UI/gi,             "Google Analytics settings"],
  [/Admin\s*[→>]\s*Key\s*Events/gi,  "the Conversions section"],
  [/Admin\s*[→>]/gi,                 "settings →"],
  [/GTM\s+Preview/gi,                "Google Tag Manager preview mode"],
  [/GTM\s+container/gi,              "Google Tag Manager"],
  [/Shopify\s+Admin\s*[→>]\s*/gi,    "Shopify settings → "],

  // Data quality terms
  [/\(not\s*set\)/gi,                "missing/blank"],
  [/\(other\)/gi,                    "unknown source"],
  [/100%\s+blank/gi,                 "completely empty"],
  [/pagePathPlusQueryString/gi,      "URL tracking field"],
  [/\bURL\s*parameter\b/gi,          "URL detail"],
  [/\bquery\s*string\b/gi,           "URL tracking detail"],
  [/\bserver-side\b/gi,              "server-based"],
  [/\bclient-side\b/gi,              "browser-based"],
  [/\bnative\b/gi,                   "built-in"],
  [/\bimplementation\b/gi,           "setup"],
  [/\bconfiguration\b/gi,            "settings"],
  [/\bintegration\b/gi,              "connection"],
  [/\bdeduplication\b/gi,            "duplicate prevention"],
  [/\bduplicate\b/gi,                "duplicate"],

  // Admin UI paths
  [/Google Analytics settings\s*[→>]\s*Admin\s*[→>]\s*/gi,  "Google Analytics → "],
  [/in the conversions section/gi,   "in the Conversions settings"],
  [/toggle is on/gi,                 "toggle is turned on"],
  [/toggle is off/gi,                "toggle is turned off"],
  [/\bimportant business action\b/gi, "conversion"],
  [/\bimportant business result\b/gi, "conversion"],
  [/\btracked action\b/gi,           "action"],
  [/\btracking workspace\b/gi,       "Google Tag Manager account"],
  [/\bworkspace\b/gi,                "account"],
  [/\bsafeguard\b/gi,                "protection"],
  [/\bgranular\b/gi,                 "detailed"],
  [/\bpersist\b/gi,                  "continue"],
  [/\bproperty\b(?!\s+type)/gi,      "account"],
  [/\bstandard\s*ecommerce\b/gi,     "standard online store tracking"],

  // Generic cleanup
  [/\bverify\b/gi,                   "confirm"],
  [/\binspect\b/gi,                  "check"],
  [/\bfired?\b/gi,                   "ran"],
  [/\bfiring\b/gi,                   "running"],
  [/\bdeployed?\b/gi,                "set up"],
  [/\binstantiated?\b/gi,            "started"],
  [/\bpopulated?\b/gi,               "filled in"],
  [/\bpropagated?\b/gi,              "passed on"],
  [/\bbool\b/gi,                     "yes/no"],
  [/\bnull\b/gi,                     "empty"],
  [/\bundefined\b/gi,                "missing"],
  [/\b(regex|regexp)\b/gi,           "pattern"],
  [/`([^`]+)`/g,                     '"$1"'],  // remove backtick code formatting
];

function cleanJargon(text) {
  if (!text) return "";
  let t = text;
  for (const [pattern, replacement] of JARGON) {
    t = t.replace(pattern, replacement);
  }
  // Strip leftover backticks
  t = t.replace(/`/g, '"');
  // Collapse multiple spaces
  t = t.replace(/  +/g, " ").trim();
  return t;
}

// ── Extract helpers ───────────────────────────────────────────────────────────

/** Pull section after a heading like "Why it matters:\n..." */
function extractSection(text, ...headings) {
  for (const h of headings) {
    const re = new RegExp(h + "[:\\s]*\\n?([\\s\\S]*?)(?:\\n\\n|$)", "i");
    const m  = text.match(re);
    if (m && m[1].trim().length > 20) return m[1].trim();
  }
  return null;
}

/** First N sentences of a string. */
function firstSentences(text, n = 2) {
  const raw = text.replace(/\n+/g, " ").trim();
  // Split on . ! ? followed by space+capital or end
  const parts = raw.match(/[^.!?]+[.!?]+(?:\s|$)/g) || [raw];
  return parts.slice(0, n).join(" ").trim();
}

/** Strip leading labels like "What to do in business terms:" */
function stripLeadingLabel(text) {
  return text
    .replace(/^(What to do in business terms|How to know it is fixed|Explanation|Why it matters)\s*:\s*/i, "")
    .replace(/^Use the original fix,?\s*(but)?\s*/i, "")
    .trim();
}

/** Convert numbered/bulleted technical steps to plain sentences. */
function simplifySteps(text, maxSteps = 4) {
  const stripped = stripLeadingLabel(text);
  // Split on numbered list or newlines
  const lines = stripped.split(/\n|\d+\.\s+/).map(l => l.trim()).filter(Boolean);
  const kept   = lines.slice(0, maxSteps).map(l => {
    // Clean jargon
    let s = cleanJargon(l);
    // Ensure it ends with a period
    if (!/[.!?]$/.test(s)) s += ".";
    return s;
  });
  return kept.join("\n");
}

// ── Per-row transform ─────────────────────────────────────────────────────────

const STATUS_PASS = /\b(correct|pass|no action|already set|confirmed present|no issues? found)\b/i;

function transformRow(row) {
  const isPass = STATUS_PASS.test(row.status || "") ||
                 STATUS_PASS.test(row.issue  || "") ||
                 STATUS_PASS.test((row.detail || "").slice(0, 200));

  if (isPass) {
    // Pass rows: short positive statement, no fix/verify needed
    const topic = cleanJargon(
      (row.summary || row.issue || "")
        .replace(/^\[[^\]]+\]\s*/,"")   // strip [TOOL | ID]
        .split(/\n/)[0]
        .slice(0, 120)
    );
    return {
      businessIssue:  `This is already set up correctly — ${firstSentences(topic, 1)} No action needed.`,
      businessDetail: `Your team has already handled this correctly. ${firstSentences(cleanJargon(row.detail || row.summary || ""), 2)}`,
      businessFix:    "",
      businessVerify: "",
    };
  }

  // ── businessIssue ──────────────────────────────────────────────────────────
  // Use the first sentence of summary (after "Explanation:"), cleaned
  const summaryClean = cleanJargon(
    (row.summary || row.issue || "")
      .replace(/^Explanation\s*:\s*/i, "")
      .replace(/^\[[^\]]+\]\s*/,"")
      .split(/\n/)[0]
  );
  const businessIssue = firstSentences(summaryClean, 2);

  // ── businessDetail ─────────────────────────────────────────────────────────
  // Extract "Why it matters" from the technical detail field
  const whySection =
    extractSection(row.detail || "", "Why it matters", "Why this matters", "Why It Matters") ||
    extractSection(row.detail || "", "Explanation") ||
    (row.detail || "").split(/\n\n/)[0] ||
    row.summary || "";

  const businessDetail = cleanJargon(
    firstSentences(
      whySection
        .replace(/^Explanation\s*:\s*/i, "")
        .replace(/Recommended action[\s\S]*/i, "")
        .replace(/Verification standard[\s\S]*/i, ""),
      5
    )
  );

  // ── businessFix ────────────────────────────────────────────────────────────
  const rawFix = row.fix || "";
  const isBoilerplateFix = /use the original fix/i.test(rawFix) || rawFix.trim().length < 20;
  const fixSource = isBoilerplateFix
    ? (row.detail || "").match(/Recommended action\s*:\s*([\s\S]*?)(?:\n\n|Verification standard|$)/i)?.[1] || rawFix
    : rawFix;

  let businessFix = simplifySteps(fixSource, 4);
  // Capitalise first letter of each step line
  businessFix = businessFix.split("\n").map(line => {
    const t = line.trimStart();
    return t.charAt(0).toUpperCase() + t.slice(1);
  }).join("\n");
  // Remove any remaining "key-event status" style phrases
  businessFix = cleanJargon(businessFix)
    .replace(/key-event status is OFF/gi,  "conversion toggle is turned off")
    .replace(/key-event status\b/gi,       "conversion setting")
    .replace(/unless the business has intentionally defined it as a true conversion/gi,
             "unless it is a real business action like a purchase or sign-up");

  // ── businessVerify ─────────────────────────────────────────────────────────
  const rawVerify   = row.verify || "";
  // Take the "Done means..." sentence as the verify outcome
  const doneMatch   = rawVerify.match(/Done means[^.]*\.[^.]*\./i);
  const verifyBase  = doneMatch
    ? doneMatch[0]
    : firstSentences(rawVerify.replace(/^\d+\.\s+/, ""), 2);

  // Turn "Done means the evidence now supports this target: X" → "After fixing this, X"
  const verifyClean = cleanJargon(
    stripLeadingLabel(verifyBase)
      .replace(/Done means the evidence now supports this target\s*:\s*/gi, "After fixing this, ")
      .replace(/Done means[^:]+:\s*/gi,                                    "After fixing this, ")
      .replace(/Done means\s*/gi,                                          "After fixing this, ")
      .replace(/Add a screenshot[^.]*\./gi,                                "")
      .replace(/Confirm the final state matches this expected condition\s*:\s*/gi, "The expected outcome: ")
      .replace(/Add a [^.]+\./gi,                                          "")
      .trim()
  );
  // Capitalise first letter
  const businessVerify = verifyClean.charAt(0).toUpperCase() + verifyClean.slice(1);

  return { businessIssue, businessDetail, businessFix, businessVerify };
}

// ── Apply to HTML ─────────────────────────────────────────────────────────────

const html  = fs.readFileSync(HTML, "utf8");
const tag   = 'id="audit-data">';
const start = html.indexOf(tag) + tag.length;
const end   = html.indexOf("</script>", start);
const rows  = JSON.parse(html.slice(start, end));

console.log(`Transforming ${rows.length} rows…`);

const improved = rows.map(row => {
  const t = transformRow(row);
  return { ...row, ...t };
});

const newHtml = html.slice(0, start) + JSON.stringify(improved) + html.slice(end);
fs.writeFileSync(HTML, newHtml, "utf8");

console.log(`Done — ${rows.length} rows updated in index.html`);

// Quick quality check on first 3 rows
console.log("\n── Sample output ──────────────────────────────────────────────");
improved.slice(0, 3).forEach((r, i) => {
  console.log(`\nRow ${i+1}: ${r.id}`);
  console.log("businessIssue: ", r.businessIssue);
  console.log("businessDetail:", r.businessDetail);
  console.log("businessFix:   ", r.businessFix?.slice(0,150));
  console.log("businessVerify:", r.businessVerify?.slice(0,150));
});
