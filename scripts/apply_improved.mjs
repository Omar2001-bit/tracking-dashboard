/**
 * apply_improved.mjs
 * Patches index.html with the improved business-language rows
 * produced by rewrite_business_lang.mjs.
 *
 * Usage:  node scripts/apply_improved.mjs
 */

import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dir  = path.dirname(fileURLToPath(import.meta.url));
const HTML   = path.join(__dir, "..", "index.html");
const IMPROVED = path.join(__dir, "_audit_improved.json");

if (!fs.existsSync(IMPROVED)) {
  console.error("Run rewrite_business_lang.mjs first — _audit_improved.json not found.");
  process.exit(1);
}

const improved = JSON.parse(fs.readFileSync(IMPROVED, "utf8"));
const html     = fs.readFileSync(HTML, "utf8");

const tag   = 'id="audit-data">';
const start = html.indexOf(tag) + tag.length;
const end   = html.indexOf("</script>", start);

if (start === -1 || end === -1) {
  console.error("Could not find audit-data script tag in index.html");
  process.exit(1);
}

const newHtml = html.slice(0, start) + JSON.stringify(improved) + html.slice(end);
fs.writeFileSync(HTML, newHtml, "utf8");
console.log(`✓ index.html patched with ${improved.length} improved rows.`);
