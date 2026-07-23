#!/usr/bin/env node
/**
 * Full pre-release / pre-tag verification gauntlet. Runs everything CI gates
 * (`guard:ci`) PLUS the maintainer-only checks CI physically can't — the ones
 * that need Playwright and/or LibreOffice — plus the zero-tolerance scored
 * suite. This is a LOCAL/maintainer gate, not a CI step: it needs Chromium
 * (`npm run setup`) and LibreOffice on PATH.
 *
 * Not in CI on purpose — if these could run headless in CI they'd already be in
 * guard:ci. Run from repo root before a release commit or tag: npm run verify:release
 */
import { spawnSync } from "node:child_process";

const npm = process.platform === "win32" ? "npm.cmd" : "npm";

/** Prerequisites — everything downstream needs a fresh dist/, so stop if these fail. */
const PREREQS = [
  ["typecheck", "tsc --noEmit — no type errors"],
  ["build:all", "build library + browser bundle from clean"],
];

/** Verification checks — run all even if one fails, so every problem surfaces at once. */
const CHECKS = [
  ["guard:ci", "9 CI-safe guards (inline, config, fields, mixed-orientation, toc-slot, internal-href, document-canvas, image-spacing, pack-smoke)"],
  ["guard:computed-parity", "oracle vs native byte-identical OOXML (Playwright)"],
  ["guard:browser-parity", "shipped browser bundle vs Node native (Playwright)"],
  ["guard:page-break", "structural page breaks: OOXML + multi-page PDF (LibreOffice)"],
  ["score:suite:strict", "full scored suite, zero-tolerance pixel regression (Chromium + LibreOffice)"],
];

/** Deliberately NOT run here — situational / not release gates. Printed for the maintainer. */
const NOT_DONE = [
  ["npm run docs:sync", "regenerate docs/TEST-SCORES.md + BENCHMARK.md from this run, then commit them"],
  ["npm run score:benchmark", "re-score html-to-docx / turbodocx for the comparison tables (only if the pitch changed)"],
  ["npm run score:style-source / score:css-cascade", "inline-vs-computed benchmarks (only if resolution changed)"],
  ["npm run score:calibration", "pipeline-noise floor (only when tuning the metric)"],
  ["npm run research:*", "word-spotcheck / novel / wild / concordance — validate the scoring metric, not the converter"],
];

function banner() {
  console.log("═══ verify:release — full pre-release gauntlet ═══");
  console.log("Needs Chromium (`npm run setup`) + LibreOffice on PATH.\n");
  console.log("Prerequisites (stop on first failure):");
  for (const [name, desc] of PREREQS) console.log(`  • ${name.padEnd(22)} ${desc}`);
  console.log("\nChecks (all run even if one fails):");
  for (const [name, desc] of CHECKS) console.log(`  • ${name.padEnd(22)} ${desc}`);
  console.log("\nDoes NOT run (situational — do these yourself if relevant):");
  for (const [cmd, desc] of NOT_DONE) console.log(`  • ${cmd}\n      ${desc}`);
  console.log("");
}

function run(name) {
  console.log(`\n─── ${name} ───`);
  return spawnSync(npm, ["run", name], { stdio: "inherit" }).status === 0;
}

banner();

for (const [name] of PREREQS) {
  if (!run(name)) {
    console.log(`\n═══ verify:release ═══\n✗ prerequisite failed: ${name} — stopping before the checks.`);
    process.exit(1);
  }
}

const failed = [];
for (const [name] of CHECKS) {
  if (!run(name)) failed.push(name);
}

console.log(`\n═══ verify:release ═══`);
if (failed.length > 0) {
  console.log(`✗ ${failed.length}/${CHECKS.length} checks failed: ${failed.join(", ")}`);
  process.exit(1);
}
console.log(`✓ all ${CHECKS.length} checks passed. Before tagging, consider the "Does NOT run" list above.`);
