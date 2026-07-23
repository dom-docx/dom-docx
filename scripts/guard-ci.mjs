#!/usr/bin/env node
/**
 * Run the CI-safe guard suite — the binary pass/fail invariants that need no
 * Playwright or LibreOffice. This is the single entry point CI and publish use
 * (`npm run guard:ci`); the underlying `guard:*` scripts stay individually
 * runnable for local iteration and are still documented one-per-name in
 * CONTRIBUTING.md. Add a CI-safe guard in exactly one place: the list below.
 *
 * Unlike a fail-fast `&&` chain, this runs every guard even after one fails and
 * prints a summary, so a broken build surfaces all failures in a single run.
 * Exits non-zero if any guard fails. Run from repo root: npm run guard:ci
 */
import { spawnSync } from "node:child_process";

/** CI-safe guards (no Playwright/LibreOffice), run in this order. */
const GUARDS = [
  "guard:inline",
  "guard:config",
  "guard:fields",
  "guard:mixed-orientation",
  "guard:toc-slot",
  "guard:internal-href",
  "guard:document-canvas",
  "guard:image-spacing",
  "guard:pack-smoke",
];

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const failed = [];

for (const guard of GUARDS) {
  console.log(`\n─── ${guard} ───`);
  const { status } = spawnSync(npm, ["run", guard], { stdio: "inherit" });
  if (status !== 0) {
    failed.push(guard);
  }
}

console.log(`\n═══ guard:ci ═══`);
if (failed.length > 0) {
  console.log(`✗ ${failed.length}/${GUARDS.length} failed: ${failed.join(", ")}`);
  process.exit(1);
}
console.log(`✓ ${GUARDS.length}/${GUARDS.length} guards passed`);
