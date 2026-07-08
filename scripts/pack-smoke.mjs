#!/usr/bin/env node
/**
 * Verify the npm tarball installs and converts without Playwright.
 * Run from repo root: npm run guard:pack-smoke
 */
import { execSync } from "node:child_process";
import { mkdir, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/** Mirrors tools/guard-result.ts's shape; duplicated here since this script runs via
 * plain `node`, not `tsx`, and can't import a .ts module directly. */
async function writeGuardResult(ok, detail) {
  const guardsDir = path.join(root, "output", "guards");
  await new Promise((resolve, reject) =>
    mkdir(guardsDir, { recursive: true }, (err) => (err ? reject(err) : resolve())),
  );
  const payload = {
    id: "pack-smoke",
    label: "Pack smoke",
    passed: ok ? 1 : 0,
    total: 1,
    ok,
    unit: detail,
    ranAt: new Date().toISOString(),
    command: "npm run guard:pack-smoke",
  };
  await writeFile(path.join(guardsDir, "pack-smoke.json"), `${JSON.stringify(payload, null, 2)}\n`, "utf-8");
}
const packLines = execSync("npm pack --silent", { cwd: root, encoding: "utf-8" })
  .trim()
  .split("\n")
  .map((line) => line.trim())
  .filter(Boolean);
const packOut = packLines.find((line) => line.endsWith(".tgz"));
if (!packOut) {
  throw new Error(`npm pack did not print a .tgz filename (got: ${packLines.join(" | ")})`);
}
const tgz = path.join(root, packOut);

const work = mkdtempSync(path.join(tmpdir(), "dom-docx-smoke-"));
try {
  writeFileSync(path.join(work, "package.json"), JSON.stringify({ type: "module", private: true }, null, 2));
  execSync(`npm install "${tgz}"`, { cwd: work, stdio: "inherit" });

  const nm = path.join(work, "node_modules");
  const pkg = JSON.parse(readFileSync(path.join(nm, "dom-docx/package.json"), "utf-8"));
  if (pkg.dependencies?.playwright || pkg.optionalDependencies?.playwright) {
    throw new Error(
      "playwright must not be a hard/optional dependency (use an optional peerDependency so `npm i` stays lean)",
    );
  }
  if (pkg.peerDependencies?.playwright && !pkg.peerDependenciesMeta?.playwright?.optional) {
    throw new Error("playwright peerDependency must be marked optional in peerDependenciesMeta");
  }
  const hasPlaywrightDir = (() => {
    try {
      readFileSync(path.join(nm, "playwright/package.json"));
      return true;
    } catch {
      return false;
    }
  })();
  if (hasPlaywrightDir) {
    throw new Error(
      "playwright was installed into a clean consumer project — an optional peerDependency must NOT auto-install",
    );
  }

  const domDocxRoot = path.join(nm, "dom-docx");
  for (const rel of [
    "dist/browser.js",
    "dist/browser.d.ts",
    "dist/browser/dom-docx.browser.js",
  ]) {
    try {
      readFileSync(path.join(domDocxRoot, rel));
    } catch {
      throw new Error(`browser entry missing from tarball: ${rel}`);
    }
  }

  const browserExport = pkg.exports?.["./browser"];
  if (browserExport?.import !== "./dist/browser.js") {
    throw new Error('exports["./browser"].import must point at ./dist/browser.js');
  }
  if (browserExport?.default !== "./dist/browser.js") {
    throw new Error('exports["./browser"].default must point at ./dist/browser.js');
  }

  const smoke = `
    import { convertHtmlToDocx } from 'dom-docx';
    import { writeFileSync } from 'fs';
    const buf = await convertHtmlToDocx('<p>pack smoke</p>');
    if (buf[0] !== 0x50 || buf[1] !== 0x4b) throw new Error('not a zip/docx');
    writeFileSync('out.docx', buf);
    console.log('ok:', buf.length, 'bytes');
  `;
  writeFileSync(path.join(work, "smoke.mjs"), smoke);
  execSync("node smoke.mjs", { cwd: work, stdio: "inherit" });

  const out = readFileSync(path.join(work, "out.docx"));
  if (out.length < 100) throw new Error("docx too small");

  // CLI: the bin entry must ship, execute, and convert from the installed tarball.
  writeFileSync(path.join(work, "cli-in.html"), "<h1>CLI smoke</h1><p>via bin</p>");
  execSync("npx --no-install dom-docx cli-in.html -o cli-out.docx", { cwd: work, stdio: "inherit" });
  const cliOut = readFileSync(path.join(work, "cli-out.docx"));
  if (cliOut[0] !== 0x50 || cliOut[1] !== 0x4b) throw new Error("CLI output is not a zip/docx");

  const browserSmoke = `
    import { convertHtmlToDocx } from 'dom-docx/browser';
    const blob = await convertHtmlToDocx('<p>browser entry</p>', { styleSource: 'inline' });
    const buf = Buffer.from(await blob.arrayBuffer());
    if (buf[0] !== 0x50 || buf[1] !== 0x4b) throw new Error('browser entry: not a zip/docx');
    console.log('browser entry ok:', buf.length, 'bytes');
  `;
  writeFileSync(path.join(work, "browser-smoke.mjs"), browserSmoke);
  execSync("node browser-smoke.mjs", { cwd: work, stdio: "inherit" });

  console.log("pack-smoke: passed (library + CLI + browser entry)");
  await writeGuardResult(true, "library + CLI + browser entry install and convert");
} catch (err) {
  await writeGuardResult(false, err instanceof Error ? err.message : String(err));
  throw err;
} finally {
  rmSync(work, { recursive: true, force: true });
  rmSync(tgz, { force: true });
}
