#!/usr/bin/env node
// Ad-hoc: convert a real, live webpage to .docx for manual inspection.
// Not part of the test suite — just a way to point the converter at arbitrary
// URLs (like the Red Hat docs page from the bug report) and see what happens.
//
// Usage:
//   node scripts/try-url.mjs <url> [contentSelector] [outFile]
//
// Examples:
//   node scripts/try-url.mjs https://docs.redhat.com/.../managing_storage_devices/index
//   node scripts/try-url.mjs https://example.com "main" out.docx
import { writeFile } from "node:fs/promises";
import { chromium } from "playwright";
import { convertHtmlToDocx } from "../src/index.ts";

const [, , url, contentSelector = "body", outFile = "try-url-output.docx"] = process.argv;
if (!url) {
  console.error("Usage: node scripts/try-url.mjs <url> [contentSelector] [outFile]");
  process.exit(1);
}

const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  await page.goto(url, { waitUntil: "networkidle" });

  const root = await page.$(contentSelector);
  if (!root) throw new Error(`no element matched selector "${contentSelector}" on ${url}`);

  // Unsanitized fetch-based resolver — fine for local, one-off inspection.
  // Do NOT ship this as-is: no host allowlist, no size cap, no private-IP guard.
  const imageResolver = async (src) => {
    try {
      const absolute = new URL(src, url).href;
      const res = await fetch(absolute);
      if (!res.ok) return null;
      const buf = new Uint8Array(await res.arrayBuffer());
      const ct = res.headers.get("content-type") ?? "";
      const type = ct.includes("png") ? "png" : ct.includes("gif") ? "gif" : ct.includes("bmp") ? "bmp" : "jpg";
      return { data: buf, type };
    } catch {
      return null;
    }
  };

  const docx = await convertHtmlToDocx(await root.innerHTML(), {
    styleSource: "computed",
    page,
    rootSelector: contentSelector,
    imageResolver,
    // default onWarning (console.warn) — leave it on so degradation is visible
  });

  await writeFile(outFile, docx);
  console.error(`wrote ${outFile} (${docx.length} bytes)`);
} finally {
  await browser.close();
}
