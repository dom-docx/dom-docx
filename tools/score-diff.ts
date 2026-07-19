/**
 * Suite score baseline — zero-ritual regression check, printed after every `score:suite`.
 *
 * How it works (like snapshot tests):
 *   - First full run with no baseline → this run is saved as the baseline automatically.
 *   - Every later run → prints ONLY what changed vs the baseline (a green line if nothing
 *     moved), so you don't eyeball 40+ numbers to spot a regression.
 *   - Brand-new cases (no "before" to regress from) are adopted into the baseline
 *     automatically.
 *   - `npm run score:pin` → deliberately re-set the baseline to the current run (e.g. after
 *     accepting new scores, or to reset to `main` before starting a change).
 *
 * The baseline is `output/suite/baseline.json` (gitignored). Visual scores vary across
 * OS/font/LibreOffice, so it's a same-machine check, not a shared gate — the committed
 * `docs/TEST-SCORES.md` is the reference contributors read for the current champion scores.
 */
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { SUITE_OUTPUT } from "./output-paths.js";

const RESULTS = path.join(SUITE_OUTPUT, "results.json");
const BASELINE = path.join(SUITE_OUTPUT, "baseline.json");
/** Δ beyond this (percentage points) counts as a real change, not scoring noise. */
const THRESHOLD = 0.5;

interface CaseScore {
  name: string;
  visualScore?: number;
}
interface Results {
  cases: CaseScore[];
  totalCaseCount?: number;
}

async function load(file: string): Promise<Results | null> {
  try {
    return JSON.parse(await readFile(file, "utf-8")) as Results;
  } catch {
    return null;
  }
}

async function save(results: Results): Promise<void> {
  await writeFile(BASELINE, JSON.stringify(results, null, 2) + "\n", "utf-8");
}

function scoreMap(r: Results): Map<string, number> {
  const m = new Map<string, number>();
  for (const c of r.cases) if (typeof c.visualScore === "number") m.set(c.name, c.visualScore);
  return m;
}

function avg(r: Results): number {
  const scores = r.cases.map((c) => c.visualScore).filter((v): v is number => typeof v === "number");
  return scores.length ? scores.reduce((a, b) => a + b, 0) / scores.length : 0;
}

const isFullRun = (r: Results): boolean =>
  r.totalCaseCount === undefined || r.cases.length >= r.totalCaseCount;

export async function pinBaseline(): Promise<void> {
  const current = await load(RESULTS);
  if (!current) throw new Error(`No results.json at ${RESULTS} — run \`npm run score:suite\` first.`);
  await save(current);
  console.log(`Pinned baseline: ${current.cases.length} cases → ${path.relative(process.cwd(), BASELINE)}`);
}

/** Print current results.json diffed against the baseline (condensed). Never throws. */
export async function printBaselineDiff(): Promise<void> {
  const current = await load(RESULTS);
  if (!current) return;
  const baseline = await load(BASELINE);

  // First full run establishes the baseline — no ritual.
  if (!baseline) {
    if (isFullRun(current)) {
      await save(current);
      console.log(
        `\nvs baseline: none yet — saved this run as the baseline (${current.cases.length} cases, avg ${avg(current).toFixed(2)}). Re-run after changes to see the diff.`,
      );
    } else {
      console.log("\nvs baseline: none yet — run the full `npm run score:suite` once to set it.");
    }
    return;
  }

  const cur = scoreMap(current);
  const base = scoreMap(baseline);
  const changed: Array<{ name: string; before: number; after: number; delta: number }> = [];
  const newCases: string[] = [];
  for (const [name, after] of cur) {
    const before = base.get(name);
    if (before === undefined) newCases.push(name);
    else if (Math.abs(after - before) >= THRESHOLD) changed.push({ name, before, after, delta: after - before });
  }
  changed.sort((a, b) => a.delta - b.delta); // regressions first

  // A new case has no "before" and can't regress — adopt it into the baseline
  // automatically. `score:pin` stays reserved for accepting CHANGED scores.
  if (newCases.length) {
    const adopted = new Set(newCases);
    baseline.cases.push(...current.cases.filter((c) => adopted.has(c.name)));
    if (current.totalCaseCount !== undefined) baseline.totalCaseCount = current.totalCaseCount;
    await save(baseline);
  }

  const n = cur.size;
  const overall = avg(current);
  const header = `━━━ vs baseline · ${n} case${n === 1 ? "" : "s"} · avg ${overall.toFixed(2)} ━━━`;
  if (changed.length === 0 && newCases.length === 0) {
    console.log(`\n${header}\n✓ no changes`);
    return;
  }

  console.log(`\n${header}`);
  const w = Math.max(...changed.map((c) => c.name.length), 4);
  for (const c of changed) {
    const mark = c.delta < 0 ? "✗" : "✓";
    console.log(
      `  ${mark} ${c.name.padEnd(w)}  ${c.before.toFixed(2)} → ${c.after.toFixed(2)}  (${c.delta >= 0 ? "+" : ""}${c.delta.toFixed(2)})`,
    );
  }
  const regressed = changed.filter((c) => c.delta < 0).length;
  const improved = changed.length - regressed;
  console.log(
    `  ${regressed} regressed · ${improved} improved · ${n - changed.length - newCases.length} unchanged (±${THRESHOLD})`,
  );
  if (newCases.length) console.log(`  new: ${newCases.join(", ")} — added to baseline`);
}

// CLI: `--pin` re-sets the baseline (npm run score:pin). No args → print the diff.
if (process.argv[1] && import.meta.url === `file://${process.argv[1]}`) {
  const run = process.argv.includes("--pin") ? pinBaseline() : printBaselineDiff();
  run.catch((err) => {
    console.error(err instanceof Error ? err.message : err);
    process.exit(2);
  });
}
