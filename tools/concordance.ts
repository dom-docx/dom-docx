/**
 * Human-label concordance harness — validates visual metrics against the
 * hand-labeled ground truth sets (methodology from
 * internal/research/visual-scoring-metric-research-2026-07-02.md §8–9).
 *
 * Pairwise concordance: across all case pairs with DIFFERENT human ratings,
 * how often does the metric rank the better-rated case higher? (ties = 0.5)
 * 50% = coin flip, 100% = perfect agreement.
 *
 * IMPORTANT: labels are only valid against the renders that were rated. The
 * harness prints labeledAt vs runAt — if the converter changed in between,
 * regenerate the labeling page and relabel before trusting the numbers.
 *
 * Run: npm run research:concordance
 *      tsx tools/concordance.ts --labels <labels.json> --results <results.json>
 */
import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SUITE_OUTPUT, WILD_OUTPUT } from "./output-paths.js";

const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const METRIC_FIELDS = [
  "visualScore",
  "layoutFidelityScore",
  "matchPercent",
  "textContentFidelityScore",
  "engineScore",
] as const;

interface LabelFile {
  labeledAt?: string;
  labels: Record<string, { rating?: number; note?: string }>;
}

interface ResultsFile {
  runAt?: string;
  cases: Array<Record<string, unknown> & { name: string }>;
}

interface JoinedCase {
  name: string;
  rating: number;
  note?: string;
  metrics: Record<string, number | undefined>;
}

function joinLabels(labels: LabelFile, results: ResultsFile): JoinedCase[] {
  const joined: JoinedCase[] = [];
  for (const [name, label] of Object.entries(labels.labels)) {
    if (!label.rating) continue;
    const row = results.cases.find((c) => c.name === name);
    if (!row) continue;
    const metrics: Record<string, number | undefined> = {};
    for (const field of METRIC_FIELDS) {
      const value = row[field];
      metrics[field] = typeof value === "number" ? value : undefined;
    }
    joined.push({ name, rating: label.rating, note: label.note, metrics });
  }
  return joined;
}

interface ConcordanceResult {
  field: string;
  concordance: number;
  pairs: number;
  groupMeans: Record<number, number>;
  discordant: Array<{ better: string; worse: string; betterScore: number; worseScore: number }>;
}

function concordanceFor(cases: JoinedCase[], field: string): ConcordanceResult | undefined {
  const scored = cases.filter((c) => c.metrics[field] !== undefined);
  if (scored.length < 2) return undefined;

  let agree = 0;
  let total = 0;
  const discordant: ConcordanceResult["discordant"] = [];
  for (let i = 0; i < scored.length; i++) {
    for (let j = i + 1; j < scored.length; j++) {
      const a = scored[i]!;
      const b = scored[j]!;
      if (a.rating === b.rating) continue;
      const better = a.rating < b.rating ? a : b; // rating 1 = best
      const worse = a.rating < b.rating ? b : a;
      total += 1;
      const bs = better.metrics[field]!;
      const ws = worse.metrics[field]!;
      if (bs > ws) agree += 1;
      else if (bs === ws) agree += 0.5;
      else {
        discordant.push({ better: better.name, worse: worse.name, betterScore: bs, worseScore: ws });
      }
    }
  }
  if (total === 0) return undefined;

  const groupMeans: Record<number, number> = {};
  for (const rating of [1, 2, 3]) {
    const group = scored.filter((c) => c.rating === rating);
    if (group.length) {
      groupMeans[rating] =
        group.reduce((sum, c) => sum + c.metrics[field]!, 0) / group.length;
    }
  }

  discordant.sort((a, b) => a.betterScore - a.worseScore - (b.betterScore - b.worseScore));
  return { field, concordance: (agree / total) * 100, pairs: total, groupMeans, discordant };
}

async function reportCorpus(
  title: string,
  labelsPath: string,
  resultsPath: string,
): Promise<void> {
  if (!existsSync(labelsPath) || !existsSync(resultsPath)) {
    console.log(`  [${title}] skipped — missing ${!existsSync(labelsPath) ? labelsPath : resultsPath}`);
    return;
  }
  const labels: LabelFile = JSON.parse(await readFile(labelsPath, "utf-8"));
  const results: ResultsFile = JSON.parse(await readFile(resultsPath, "utf-8"));
  const joined = joinLabels(labels, results);

  console.log(`\n  [${title}] ${joined.length} labeled cases joined`);
  console.log(`    labels: ${labels.labeledAt ?? "?"} · run: ${results.runAt ?? "?"}`);
  if (labels.labeledAt && results.runAt && labels.labeledAt < results.runAt) {
    console.log(
      "    ⚠ labels PREDATE this run — renders may have changed; relabel before trusting these numbers",
    );
  }

  const header = `    ${"metric".padEnd(26)} ${"concord".padStart(8)} ${"pairs".padStart(6)}   group means (r1 / r2 / r3)`;
  console.log(header);
  for (const field of METRIC_FIELDS) {
    const r = concordanceFor(joined, field);
    if (!r) continue;
    const means = [1, 2, 3]
      .map((g) => (r.groupMeans[g] !== undefined ? r.groupMeans[g]!.toFixed(1) : "—"))
      .join(" / ");
    console.log(
      `    ${field.padEnd(26)} ${r.concordance.toFixed(1).padStart(7)}% ${String(r.pairs).padStart(6)}   ${means}`,
    );
  }

  const primary = concordanceFor(joined, "visualScore");
  if (primary && primary.discordant.length > 0) {
    console.log(`    worst discordant pairs (visualScore):`);
    for (const d of primary.discordant.slice(0, 5)) {
      console.log(
        `      human prefers ${d.better} (${d.betterScore.toFixed(1)}) over ${d.worse} (${d.worseScore.toFixed(1)})`,
      );
    }
  }
}

async function main(): Promise<void> {
  const argLabels = process.argv.indexOf("--labels");
  const argResults = process.argv.indexOf("--results");
  console.log("── Metric ↔ human-label concordance ──");

  if (argLabels > -1 && argResults > -1) {
    await reportCorpus(
      "custom",
      path.resolve(process.argv[argLabels + 1]!),
      path.resolve(process.argv[argResults + 1]!),
    );
    return;
  }

  await reportCorpus(
    "suite",
    path.join(REPO_ROOT, "internal", "research", "human-labels.json"),
    path.join(SUITE_OUTPUT, "results.json"),
  );
  await reportCorpus(
    "wild",
    path.join(REPO_ROOT, "internal", "research", "wild-labels.json"),
    path.join(WILD_OUTPUT, "results.json"),
  );
}

main();
