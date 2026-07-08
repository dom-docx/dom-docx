import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { GUARDS_OUTPUT } from "./output-paths.js";

/**
 * Uniform result shape for the binary pass/fail `guard:*` scripts (inline path,
 * config options, computed parity, browser-bundle parity, pack smoke). Each guard
 * writes one of these to `output/guards/<id>.json`; `tools/docs-sync.ts` reads
 * whichever are present and renders a single "Guard status" table, so pass/fail
 * counts in docs can't go stale the way hand-typed prose did.
 */
export interface GuardResult {
  /** Stable id, also the JSON filename (without extension). */
  id: string;
  /** Short human label for the docs table, e.g. "Inline path". */
  label: string;
  passed: number;
  total: number;
  ok: boolean;
  /** One-line description of what "passed" means for this guard, e.g. "byte-identical". */
  unit: string;
  ranAt: string;
  /** Command a maintainer runs to reproduce this guard. */
  command: string;
}

export async function writeGuardResult(
  result: Omit<GuardResult, "ranAt">,
): Promise<void> {
  await mkdir(GUARDS_OUTPUT, { recursive: true });
  const payload: GuardResult = { ...result, ranAt: new Date().toISOString() };
  await writeFile(
    path.join(GUARDS_OUTPUT, `${result.id}.json`),
    `${JSON.stringify(payload, null, 2)}\n`,
    "utf-8",
  );
}
