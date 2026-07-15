# Contributing

## Repository layout

| Path            | Purpose                                                                                       |
| --------------- | --------------------------------------------------------------------------------------------- |
| **`src/`**      | Published library — `index.ts`, `browser.ts`, `converter.ts`, `converter/*`                   |
| **`tools/`**    | Visual harness, benchmarks, scoring — **not** shipped on npm                                  |
| **`examples/`** | Committed sample HTML, DOCX output, side-by-side previews                                     |
| **`docs/`**     | Dev documentation (scores, benchmarks, scoring methodology)                                   |
| **`scripts/`**  | Browser bundle build + pack smoke test                                                        |
| **`output/`**   | Gitignored harness artifacts (`suite/`, `benchmark/`, `showcase/`, `css-cascade/`, `guards/`) |

The npm package includes only `dist/`, `README.md`, `LICENSE`, `API.md`, and `examples/` (see `"files"` in `package.json`).

Paths are centralized in `tools/output-paths.ts`.

## Prerequisites

- **Node.js ≥ 20** (see `.nvmrc`)
- **Playwright Chromium** — `npm run setup` (harness + `styleSource: "computed"`)
- **LibreOffice** (`soffice`) — PDF rasterization for visual scoring only; not needed for `npm run build` or inline conversion

## Test and score commands

```bash
npm install
npm run setup          # playwright install chromium
npm run typecheck
npm run build          # dist/ for npm pack
```

Everything below groups into four tiers by what it's actually for — a scored regression, a binary invariant, or a check on the _metric_ rather than the converter. Prefixes match the group (`score:*`, `guard:*`, `research:*`); `showcase` and the build/setup commands above don't fit any of the four and stand on their own.

### 1. Core score — the primary signal, run every iteration

- **`score:suite`** — the core loop (`tools/validator.ts`). Converts all cases from `tools/generator.ts` to DOCX, screenshots the source HTML (Chromium) and the converted DOCX (LibreOffice → PDF → PNG), scores layout fidelity + editability + compile speed, writes `output/suite/results.json`. Everything else in this project is judged against this number. Run one case with `SUITE_ONLY=<name> npm run score:suite`.
- **`score:suite:priority`** — same script with `--priority`, only the ~10 cases named in `PRIORITY_LOOP_CASE_NAMES` (`tools/generator.ts`), for a fast dev loop. Writes a separate `results-priority.json` that `docs:sync` does **not** read — this is for local iteration, not the record.
- **`score:suite:strict`** — same full suite with `--strict-visual`, a zero-tolerance pixel-regression flag for CI-grade runs. Writes to the same `results.json` as plain `score:suite`.
- **`docs:sync`** — reads the JSON these scripts write and regenerates `docs/TEST-SCORES.md` + `docs/BENCHMARK.md`, so per-case tables and pass counts can't drift from what the harness actually measured. Optional sections (style-source, CSS cascade, guard status) are preserved verbatim from the last run if their source JSON is missing, rather than regressing to a placeholder.

```bash
npm run score:suite
npm run score:suite:priority
npm run score:suite:strict
npm run docs:sync
```

### 2. Comparative scoring — periodic, feeds docs via `docs:sync`, not part of every dev loop

- **`score:benchmark`** — scores `html-to-docx` and `@turbodocx/html-to-docx` through the _identical_ harness (same cases, same scoring) so dom-docx's numbers are directly comparable. Reads dom-docx's own baseline from `output/suite/results.json` to compute deltas — run `score:suite` first. Takes an optional arg to run just one library.
- **`score:style-source`** — inline vs computed-oracle vs computed-native, run against the **main suite** (`tools/generator.ts`), which is 100% inline `style=""` with no `<style>` blocks or classes at all. Since inline resolution already sees everything on this suite, this benchmark isn't measuring "does computed find more" — there's nothing more to find. It measures two other things: (1) whether switching to computed resolution regresses anything on ordinary inline content (the three lanes land within ~0.35pp of each other, which is the proof it doesn't), and (2) the performance/architecture cost between the two ways of _doing_ computed resolution — **oracle** spawns its own Playwright page per call (~630ms), **native** reads an already-rendered page (~27ms), same output. Needs a fresh `score:suite` baseline.
- **`score:css-cascade`** — inline vs computed, run against a **separate, purpose-built fixture** (`tools/css-cascade-cases.ts`) where styles live in `<style>` blocks and classes with **no** matching inline style on the styled elements (one deliberate exception, `stylesheet-inline-wins`, proving inline still overrides class when both are present). This is where the inline resolver is actually blind — it scores far lower here because it structurally cannot see stylesheet rules — so this is the capability/correctness benchmark: does computed resolution correctly implement the cascade the inline path can't see at all.

  **`score:style-source` and `score:css-cascade` are not redundant** even though both compare inline vs computed — they run against different fixtures to answer different questions: one checks for regression + measures perf cost using content both paths see identically, the other checks a real capability gap using content deliberately invisible to one of the two paths.

- **`score:calibration`** — pushes the _same_ HTML through both pipeline sides (Chromium screenshot vs Chromium-printed PDF) with **no conversion involved**, to measure how much score deficit is pipeline/rendering noise vs an actual conversion defect. Add `-- --full` to run all cases instead of the priority subset.

```bash
npm run score:benchmark
npm run score:benchmark -- turbodocx
npm run score:benchmark -- html-to-docx
npm run score:style-source
npm run score:css-cascade
npm run score:calibration
npm run score:calibration -- --full
```

### 3. Guards — binary pass/fail invariants

Each writes a result to `output/guards/<id>.json` (via `tools/guard-result.ts`, or an inlined equivalent in `scripts/pack-smoke.mjs` since it runs via plain `node`); `docs:sync` reads whichever are present into a single status table in BENCHMARK.md. `guard:inline`, `guard:config`, `guard:toc-slot`, `guard:internal-href`, `guard:document-canvas`, `guard:image-spacing`, and `guard:pack-smoke` need no Playwright/LibreOffice and run in CI (`guard:document-canvas` also runs an optional Playwright section when Chromium is installed); the remaining ones need Playwright/LibreOffice and are maintainer-only.

- **`guard:inline`** — converts every case via default options and via explicit `{ styleSource: "inline" }`; asserts byte-identical normalized `word/*.xml`. Catches accidental drift in the default path.
- **`guard:config`** — a battery of named assertions (one per `DocumentConfig` field — `pageSize`, `margins`, `defaultFont`, `metadata`, `headerHtml`/`footerHtml`, `pageNumber`, `lang`/`direction`, `coverHtml`, `tocHtml`, …) that each produces the correct OOXML. **Runs every assertion through both public entries** — the Node `convertHtmlToDocx` and the browser `convertHtmlToDocxUint8Array` (its inline path runs headless) — because option forwarding is duplicated per entry and has drifted before (a new option reaching one entry but not the other, with no compiler error).
- **`guard:pack-smoke`** — `npm pack`s the real tarball, installs it into a clean temp project, and asserts: Playwright isn't a hard/optional dependency, the browser bundle files ship in the tarball, and the library/CLI/browser entry points each actually convert HTML to a valid `.docx`.
- **`guard:computed-parity`** — computed-oracle and computed-native must emit byte-identical OOXML for identical HTML. This is what backs the claim that the "native" lane (a Playwright-driven stand-in used in the test harness) faithfully represents the real browser deployment.
- **`guard:browser-parity`** — chained script (`build:browser && browser-spike.ts && browser-build-parity.ts`) asserting the esbuild browser IIFE bundle (`dom-docx/browser`) produces byte-identical output to the Node computed-native path.
- **`guard:page-break`** — structural page-break test (OOXML `w:pageBreakBefore` + multi-page PDF). Not part of the visual suite — explicit breaks can't be scored with single-page pixel compare.
- **`guard:toc-slot`** — structural test for the `tocHtml` option (caller-provided table of contents). Asserts the slot fragment renders after the cover and before the body; that its in-page links (`<a href="#id">`) become internal hyperlinks pointing at real `id` bookmarks in the body; the cover → toc → body ordering; and OOXML schema validity of the whole document. (In-page linking itself is covered by `guard:internal-href`.)
- **`guard:internal-href`** — structural test for same-document links (`href="#id"`). Asserts internal hyperlinks (`w:hyperlink w:anchor`), matching bookmarks on `id` / legacy `a[name]` targets, URI-decoded fragments, that external URLs still use relationships, and OOXML schema validity. CI — no Playwright/LibreOffice.
- **`guard:document-canvas`** — DOCX exports must not follow dark-theme text colors. Asserts near-white computed `color` with no dark fill is dropped (so Word doesn’t get invisible light-on-white runs), while light text on a dark shaded block is kept; plus an optional Playwright check that the Node computed path forces `prefers-color-scheme: light`. CI for the remap/OOXML checks.
- **`guard:image-spacing`** — image paragraphs keep a floor of before/after spacing so figures don’t smash into the next heading or body when web margins were zeroed (flex/grid gap layouts). Skipped inside flex cards. CI — no Playwright/LibreOffice.

```bash
npm run guard:inline             # CI
npm run guard:config             # CI
npm run guard:toc-slot           # CI
npm run guard:internal-href      # CI
npm run guard:document-canvas    # CI (Playwright section optional)
npm run guard:image-spacing      # CI
npm run guard:pack-smoke         # CI
npm run guard:page-break         # structural page breaks (needs LibreOffice)
npm run guard:computed-parity    # maintainer-only, needs Playwright
npm run guard:browser-parity     # maintainer-only, needs Playwright + a built bundle
```

### 4. Research tools — validate the metric itself, not the converter

Maintainer-only, run occasionally when tuning or auditing the scoring metric rather than the converter. Not part of what gates a release.

- **`research:word-spotcheck`** — the loop scores against LibreOffice, but the real consumer is Microsoft Word, and the two disagree on some layout rules. Renders 5 anchor cases through **both** renderers and reports the adjusted-visual delta, quantifying how much of the metric is LibreOffice-specific artifact vs a real conversion defect. Needs Word on macOS; skips cleanly when unavailable.
- **Pages / Google Docs** — not gated; see `internal/TODO.md` → _Viewer compatibility_ (Pages can mis-render page breaks; Docs import is lossy).
- **`research:multipage`** — a 14-section stress document exercising page-break/pagination correctness, a dimension the main suite (mostly single-page cases) doesn't cover.
- **`research:novel`** — randomly generated, seeded HTML structures, looking for structural-robustness bugs that a curated, hand-written suite wouldn't happen to hit.
- **`research:wild-corpus`** / **`research:wild`** — build, then score, a corpus of real-world pages the converter was **not** tuned on (email templates, legacy table layouts, wiki tables, book prose, rendered markdown) — a check against overfitting to the curated suite.
- **`research:label`** — generates a blind hand-labeling UI over suite renders (scores hidden, so the labeler isn't anchored by the metric) to build the human ground-truth ratings the other research tools validate against.
- **`research:concordance`** — cross-references those human labels against the metric's own scores: for case pairs where humans rated one clearly better, does the metric rank them the same way? This validates that the scoring formula tracks human judgment, not just its own internal consistency.

```bash
npm run research:word-spotcheck
npm run research:multipage
npm run research:novel
npm run research:wild-corpus
npm run research:wild
npm run research:label
npm run research:concordance
```

### Showcase — demo generation, not a test

```bash
npm run showcase   # 9 realistic documents → examples/ + output/showcase/ (not part of the regression loop)
```

## Adding regression cases

Edit `tools/generator.ts` (loop cases) — give each a `description`, which is what `docs:sync` uses to render TEST-SCORES.md; a case without one is flagged rather than silently missing. For rich demos, use `tools/showcase.ts` and run `npm run showcase`.

Scoring methodology: [docs/SCORING.md](./docs/SCORING.md). HTML authoring guide: [AGENTS.md](./AGENTS.md). Maintainer backlog: `internal/TODO.md` (gitignored, local only).

## Release to npm

CI (`.github/workflows/ci.yml`) runs on every push/PR: typecheck, build, browser bundle, `guard:inline`, `guard:config`, `guard:toc-slot`, `guard:internal-href`, `guard:document-canvas`, `guard:image-spacing`, and `guard:pack-smoke` — no Playwright or LibreOffice.

Publishing uses **npm Trusted Publishing (OIDC)** — there is no `NPM_TOKEN` secret. The publish workflow (`.github/workflows/publish.yml`) runs when a semver tag matching `v*.*.*` is pushed to GitHub. It verifies the tag matches `"version"` in `package.json`, re-runs the same CI checks as above, then runs `npm publish --provenance --access public` (`prepack` builds the library and browser bundle).

### Typical release flow

From `main` with a clean working tree:

```bash
npm version patch   # or minor / major — bumps package.json, commits, creates vX.Y.Z tag
git push origin main --follow-tags
```

`npm version` updates `package.json` (and `package-lock.json` if present), creates a release commit, and tags it (e.g. `v0.1.3` ↔ `"0.1.3"`). Pushing the tag triggers the workflow; the tag **must** match `package.json` or publish fails.

To bump without auto-commit/tag (manual control):

```bash
npm version patch --no-git-tag-version
# edit CHANGELOG if you keep one, then:
git add package.json package-lock.json
git commit -m "Release 0.1.3"
git tag v0.1.3
git push origin main && git push origin v0.1.3
```

Visual regression (`npm run score:suite`) is maintainer-local — Chromium + LibreOffice — and is not run in GitHub Actions before publish.
