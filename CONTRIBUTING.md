# Contributing

## Repository layout

| Path | Purpose |
|------|---------|
| **`src/`** | Published library — `index.ts`, `browser.ts`, `converter.ts`, `converter/*` |
| **`tools/`** | Visual harness, benchmarks, scoring — **not** shipped on npm |
| **`examples/`** | Committed sample HTML, DOCX output, side-by-side previews |
| **`docs/`** | Dev documentation (scores, benchmarks, scoring methodology) |
| **`scripts/`** | Browser bundle build + pack smoke test |
| **`output/`** | Gitignored harness artifacts (`suite/`, `benchmark/`, `showcase/`, `css-cascade/`, `guards/`) |

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

Everything below groups into four tiers by what it's actually for — a scored regression, a binary invariant, or a check on the *metric* rather than the converter. Prefixes match the group (`score:*`, `guard:*`, `research:*`); `showcase` and the build/setup commands above don't fit any of the four and stand on their own.

### 1. Core score — the primary signal, run every iteration

```bash
npm run score:suite            # full regression suite (cases: tools/generator.ts; needs LO + Chromium)
npm run score:suite:priority   # fast subset of the same cases, for the dev loop
npm run score:suite:strict     # full suite, zero-tolerance pixel regression (CI)
npm run docs:sync              # regenerate docs/TEST-SCORES.md + docs/BENCHMARK.md from the JSON above
```

### 2. Comparative scoring — periodic, feeds docs via `docs:sync`, not part of every dev loop

```bash
npm run score:benchmark              # dom-docx vs html-to-docx + TurboDocx, same harness
npm run score:benchmark -- turbodocx
npm run score:benchmark -- html-to-docx
npm run score:style-source           # inline vs computed-oracle vs computed-native (needs a fresh score:suite baseline)
npm run score:css-cascade            # stylesheet / class selector suite
npm run score:calibration            # pipeline-noise ceiling (no conversion); add -- --full for all cases
```

### 3. Guards — binary pass/fail invariants

Each writes a result to `output/guards/<id>.json`; `docs:sync` reads whichever are present into a single status table. `guard:inline`, `guard:config`, and `guard:pack-smoke` need no Playwright/LibreOffice and run in CI; the other two need Playwright and are maintainer-only.

```bash
npm run guard:inline             # default styleSource vs explicit "inline", byte-equivalence (CI)
npm run guard:config             # ConvertOptions → correct OOXML (CI)
npm run guard:pack-smoke         # npm tarball installs/converts without Playwright (CI)
npm run guard:computed-parity    # computed-oracle vs computed-native, byte-identical (maintainer-only)
npm run guard:browser-parity     # browser bundle vs Node computed-native (maintainer-only)
```

### 4. Research tools — validate the metric itself, not the converter

Maintainer-only, run occasionally when tuning or auditing the scoring metric — human labeling, real-world corpora, renderer-drift spot checks. Not part of what gates a release.

```bash
npm run research:word-spotcheck   # LibreOffice- vs Word-rendered score delta (needs Word on macOS)
npm run research:multipage        # 14-section stress doc, pagination correctness
npm run research:novel            # randomly generated/seeded HTML, structural robustness
npm run research:wild-corpus       # build the real-world corpus (tsx tools/wild-corpus-build.ts)
npm run research:wild             # score dom-docx against that corpus
npm run research:label            # hand-labeling UI over suite renders (human ground truth)
npm run research:concordance      # does the visual metric rank cases the way humans do?
```

### Showcase — demo generation, not a test

```bash
npm run showcase   # 9 realistic documents → examples/ + output/showcase/ (not part of the regression loop)
```

## Adding regression cases

Edit `tools/generator.ts` (loop cases) — give each a `description`, which is what `docs:sync` uses to render TEST-SCORES.md; a case without one is flagged rather than silently missing. For rich demos, use `tools/showcase.ts` and run `npm run showcase`.

Scoring methodology: [docs/SCORING.md](./docs/SCORING.md). HTML authoring guide: [AGENTS.md](./AGENTS.md). Maintainer backlog: `internal/TODO.md` (gitignored, local only).

## Release to npm

CI (`.github/workflows/ci.yml`) runs on every push/PR: typecheck, build, browser bundle, inline guard (all cases), config tests, and pack smoke — no Playwright or LibreOffice.

Publishing (`.github/workflows/publish.yml`) runs when a semver tag is pushed:

1. Set `"version"` in `package.json` (e.g. `0.1.1`).
2. Commit and push to `main`.
3. Tag and push: `git tag v0.1.1 && git push origin v0.1.1`

The tag must match the package version (`v0.1.1` ↔ `"0.1.1"`). The workflow re-runs the CI checks, then `npm publish --provenance --access public`.

Add an **`NPM_TOKEN`** repository secret (npmjs.com → Access Tokens → granular, publish-only for `dom-docx`). Visual regression (`npm run score:suite`) stays manual — it needs Chromium + LibreOffice and is not run in GitHub Actions.
