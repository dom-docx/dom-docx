# Contributing

## Repository layout

| Path | Purpose |
|------|---------|
| **`src/`** | Published library — `index.ts`, `browser.ts`, `converter.ts`, `converter/*` |
| **`tools/`** | Visual harness, benchmarks, scoring — **not** shipped on npm |
| **`examples/`** | Committed sample HTML, DOCX output, side-by-side previews |
| **`docs/`** | Dev documentation (scores, benchmarks, scoring methodology) |
| **`scripts/`** | Browser bundle build + pack smoke test |
| **`output/`** | Gitignored harness artifacts (`suite/`, `benchmark/`, `showcase/`, `css-cascade/`) |

The npm package includes only `dist/`, `README.md`, `LICENSE`, `API.md`, and `examples/` (see `"files"` in `package.json`).

Paths are centralized in `tools/output-paths.ts`.

## Prerequisites

- **Node.js ≥ 20** (see `.nvmrc`)
- **Playwright Chromium** — `npm run setup` (harness + `styleSource: "computed"`)
- **LibreOffice** (`soffice`) — PDF rasterization for visual scoring only; not needed for `npm run build` or inline conversion

## Common commands

```bash
npm install
npm run setup          # playwright install chromium
npm run typecheck
npm run build          # dist/ for npm pack
npm run test:inline-guard   # fast: inline path unchanged
npm run test:config         # ConvertOptions OOXML checks
npm run test:suite           # full visual regression suite (cases: tools/generator.ts; needs LO + Chromium)
npm run test:suite:priority  # fast subset of the same cases
# Run one case: SUITE_ONLY=rasterize-in-place-chart npm run test:suite
npm run test:benchmark      # vs html-to-docx + TurboDocx
npm run test:showcase       # refresh examples/
npm run test:pack-smoke     # verify npm tarball installs without Playwright
```

Benchmark sub-commands use npm args (no separate scripts):

```bash
npm run test:benchmark -- turbodocx
npm run test:benchmark -- html-to-docx
npm run test:calibration -- --full   # all cases (default is the priority subset)
```

## Maintainer-only harness commands

These are not npm scripts — run directly when debugging a specific lane:

```bash
npm run build:browser && tsx scripts/browser-spike.ts
npm run build:browser && tsx scripts/browser-spike.ts && tsx tools/browser-build-parity.ts
tsx tools/computed-parity-guard.ts
tsx tools/style-source-benchmark.ts
tsx tools/css-cascade-runner.ts
tsx tools/novel-runner.ts
tsx tools/word-spotcheck.ts      # Word on macOS
tsx tools/multipage-test.ts
```

## Adding regression cases

Edit `tools/generator.ts` (loop cases). For rich demos, use `tools/showcase.ts` and run `npm run test:showcase`.

Scoring methodology: [docs/SCORING.md](./docs/SCORING.md). HTML authoring guide: [AGENTS.md](./AGENTS.md). Maintainer backlog: `internal/TODO.md` (gitignored, local only).

## Release to npm

CI (`.github/workflows/ci.yml`) runs on every push/PR: typecheck, build, browser bundle, inline guard (all cases), config tests, and pack smoke — no Playwright or LibreOffice.

Publishing (`.github/workflows/publish.yml`) runs when a semver tag is pushed:

1. Set `"version"` in `package.json` (e.g. `0.1.1`).
2. Commit and push to `main`.
3. Tag and push: `git tag v0.1.1 && git push origin v0.1.1`

The tag must match the package version (`v0.1.1` ↔ `"0.1.1"`). The workflow re-runs the CI checks, then `npm publish --provenance --access public`.

Add an **`NPM_TOKEN`** repository secret (npmjs.com → Access Tokens → granular, publish-only for `dom-docx`). Visual regression (`npm run test:suite`) stays manual — it needs Chromium + LibreOffice and is not run in GitHub Actions.
