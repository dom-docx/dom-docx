import type { TestCase } from "./generator.js";

/** Stylesheet / class / selector cases — inline resolver ignores `<style>` and classes. */
export interface CssCascadeCase extends TestCase {
  /** Why this case exists in the cascade suite. */
  notes: string;
  /** Minimum adjusted visual for computed path (default 70). */
  computedMinVisual?: number;
  /** Computed must beat inline by at least this many pp (default 5). Use 0 to skip. */
  minComputedAdvantage?: number;
}

const CSS_CASCADE_CASES: CssCascadeCase[] = [
  {
    name: "stylesheet-p-color",
    notes: "Element selector sets paragraph color; no inline style on p.",
    html: `
      <style>
        p { color: #e63946; font-size: 18px; }
      </style>
      <p>Red paragraph text from a stylesheet rule.</p>
    `,
    computedMinVisual: 85,
    minComputedAdvantage: 3,
  },
  {
    name: "stylesheet-class-banner",
    notes: "Class selector for shaded banner block.",
    html: `
      <style>
        .banner {
          background: #eaeaea;
          padding: 10px 16px;
          margin-bottom: 12px;
        }
        .banner h1 {
          margin: 0;
          font-size: 20px;
          color: #1a1a2e;
        }
      </style>
      <div class="banner">
        <h1>Stylesheet Banner Title</h1>
      </div>
      <p>Body text below the banner.</p>
    `,
    computedMinVisual: 88,
    minComputedAdvantage: 1,
  },
  {
    name: "stylesheet-descendant-table",
    notes: "Descendant selector styles header row via class on tr.",
    html: `
      <style>
        table.data tr.header td {
          background: #1a1a2e;
          color: #f1faee;
          font-weight: bold;
        }
        table.data td.num { text-align: right; }
      </style>
      <table class="data" border="1" cellpadding="8" style="border-collapse:collapse;width:100%">
        <tr class="header">
          <td>Item</td>
          <td class="num">Value</td>
        </tr>
        <tr>
          <td>Alpha</td>
          <td class="num">100</td>
        </tr>
        <tr>
          <td>Beta</td>
          <td class="num">250</td>
        </tr>
      </table>
    `,
    computedMinVisual: 90,
    minComputedAdvantage: 40,
  },
  {
    name: "stylesheet-strong-em",
    notes: "Type selectors set strong/em presentation inside styled paragraph.",
    html: `
      <style>
        p.lead { color: #457b9d; font-size: 16px; }
        p.lead strong { color: #1a1a2e; font-weight: 700; }
        p.lead em { color: #e76f51; font-style: italic; }
      </style>
      <p class="lead">
        Lead sentence with <strong>dark bold</strong> and <em>orange italic</em> from CSS.
      </p>
    `,
    computedMinVisual: 88,
    minComputedAdvantage: 2,
  },
  {
    name: "stylesheet-id-title",
    notes: "ID selector sets heading size and color; subtitle via class.",
    html: `
      <style>
        #doc-title {
          color: #1a1a2e;
          font-size: 22px;
          margin: 0 0 8px 0;
        }
        .subtitle { color: #666; font-size: 13px; }
      </style>
      <h1 id="doc-title">Document Title</h1>
      <p class="subtitle">Prepared for stylesheet cascade testing.</p>
    `,
    computedMinVisual: 80,
    minComputedAdvantage: 0,
  },
  {
    name: "stylesheet-list-classes",
    notes: "Class on list items for status colors and row shading.",
    html: `
      <style>
        ul.tasks { margin-top: 8px; }
        li.done { color: #2a9d8f; background: #cfc; padding: 4px 8px; }
        li.blocked { color: #e63946; background: #fcc; padding: 4px 8px; }
      </style>
      <p>Action items:</p>
      <ul class="tasks">
        <li class="done">Ship feature A</li>
        <li class="blocked">Fix regression B</li>
        <li>Write documentation</li>
      </ul>
    `,
    computedMinVisual: 85,
    minComputedAdvantage: 3,
  },
  {
    name: "stylesheet-section-theme",
    notes: "Section class applies dark card chrome; inline must render it unstyled (leak check), computed applies the full theme.",
    html: `
      <style>
        section.card {
          background: #457b9d;
          color: #f1faee;
          padding: 16px 20px;
          margin: 12px 0;
        }
        section.card h2 {
          color: #ffdd57;
          font-size: 18px;
          margin: 0 0 8px 0;
        }
        section.card p { margin: 0; }
      </style>
      <section class="card">
        <h2>Card Title</h2>
        <p>Card body styled entirely from the stylesheet.</p>
      </section>
    `,
    computedMinVisual: 50,
    minComputedAdvantage: 40,
  },
  {
    name: "stylesheet-inline-wins",
    notes: "Inline style overrides class color — both paths should honor inline on that node.",
    html: `
      <style>
        .muted { color: #666; }
        .accent { color: #2a9d8f; }
      </style>
      <p class="muted">Muted gray from class.</p>
      <p class="muted" style="color:#e63946">Inline red overrides class gray.</p>
      <p class="accent">Teal from class alone.</p>
    `,
    computedMinVisual: 85,
    minComputedAdvantage: 0,
  },
  {
    name: "stylesheet-blockquote-rule",
    notes: "Element rule styles blockquote: left bar + background + padding together.",
    html: `
      <style>
        blockquote {
          border-left: 4px solid #333;
          padding: 8px 12px;
          margin: 8px 0;
          color: #444;
          background: #f5f5f5;
        }
      </style>
      <blockquote>
        <p>Quote styled only via stylesheet blockquote rule.</p>
      </blockquote>
    `,
    computedMinVisual: 80,
    minComputedAdvantage: 5,
  },
  {
    name: "stylesheet-callout-class",
    notes: "Class rule styles a callout div: left bar + background + padding together.",
    html: `
      <style>
        .callout {
          border-left: 4px solid #333;
          padding: 8px 12px;
          margin: 8px 0;
          color: #444;
          background: #f5f5f5;
        }
      </style>
      <div class="callout">
        <p>Callout styled only via stylesheet class rule.</p>
      </div>
    `,
    computedMinVisual: 80,
    minComputedAdvantage: 5,
  },
];

export function generateCssCascadeCases(): CssCascadeCase[] {
  return CSS_CASCADE_CASES.map((c) => ({ ...c }));
}

export const DEFAULT_COMPUTED_MIN_VISUAL = 70;
export const DEFAULT_MIN_COMPUTED_ADVANTAGE = 5;
