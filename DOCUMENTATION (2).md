# Patent Portfolio Bundling Template — Project Documentation

> **Living document.** This file is updated whenever a feature is added, changed, or removed. See the [Version History](#version-history) section for the changelog.

---

## Table of contents

1. [Overview](#overview)
2. [Version history](#version-history)
3. [File inventory](#file-inventory)
4. [Conceptual framework](#conceptual-framework)
5. [Quick reference — Bundle catalog (summary)](#quick-reference--bundle-catalog-summary)
6. [Sheet-by-sheet reference](#sheet-by-sheet-reference)
7. [Configuration system](#configuration-system)
8. [Short-lifecycle portfolio context (v5)](#short-lifecycle-portfolio-context-v5)
9. [PatSeer Import sheet (v6)](#patseer-import-sheet-v6)
10. [Buyer archetype framework](#buyer-archetype-framework)
11. [Bundle catalog — full detail](#bundle-catalog--full-detail)
12. [Threshold catalog](#threshold-catalog)
13. [Quality gates](#quality-gates)
14. [Bundle composition strategies](#bundle-composition-strategies)
15. [Workflows / how-to guides](#workflows--how-to-guides)
16. [Maintenance notes](#maintenance-notes)
17. [Glossary](#glossary)

---

## Overview

### What this project is

A technical-attribute-based framework, delivered as an Excel workbook, for grouping patents in a portfolio into sellable bundles for **outright sale**. The framework is intentionally limited to **technical aspects only** — no patent valuation, no pricing, no financial modeling. The goal is to decide *which patents belong together* and *why a buyer would want them as a unit*, based on technical attributes alone.

The project has two tightly connected workstreams that share a common asset-analysis foundation:

1. **Patent bundling** — scoring, routing, and quality-scoring patents into sellable packages (the Excel workbook)
2. **Value proposition writing** — writing IAM Market listing copy for those packages (the Value Proposition Framework and Buyer Profiles document)

### Problem it solves

A typical patent portfolio shortlisted for sale contains patents with very different characteristics — pioneer patents and improvement patents, broad and narrow claims, strong and weak claims, vulnerable and bulletproof patents, mature and emerging tech areas, multiple jurisdictions, varying remaining terms. Selling them as a single lot is rarely optimal. Selling them one-by-one is slow and leaves volume value on the table. The right answer is **structured bundling** — composing each bundle around a coherent technical narrative that targets a specific buyer profile.

### What it produces

For a given portfolio of patents (entered as rows in a spreadsheet), the framework:
- Scores every patent across 42 technical attributes (Groups A–I)
- Auto-routes each patent into one or more of 33 predefined bundle types based on logical rules
- Quality-scores each resulting bundle on coverage depth, detectability, term, invalidity exposure, continuation-optionality, and more
- Lets you toggle bundles, gates, and thresholds for a given analysis
- Lets you save and load named configurations as presets
- Provides three short-lifecycle portfolio presets tuned to assets with ≤8 years remaining term
- Accepts PatSeer export data directly via the PatSeer Import sheet, auto-deriving 13 of 42 attributes and computing all 33 bundle qualifications inline

---

## Version history

### v1 — Foundation
- 7 sheets: README, Attribute Dictionary, Patent Portfolio, Bundle Rules, Bundle Assignment, Bundle Quality Scorecard, Sample Bundles
- 26 technical attributes across 7 groups (A–G)
- 21 bundle types
- 12 sample patents pre-populated
- Bundle Assignment auto-computes TRUE/FALSE qualification per patent per bundle
- Scorecard auto-computes 8 quality metrics per bundle
- Data validation dropdowns on all categorical fields

### v2 — Quality, vulnerability, and market signals
- Added **Group H — Patent Quality & Vulnerability** (10 attributes): claim strength, prior-art exposure, prosecution risk, divided-infringement risk, forward citations, backward citations, litigation/PTAB history, chain-of-title, EoU availability, encumbrance status
- Added **Group I — Market/Buyer Signals** (4 attributes): product-mapping confidence, implementation maturity, adjacent-market re-read, workaround complexity
- Added **12 new bundle types (22–33)**: Anchor-and-Halo, Picket-Fence, Strong-Core+Tail, Continuation-Live, EoU-Backed, Battle-Tested, Clean-Title, High-Citation, Adjacent Re-Read, Salvage Volume, Pre-Expiry, Provenance-Coherent
- Added **5 quality-gate columns** on Bundle Quality Scorecard: Weakest H1, Invalidity Exposure %, EoU-ready %, Survived %, Continuation-optionality %
- Added **Bundle Composition Strategy sheet** documenting 10 compositional approaches
- Populated v1 sample patents with realistic H/I attribute values

### v3 — Configurability
- New **Configuration sheet**: Active Preset dropdown, Edit Mode toggle, live summary line, parameter table with 33 bundle toggles + 18 thresholds + 5 gate toggles, each with Active Value / From Preset / Manual Override columns
- New **Presets sheet** with 6 starter presets (All ON, NPE / Counter-Assertion, Operating Company / FTO, Defensive Aggregator, Standards Licensee, EV Powertrain Sale) + 4 empty Custom Preset slots
- Bundle Assignment formulas rewired — wrapped in `IF(<bundle_enabled>="No","",...)`. Disabled columns show `[DISABLED]` and are grayed out
- Bundle Quality Scorecard rewired — disabled bundles/gates gray out similarly
- 18 thresholds made tunable via Configuration sheet cell references

### v4 — Attribute Derivation Procedure
- New **Attribute Procedure sheet** (inserted between Attribute Dictionary and Patent Portfolio):
  - Field-by-field procedure for all 42 attributes
  - Recommended 10-step workflow (data pull → claim reading → SEP check → detectability → product mapping → geography → themes → quality review → EoU audit → market signals)
  - Master per-attribute procedure table: where to find / how to derive / data source for each attribute
  - Free vs. paid data source reference (19 tools documented, including PatSeer, Google Patents, Espacenet, Lex Machina, TechInsights)
  - Practical notes on calibration and effort allocation
- AI prompt integration: `AI_Prompts_for_Attribute_Scoring.md` delivered as companion file with ready-to-use LLM prompts covering 20 attributes across 7 prompt blocks (A2, A3–A5, B1, C2+H4, D1–D3, H1+H4, I2–I4)

### v5 — Short-Lifecycle Portfolio Context
- **Three new presets** added to Presets sheet (columns 13–15), fully wired into the Configuration system:

  | Preset | PreExpiry window | Anchor H1 floor | Strength_term_min | Active bundles |
  |---|---|---|---|---|
  | Short Lifecycle — Critical (1-2yr) | 1–2 yr | **3** (raised) | 1 | B11–13, 19, 22, 26–28, 31–32 |
  | Short Lifecycle — Monetization (3-5yr) | 3–5 yr | 2 | 3 | Adds B5, 15–17, 24 |
  | Short Lifecycle — Strategic (6-8yr) | 6–8 yr | 2 | **6** (raised) | Adds B1, 3, 20, 29–30, 33; drops B31 |

- **30 new sample patents** (SL-001 to SL-030) added to Patent Portfolio:
  - 10 per term band (Critical / Monetization / Strategic)
  - Spanning 5 technology domains: wireless, battery/materials, edge AI, semiconductors, sensors
  - Each pre-scored across all 42 attributes with realistic values
  - Band separator rows in Patent Portfolio and Bundle Assignment: color-coded (red/amber/green)

- **Portfolio Context Guide sheet** — new reference sheet (position 3) with 7 sections:
  1. Term band overview table (3 bands × 7 dimensions)
  2. Step-by-step how-to for using the presets
  3. Critical signals grid (10 key attributes × 3 bands, with why-it-matters per band)
  4. Deal killer table with recommended actions per band
  5. IAM Market language register by band (urgency framing, value hook, buyer language, what to avoid)
  6. 3-tier composition strategy (Tier A: EoU-backed anchors / Tier B: program-starters / Tier C: salvage)
  7. Full 30-patent indexed reference (ID, title, band, key signals, tier)

- Configuration dropdown extended to cover columns C–O (15 presets total)
- Bundle Assignment extended for all 30 new patents with formulas matching the existing pattern

### v6 — PatSeer Import Sheet
- New **PatSeer Import sheet** (second sheet, after README) — three-zone layout:

  **Zone 1 — Column Mapping Table (rows 10–36):**
  - 25 rows mapping PatSeer export column headers to v5 attribute codes
  - Columns: PatSeer header (editable, blue) | v5 attribute code | transform rule | notes
  - Pre-populated with standard PatSeer export field names
  - Edit col A if your PatSeer headers differ — all formulas update automatically

  **Zone 2 — Paste Area (cols F–AD, rows 39 header + 40–239 data):**
  - 200-row paste zone; paste PatSeer Excel export directly
  - Cols A–E: manual-input helpers (row number, notes, E3 override, E2 override, E5 override) — amber background

  **Zone 3 — Mapped Output + Bundle Results (cols AE–DD):**
  - 42 attribute columns (AE–BT): teal header = auto-filled from PatSeer; amber header = manual scoring required
  - Completeness % column (BU): traffic-light conditional formatting (green ≥60%, amber 30–60%, red <30%)
  - 33 bundle qualification columns (BW–DC): TRUE/FALSE, green/red conditional formatting, respects active Configuration preset
  - Total qualifying bundle count column (DD)

- **PatSeer → v5 attribute coverage:**

  | v5 Attribute | PatSeer field | Transform |
  |---|---|---|
  | Patent ID | Publication Number | Direct |
  | Title | Title | Direct |
  | A1 | Technology Domain (fallback: CPC Main) | Direct / fallback |
  | A2 | Technology Sub-Domain (fallback: CPC All) | Direct / fallback |
  | C3 | Independent Claim Count | Numeric |
  | E1 | Simple Family Size | Numeric |
  | E2 | Legal Status | Map: granted/pending/expired |
  | E4 | Remaining Life (Years) OR computed from Estimated Expiry Date | Numeric / date-to-years |
  | E5 | Maintenance Status | Direct |
  | F1 | Simple Family Members | Extract country list |
  | H5 | Forward Citation Count | Numeric |
  | H6 | Backward Citation Count | Numeric |
  | H7 | Litigation / PTAB Flag | Map: survived/pending/none |

  Remaining 29 attributes (A3–A5, B1–B3, C1–C2, C4, D1–D3, E3, F2–F3, G1–G3, H1–H4, H8–H10, I1–I4) require manual or AI-assisted scoring.

- Freeze panes set at `AE40` — paste area and output stay in view simultaneously
- Data validation on col C (E3 Continuation): Yes/No dropdown

---

## File inventory

| File | Purpose | Status |
|---|---|---|
| `Patent_Bundling_Template_v6.xlsx` | **Active workbook.** Always use the latest version. | Current |
| `Patent_Bundling_Template_v5.xlsx` | v5 — kept for reference. | Reference |
| `Patent_Bundling_Template_v4.xlsx` | v4 — kept for reference. | Reference |
| `Patent_Bundling_Template_v3.xlsx` | v3 — kept for reference. | Reference |
| `DOCUMENTATION.md` | This document. Updated with every feature change. | Current |
| `Value_Proposition_Framework_v3.md` | 20-section framework for writing IAM Market listing copy. Covers buyer archetypes, signal tiers, canonical listing patterns, 15-step authoring workflow. | Current |
| `Patent_Buyer_Profiles_v1.docx` | Comprehensive stakeholder intelligence document. 12 buyer archetypes with full deep-dive profiles and 3 supplementary profiles. Companion to the workbook and value prop framework. | Current |
| `AI_Prompts_for_Attribute_Scoring.md` | Ready-to-use LLM prompts for scoring 20 workbook attributes across 7 prompt blocks. | Current |
| `Attribute_Derivation_Procedure.docx` | Word version of the Attribute Procedure sheet — field-by-field procedure for all 42 attributes. | Current |
| `build_v3.py` | Build script for v3 (extends v2 to v3). | Reference |
| `build_v4.py` | Build script for v4 (adds Attribute Procedure sheet). | Reference |
| `build_v5.py` | Build script for v5 (adds short-lifecycle presets, 30 patents, Portfolio Context Guide). Runs on top of v4 output. | Reference |
| `build_v6.py` | Build script for v6 (adds PatSeer Import sheet). Runs on top of v5 output. | Current |

Build scripts are kept because they encode the exact construction logic — if you ever need to extend or modify the template programmatically, edit the most recent script and rebuild. The chain is: build_v3 → build_v4 → build_v5 → build_v6.

---

## Conceptual framework

### The two-stage model

**Stage 1 — Score patents on technical attributes.** Each patent in the portfolio is scored across 42 attributes organized into 9 groups (A through I). Some attributes are categorical (e.g., stack layer), some are numeric (e.g., remaining term in years), some are 0–3 ratings (e.g., claim breadth, detectability).

**Stage 2 — Route patents into bundles using rules.** Each of the 33 bundle types has a routing rule expressed in terms of the attributes. For each (patent, bundle) pair, the rule evaluates to TRUE or FALSE. A single patent typically qualifies for multiple bundles — that's by design, because it gives flexibility in how the portfolio is packaged for different buyer types.

### Attribute groups

| Group | Theme | # of attributes | Examples |
|---|---|---|---|
| **A** | Technology Classification | 5 | Primary domain, stack layer, subsystem, use-case |
| **B** | Standards & Ecosystem | 3 | SEP potential, standard tagged, interface role |
| **C** | Claim & Scope | 4 | Claim type, breadth, count, design-around difficulty |
| **D** | Detectability & Enforcement | 3 | External detectability, teardown detectability, reads-on-products |
| **E** | Family & Lifecycle | 5 | Family size, prosecution status, continuation, term, maintenance |
| **F** | Geographic | 3 | Jurisdictions, trilateral coverage, major-market score |
| **G** | Strategic & Thematic | 3 | Convergence theme, generation, cross-industry applicability |
| **H** | Patent Quality & Vulnerability | 10 | Claim strength, prior-art exposure, litigation history, EoU availability |
| **I** | Market/Buyer Signals | 4 | Product-mapping confidence, implementation maturity, adjacent re-read |

### Why 33 bundles

The 33 bundle types fall into five conceptual layers:

1. **Technology-anchored bundles (1–8):** Group by what patents are about technically — domain, SEP, product architecture, stack layer, use-case, manufacturing, materials, algorithms.
2. **Strategic-property bundles (9–17):** Group by enforceable/strategic property — interoperability, generation, claim-type, detectability, geography, family, lifecycle, foundational+improvement, cross-industry.
3. **Theme-and-context bundles (18–21):** Group by thematic or contextual signals — convergent theme, defensive, whitespace, prosecution status.
4. **Composition-pattern bundles (22–25):** Bundles formed by deliberate composition — anchor+halo, picket-fence, strong-core+tail, continuation-live.
5. **Quality/litigation/lifecycle bundles (26–33):** Bundles defined by buyer-facing quality signals — EoU-backed, battle-tested, clean-title, high-citation, adjacent re-read, salvage, pre-expiry, provenance-coherent.

---

## Quick reference — Bundle catalog (summary)

| # | Bundle | Routing in plain English | Primary attributes |
|---|---|---|---|
| 1 | Tech Domain | Same primary domain | A1 |
| 2 | SEP | SEP potential high AND mapped to a standard | B1, B2 |
| 3 | Product Architecture | Spans subsystems of one product | A4 |
| 4 | Stack Layer | Same stack layer | A3 |
| 5 | Use-Case | Same use-case tag | A5 |
| 6 | Manufacturing / Process | Process/fab/manufacturing domain + Method claims | A1, C1 |
| 7 | Materials & Chemistry | Materials/chem domain + Apparatus/Method claims | A1, C1 |
| 8 | Algorithm / Software | App/Middleware/Cloud + Method/CRM | A3, C1 |
| 9 | Interoperability | Interface role high | B3 |
| 10 | Generational Roadmap | Same generation tag | G2 |
| 11 | Claim-Type | Same claim type | C1 |
| 12 | Detectability | Externally or teardown-detectable | D1, D2 |
| 13 | Geographic | Trilateral or shared jurisdiction set | F1, F2 |
| 14 | Family-Tree | Sufficient family size | E1 |
| 15 | Lifecycle / Term | Any patent with known remaining term | E4 |
| 16 | Foundational + Improvement | Pioneer + narrow improvements together | C2 |
| 17 | Cross-Industry | High cross-industry applicability | G3 |
| 18 | Convergent Theme | Shared convergence tag | G1 |
| 19 | Defensive / Counter-Assertion | Reads on known products | D3 |
| 20 | Whitespace / Design-Around | Hard to design around AND in same domain | A1, C4 |
| 21 | Prosecution-Status | Pending or has live continuation | E2, E3 |
| 22 | Anchor-and-Halo | Anchor (strong claim + detectable) with halo | H1, C2, D2 |
| 23 | Picket-Fence | Multiple narrow patents cluster around tech/standard | A1, B2, C4 |
| 24 | Strong-Core + Tail | Mix of strong and weak patents in same domain | H1, A1 |
| 25 | Continuation-Live | Has live continuation | E3 |
| 26 | EoU-Backed | EoU claim chart available | H9 |
| 27 | Battle-Tested | Survived validity challenge | H7 |
| 28 | Clean-Chain-of-Title | Clean ownership AND no encumbrances | H8, H10 |
| 29 | High-Citation | Forward citations above threshold | H5 |
| 30 | Adjacent-Industry Re-Read | Re-read potential AND cross-industry | I3, G3 |
| 31 | Salvage / Volume Lot | Weak claim OR near expiry OR high prior-art risk | H1, E4, H2 |
| 32 | Pre-Expiry / Last-Window | Within configurable year range | E4 |
| 33 | Provenance-Coherent | Same domain + same subsystem | A1, A4 |

---

## Sheet-by-sheet reference

The workbook contains 13 sheets in this order (v6):

### 1. README
Overview, color legend, sheet directory, usage instructions. Updated at each version with an addendum section.

### 2. PatSeer Import *(v6 — new)*
Three-zone import and staging sheet. See [PatSeer Import sheet (v6)](#patseer-import-sheet-v6) for full detail.
- **Zone 1:** Column Mapping Table (rows 10–36) — editable PatSeer field → v5 attribute mapping
- **Zone 2:** Paste Area (cols F–AD, rows 40–239) — raw PatSeer export goes here
- **Zone 3:** Mapped Output + 33 Bundle Results (cols AE–DD) — auto-computed

### 3. Configuration
Control panel for the entire analysis. See [Configuration system](#configuration-system) for full detail.
- **Active Preset dropdown (B4):** selects from 15 named presets (6 original + 3 short-lifecycle + 4 custom + 2 more custom)
- **Edit Mode toggle (B5):** Yes = Manual Override mode; No = preset mode
- **Summary line (B6):** live count of active bundles and gates

### 4. Portfolio Context Guide *(v5 — new)*
Reference sheet for short-lifecycle portfolio sales strategy. 7 sections:
- Term band overview table
- Preset usage how-to
- Critical signals by band
- Deal killer table
- IAM Market language register
- 3-tier composition strategy
- 30-patent indexed reference (SL-001 to SL-030)

### 5. Presets
Parameter × preset matrix. 15 preset columns total:
- Cols C–H: 6 original presets (All ON, NPE/Counter-Assertion, OC/FTO, Defensive Aggregator, Standards Licensee, EV Powertrain Sale)
- Cols I–L: 4 empty Custom Preset slots
- Cols M–O: 3 short-lifecycle presets (Critical 1-2yr, Monetization 3-5yr, Strategic 6-8yr) *(v5)*

### 6. Attribute Dictionary
Authoritative list of all 42 attributes (A1–I4) with group, code, name, description, scale, and example value. Read-only reference.

### 7. Attribute Procedure *(v4 — new)*
Field-by-field procedure for populating all 42 attributes from raw patent data. Four sections:
- Type legend (Direct lookup / Derived-rubric / Computed / Internal artifact)
- 10-step recommended workflow with time/effort guidance
- Master procedure table (one row per attribute: where to find / how to derive / data source)
- Free vs. paid data source reference (19 tools)
- Practical calibration notes

### 8. Patent Portfolio
Master input sheet. 42 attribute columns, one row per patent.
- **Original sample patents:** rows 3–14 (P-001 to P-012, spanning wireless / EV / AI / materials)
- **Short-lifecycle sample patents (v5):** rows 24–55 (SL-001 to SL-030), separated by color-coded band banners
- Band A (Critical 1-2yr): rows 24–33, red background
- Band B (Monetization 3-5yr): rows 35–44, amber background
- Band C (Strategic 6-8yr): rows 46–55, green background
- To add your own patents: insert rows, use dropdowns on categorical fields, copy formulas from adjacent rows into Bundle Assignment

### 9. Bundle Rules
Reference documentation for all 33 bundle types: number, name, value proposition, routing rule text, primary attributes. Read-only.

### 10. Bundle Assignment
Auto-computed TRUE/FALSE matrix. One row per patent, one column per bundle, Total column at the end.
- All cells are formulas — do not edit
- Disabled bundles: `[DISABLED]` header, grayed out, blank (not FALSE) so they don't inflate totals
- Extended to cover all 42 Patent Portfolio rows (12 original + 30 short-lifecycle)

### 11. Bundle Quality Scorecard
Auto-aggregated quality metrics per bundle. Covers all active (non-disabled) bundles.
- Strength flag (STRONG / MODERATE / WEAK) uses Configuration thresholds
- Gate columns respect individual gate toggles on Configuration

### 12. Sample Bundles
Three worked examples (5G SEP cluster, EV powertrain, edge AI). Illustrative — do not edit.

### 13. Bundle Composition Strategy
Reference documentation for 10 compositional approaches plus a pre-offering checklist. Read-only.

---

## Configuration system

### How it works

The Configuration sheet has every parameter as its own row with four columns:

| Column | Name | Purpose |
|---|---|---|
| A | Parameter | Internal key (e.g., `B2_enabled`, `SEP_B1_cutoff`) |
| B | Description | Human-readable description |
| C | **Active Value** | The value all downstream formulas read |
| D | From Preset | INDEX/MATCH lookup from the Presets sheet |
| E | Manual Override | Hand-editable; used when Edit Mode = Yes |

**Active Value formula:** `=IF($B$5="Yes", E<row>, D<row>)`

**From Preset formula:** `=INDEX(Presets!$C<row>:$O<row>, MATCH($B$4, Presets!$C$1:$O$1, 0))`  *(extended to col O in v5)*

### Two operating modes

- **Preset mode (Edit Mode = No):** Pick a preset from dropdown → all Active Values update automatically.
- **Edit Mode (Edit Mode = Yes):** Active Values read from Manual Override column → hand-edit any individual parameter.

### Preset catalog

| # | Preset name | Primary use case |
|---|---|---|
| 1 | All ON (default) | Full baseline view — all bundles active |
| 2 | NPE / Counter-Assertion | Assertion-ready: detectability, EoU, survived challenges |
| 3 | Operating Company / FTO | Subsystem coverage, interfaces, claim breadth, continuation |
| 4 | Defensive Aggregator | Volume-driven: salvage, geographic breadth, claim variety |
| 5 | Standards Licensee | SEP, interoperability, generational, geographic, battle-tested |
| 6 | EV Powertrain Sale | EV subsystems, materials, manufacturing, thermal, generational |
| 7–10 | Custom Preset 1–4 | Empty slots — overwrite with your own configuration |
| 11 | Short Lifecycle — Critical (1-2yr) | Last-window enforcement; EoU-backed; Anchor H1 ≥ 3 |
| 12 | Short Lifecycle — Monetization (3-5yr) | Primary assertion/licensing window; EoU + anchor+halo core |
| 13 | Short Lifecycle — Strategic (6-8yr) | Broadest buyer pool; OC-DEF viable; full bundle range |

### Saving a custom configuration as a preset

1. Set Edit Mode = Yes; tune Manual Override values.
2. Go to Presets sheet; find an empty Custom Preset column.
3. Enter values row-by-row (or paste-special values from Configuration col C).
4. Rename row 1 of that column.
5. Set Edit Mode = No; pick your new preset from the dropdown.

### Extending beyond 15 presets

Add columns P onward on the Presets sheet, then update the data validation source for B4 on Configuration to cover the new last column.

---

## Short-lifecycle portfolio context (v5)

### The core reframe

With ≤8 years remaining term, "time left" stops being a neutral attribute and becomes the central commercial fact. Every buyer evaluates it differently. The three presets encode this: same routing logic, different threshold windows, different bundle focus, different buyer targets.

### Term band decision tree

```
E4 ≤ 2 years  →  Critical Window    →  Preset: Short Lifecycle — Critical (1-2yr)
               →  Primary buyers: NPE-LIT, LIT-FIN
               →  Value driver: EoU chart + named defendant

E4 3–5 years  →  Monetization Window  →  Preset: Short Lifecycle — Monetization (3-5yr)
               →  Primary buyers: NPE-LIT, LIT-FIN, NPE-LIC
               →  Value driver: Enforcement window + broad claims

E4 6–8 years  →  Strategic Window   →  Preset: Short Lifecycle — Strategic (6-8yr)
               →  Primary buyers: OC-DEF, NPE-LIC, IP-FUND
               →  Value driver: Coverage depth + product mapping
```

### Critical bundle for each band

- **Critical (1-2yr):** B26 EoU-Backed + B27 Battle-Tested + B32 Pre-Expiry. Without B26, asset cannot be sold in this band.
- **Monetization (3-5yr):** B26 EoU-Backed + B22 Anchor+Halo + B32 Pre-Expiry. B24 Strong-Core+Tail allows mixing assertion anchors with a salvage tail.
- **Strategic (6-8yr):** B1 Tech Domain + B19 Defensive + B22 Anchor+Halo for OC buyers. B17 Cross-Industry + B30 Adjacent Re-Read for NPE-LIC and IP-FUND.

### 3-tier composition rule (applies to all bands)

| Tier | Qualifying criteria | Positioning | Key bundles |
|---|---|---|---|
| **A — Premium** | EoU-Backed + Survived-Challenge anchor | Price to reflect assertion readiness | B26 + B27 + B22 |
| **B — Standard** | Detectability ≥ 2, no full EoU yet | "Program-starter" lot, 60–90d claim-charting needed | B12 + B22 + B28 |
| **C — Volume/Salvage** | H1 ≤ 1 or high invalidity exposure | DEF-AGG target only; price to clear | B31 |

Never mix tiers in the same IAM Market listing.

### Universal deal killers (all bands)

- E4 < 1 year at close — case cannot be filed and resolved before expiry
- H4 ≤ 1 (divided infringement) — kills litigation route
- H10 ≠ None (FRAND/encumbrance) — restricts target set and caps royalty rates
- H8 = Clouded title — doubles due diligence, OC buyers will walk
- Pending IPR (H7 = Pending) — kills NPE-LIT and LIT-FIN interest until resolved

---

## PatSeer Import sheet (v6)

### Purpose

A low-friction onramp from PatSeer export → bundle qualification. Paste a PatSeer Excel export, get all computable bundle results immediately, with amber-highlighted cells showing exactly which attributes still need manual or AI-assisted scoring.

### How to use

1. In PatSeer: export your portfolio as Excel/CSV. Select field groups: Bibliographic + Parties + Classifications + Family + Citations + Other/Scores.
2. In the workbook: go to PatSeer Import sheet. Paste your export starting at **cell F40** (the paste area starts at column F, row 40).
3. If the paste duplicated a header row into row 40, delete row 40's data.
4. Zone 3 (cols AE onwards) auto-populates immediately — 13 attributes fill from PatSeer, 29 stay amber.
5. Score amber cells manually or use `AI_Prompts_for_Attribute_Scoring.md`.
6. Check Completeness % column (BU) — aim for ≥60% before treating results as reliable.
7. Review bundle TRUE/FALSE results (cols BW–DC). These respect the currently active Configuration preset.
8. Copy completed rows to the Patent Portfolio sheet for full Scorecard analysis.

### If your PatSeer column headers differ

Edit **column A of the Mapping Table** (rows 12–36) to match your actual export headers. All Zone 3 formulas reference the mapping table, so changing a header in col A is the only edit needed.

### PatSeer fields not in the standard export

If PatSeer doesn't include "Remaining Life (Years)" in your subscription tier, the E4 formula falls back to computing from "Estimated Expiry Date" automatically. If neither field is exported, E4 stays blank and must be entered manually.

### Extending beyond 200 patents

Add rows below row 239 in the paste area and copy the Zone 3 formula row (row 239) down to match. Update the freeze pane if needed.

---

## Buyer archetype framework

### Overview

Documented in `Patent_Buyer_Profiles_v1.docx`. The framework defines 12 buyer archetypes active in the patent secondary market, each with a full profile covering: organization types, mandate, top workbook attributes evaluated, buy signals vs. deal killers, preferred bundle types, language register, where to find them, and outreach tactics.

### Archetype reference table

| ID | Type | Primary goal | Top signal | Preferred bundles |
|---|---|---|---|---|
| OC-DEF | Operating Co. — Defensive | FTO / litigation shield | D3 + H8 Clean + E4 long | B19, B20, B28, B16, B13 |
| OC-OFF | Operating Co. — Offensive | Assert / leverage competitors | H9 Full + D1 high + H7 Survived | B26, B12, B27, B22, B23 |
| OC-EXP | Operating Co. — Expansion | Enter new markets / block rivals | G3 + G1 convergent + I3 | B17, B30, B18, B1 |
| NPE-LIC | NPE — Licensing-First | Royalty revenue via licensing | C2 broad + G3 + I3 | B17, B30, B5, B16, B25 |
| NPE-LIT | NPE — Litigation-Funded | Litigation damages / settlement | H9 Full + H7 Survived + D1 high | B26, B27, B12, B32, B22 |
| DEF-AGG | Defensive Aggregator | Shield members from NPE | D3 + H8 + volume/price | B19, B31, B17, B28, B5 |
| LIT-FIN | Litigation Finance | Return on capital from cases | H9 + E4 1-5yr + H7 Survived | B32, B26, B27, B12, B28 |
| STD-LIC | Standards Licensee | FRAND leverage or SEP blocking | B1 ≥ 2 + B2 + F2 trilateral | B2, B9, B13, B23, B10 |
| CORP-VC | Corp Dev / M&A | IP as part of acquisition | H5 + provenance + family tree | B33, B14, B29, B3 |
| IP-FUND | Patent Investment Fund | Portfolio IRR + resale optionality | H5 + I3 + E3 continuation | B29, B30, B25, B17, B10 |
| UNIV | University / Gov TTO | Commercialize research IP | H5 academic + C2 broad | B1, B16, B29, B21 |
| INSUR | IP Insurance / Risk Mgmt | Underwrite validity risk | H7 + H2 low + H8 Clean | B27, B28, B26, B33 |

### Signal tier taxonomy

Signals are classified in four tiers (defined in `Value_Proposition_Framework_v3.md`):

| Tier | Signal type | Example |
|---|---|---|
| T1 | Verifiable patent facts | H9 Full EoU, H7 Survived IPR, H8 Clean title |
| T2 | Derived from patent data | C2 broad claims, E4 long term, H5 forward citations |
| T3 | Market-context assertions | G3 cross-industry, I3 adjacent re-read |
| T4 | Positioning claims | Market size statements, buyer rationale framing |

Mixing T1 and T4 signals without discipline undermines listing credibility. The Market Context Library (MCL) in `Value_Proposition_Framework_v3.md` controls this.

### Bundle-to-buyer matching

For each bundle type, the primary and secondary buyer archetypes are documented in `Patent_Buyer_Profiles_v1.docx` (Section 7: Buyer Matching Guide). Quick reference:

- EoU-Backed (B26) + Battle-Tested (B27) → NPE-LIT, LIT-FIN (primary); OC-OFF (secondary)
- Pre-Expiry (B32) → LIT-FIN, NPE-LIT (primary); DEF-AGG (secondary)
- Defensive/Counter-Assertion (B19) → OC-DEF, DEF-AGG (primary)
- Cross-Industry (B17) + Adjacent Re-Read (B30) → NPE-LIC, IP-FUND (primary)
- Salvage/Volume Lot (B31) → DEF-AGG (primary); IP-FUND tail (secondary)
- SEP (B2) + Interoperability (B9) → STD-LIC (primary); NPE-LIC, DEF-AGG (secondary)
- High-Citation (B29) + Provenance-Coherent (B33) → IP-FUND, CORP-VC (primary)

---

## Bundle catalog — full detail

Authoritative bundle definitions. Each entry lists the bundle number, name, value proposition, routing rule, primary attributes, and any configurable thresholds.

---

### Bundle 1 — Tech Domain
- **Value proposition:** Concentrated coverage in one technical field; useful for entering or fortifying a domain position.
- **Routing rule:** A1 (primary technology domain) is populated.
- **Primary attributes:** A1
- **Tunable thresholds:** None

### Bundle 2 — SEP (Standard-Essential)
- **Value proposition:** Bundle of assets every implementer of a standard must license or design around.
- **Routing rule:** B1 ≥ `SEP_B1_cutoff` AND B2 is populated.
- **Primary attributes:** B1, B2
- **Tunable thresholds:** `SEP_B1_cutoff` (default 2)

### Bundle 3 — Product Architecture
- **Value proposition:** "Build a product" kit — full vertical coverage across subsystems of one end-product.
- **Routing rule:** A4 (subsystem) is populated.
- **Primary attributes:** A4
- **Tunable thresholds:** None

### Bundle 4 — Stack Layer
- **Value proposition:** Targets buyers operating at a specific layer of the technology stack.
- **Routing rule:** A3 (stack layer) is populated.
- **Primary attributes:** A3
- **Tunable thresholds:** None

### Bundle 5 — Use-Case
- **Value proposition:** All tools to address one customer problem, even if underlying tech is heterogeneous.
- **Routing rule:** A5 (use-case) is populated.
- **Primary attributes:** A5
- **Tunable thresholds:** None

### Bundle 6 — Manufacturing / Process
- **Value proposition:** Targets fabs, contract manufacturers, and equipment makers.
- **Routing rule:** A1 contains "process", "fab", or "manufac" AND C1 = "Method".
- **Primary attributes:** A1, C1
- **Tunable thresholds:** None

### Bundle 7 — Materials & Chemistry
- **Value proposition:** Focused IP shield around composition, synthesis, and use of a material system.
- **Routing rule:** A1 contains "material", "chem", "battery", "electrolyte", or "polymer" AND C1 ∈ {Apparatus, Method}.
- **Primary attributes:** A1, C1
- **Tunable thresholds:** None

### Bundle 8 — Algorithm / Software
- **Value proposition:** Implementation-agnostic methods deployable across many products.
- **Routing rule:** A3 ∈ {App, Middleware, Cloud} AND C1 ∈ {Method, CRM}.
- **Primary attributes:** A3, C1
- **Tunable thresholds:** None

### Bundle 9 — Interoperability
- **Value proposition:** Critical chokepoints for any player whose product must talk to others.
- **Routing rule:** B3 ≥ `Interface_B3_cutoff`.
- **Primary attributes:** B3
- **Tunable thresholds:** `Interface_B3_cutoff` (default 2)

### Bundle 10 — Generational Roadmap
- **Value proposition:** Buyer picks legacy (cheap defensive) or next-gen (forward-looking offensive).
- **Routing rule:** G2 (generation) is populated.
- **Primary attributes:** G2
- **Tunable thresholds:** None

### Bundle 11 — Claim-Type
- **Value proposition:** Different enforcement profiles; buyers pick by litigation/licensing strategy.
- **Routing rule:** C1 (claim type) is populated.
- **Primary attributes:** C1
- **Tunable thresholds:** None

### Bundle 12 — Detectability
- **Value proposition:** Litigation-ready bundle — infringement is easy to spot.
- **Routing rule:** D1 ≥ `Detect_D1_cutoff` OR D2 ≥ `Detect_D2_cutoff`.
- **Primary attributes:** D1, D2
- **Tunable thresholds:** `Detect_D1_cutoff` (default 2), `Detect_D2_cutoff` (default 2)

### Bundle 13 — Geographic
- **Value proposition:** Aligned with the buyer's market footprint and enforcement venues.
- **Routing rule:** F2 = "Yes" (trilateral coverage).
- **Primary attributes:** F1, F2
- **Tunable thresholds:** None

### Bundle 14 — Family-Tree
- **Value proposition:** Complete families together — parent + continuations + foreign counterparts.
- **Routing rule:** E1 ≥ `Family_E1_min`.
- **Primary attributes:** E1
- **Tunable thresholds:** `Family_E1_min` (default 2)

### Bundle 15 — Lifecycle / Term
- **Value proposition:** Long-life = strategic; short-life = immediate licensing.
- **Routing rule:** E4 > 0 (remaining term known).
- **Primary attributes:** E4
- **Tunable thresholds:** None

### Bundle 16 — Foundational + Improvement
- **Value proposition:** Combines blocking broad claims with defensive narrow follow-ons.
- **Routing rule:** C2 = 3 (pioneer) OR C2 ≤ 1 (improvements).
- **Primary attributes:** C2
- **Tunable thresholds:** None
- **Note:** This rule sweeps in both ends of the claim-breadth spectrum. Refine by requiring shared A1 — see Bundle Composition Strategy sheet.

### Bundle 17 — Cross-Industry
- **Value proposition:** Maximizes addressable buyer pool — many industries are potential acquirers.
- **Routing rule:** G3 ≥ `CrossIndustry_G3_cutoff`.
- **Primary attributes:** G3
- **Tunable thresholds:** `CrossIndustry_G3_cutoff` (default 2)

### Bundle 18 — Convergent Theme
- **Value proposition:** Aligns with current investment trends; commands attention from active builders.
- **Routing rule:** G1 (convergence theme) is populated.
- **Primary attributes:** G1
- **Tunable thresholds:** None

### Bundle 19 — Defensive / Counter-Assertion
- **Value proposition:** Counter-assertion ammunition — defined by who the patents read on.
- **Routing rule:** D3 ≥ `Defensive_D3_cutoff`.
- **Primary attributes:** D3
- **Tunable thresholds:** `Defensive_D3_cutoff` (default 2)

### Bundle 20 — Whitespace / Design-Around
- **Value proposition:** Closes escape routes around a known core patent or feature.
- **Routing rule:** C4 ≥ `Whitespace_C4_cutoff` AND A1 is populated.
- **Primary attributes:** A1, C4
- **Tunable thresholds:** `Whitespace_C4_cutoff` (default 2)

### Bundle 21 — Prosecution-Status
- **Value proposition:** Pending applications offer claim-tailoring flexibility for targeted-product buyers.
- **Routing rule:** E2 = "Pending" OR E3 = "Yes".
- **Primary attributes:** E2, E3
- **Tunable thresholds:** None

### Bundle 22 — Anchor-and-Halo
- **Value proposition:** One or two strong anchors fortified by narrower halo patents that close design-around routes.
- **Routing rule:** H1 ≥ `Anchor_H1_cutoff` AND C2 ≥ 1.
- **Primary attributes:** H1, C2, D2, A1
- **Tunable thresholds:** `Anchor_H1_cutoff` (default 2; raised to 3 in Critical (1-2yr) preset)

### Bundle 23 — Picket-Fence / Cluster-Around-Standard
- **Value proposition:** Multiple narrow patents collectively encircling a known commercial technology or specification.
- **Routing rule:** C4 ≥ 1 AND (A1 populated OR B2 populated).
- **Primary attributes:** A1, B2, C4
- **Tunable thresholds:** None

### Bundle 24 — Strong-Core + Quality-Diluted Tail
- **Value proposition:** A small set of high-quality assets carries the bundle; tail provides volume and continuation optionality.
- **Routing rule:** A1 is populated.
- **Primary attributes:** H1, A1
- **Tunable thresholds:** None

### Bundle 25 — Continuation-Live
- **Value proposition:** Buyer can shape future claims to match a specific target product.
- **Routing rule:** E3 = "Yes".
- **Primary attributes:** E3
- **Tunable thresholds:** None

### Bundle 26 — EoU-Backed / Litigation-Ready
- **Value proposition:** Every patent ships with an Evidence-of-Use claim chart mapped to a named product or standard.
- **Routing rule:** H9 ∈ {"Partial", "Full"}.
- **Primary attributes:** H9
- **Tunable thresholds:** None

### Bundle 27 — Survived-Challenge / Battle-Tested
- **Value proposition:** Reduces buyer's invalidity risk — assets are legally vetted.
- **Routing rule:** H7 = "Survived".
- **Primary attributes:** H7
- **Tunable thresholds:** None

### Bundle 28 — Clean-Chain-of-Title
- **Value proposition:** Transacts faster and with lower legal cost.
- **Routing rule:** H8 = "Clean" AND H10 = "None".
- **Primary attributes:** H8, H10
- **Tunable thresholds:** None

### Bundle 29 — High-Citation / Technical-Influence
- **Value proposition:** Forward citations signal technical importance and point to obvious future target companies.
- **Routing rule:** H5 ≥ `HighCitation_H5_min`.
- **Primary attributes:** H5
- **Tunable thresholds:** `HighCitation_H5_min` (default 15; lowered to 10 in short-lifecycle presets)

### Bundle 30 — Adjacent-Industry Re-Read
- **Value proposition:** Gives the buyer a "second-life" thesis — opens new buyer pools beyond the original target industry.
- **Routing rule:** I3 ≥ 2 AND G3 ≥ 2.
- **Primary attributes:** I3, G3
- **Tunable thresholds:** None *(currently using literal 2; could be made tunable)*

### Bundle 31 — Salvage / Defensive-Volume Lot
- **Value proposition:** Deliberate volume lot of weak, narrow, or near-expiry patents priced for defensive aggregators.
- **Routing rule:** H1 ≤ `Salvage_H1_max` OR E4 < `Salvage_E4_max` OR H2 ≤ `Salvage_H2_max`.
- **Primary attributes:** H1, E4, H2
- **Tunable thresholds:** `Salvage_H1_max` (default 1), `Salvage_E4_max` (default 5), `Salvage_H2_max` (default 1)

### Bundle 32 — Pre-Expiry / Last-Window
- **Value proposition:** Buyers running short-cycle licensing campaigns. Short term is a feature for litigation-finance buyers.
- **Routing rule:** `PreExpiry_min_years` ≤ E4 ≤ `PreExpiry_max_years`.
- **Primary attributes:** E4
- **Tunable thresholds:** `PreExpiry_min_years`, `PreExpiry_max_years` (defaults 1 and 4; adjusted per short-lifecycle preset)

### Bundle 33 — Provenance-Coherent
- **Value proposition:** Shared specifications and consistent terminology reduce claim-construction risk.
- **Routing rule:** A1 is populated AND A4 is populated.
- **Primary attributes:** A1, A4
- **Tunable thresholds:** None

---

## Threshold catalog

All 18 tunable thresholds available on the Configuration sheet:

| Threshold key | Default | Used by | What it controls |
|---|---|---|---|
| `SEP_B1_cutoff` | 2 | B2 | Minimum B1 (SEP potential) for SEP bundle |
| `Interface_B3_cutoff` | 2 | B9 | Minimum B3 (interface role) for Interop bundle |
| `Detect_D1_cutoff` | 2 | B12 | Minimum D1 (external detectability) |
| `Detect_D2_cutoff` | 2 | B12 | Minimum D2 (teardown detectability) |
| `Family_E1_min` | 2 | B14 | Minimum E1 (family size) |
| `CrossIndustry_G3_cutoff` | 2 | B17 | Minimum G3 (cross-industry score) |
| `Defensive_D3_cutoff` | 2 | B19 | Minimum D3 (reads on products) |
| `Whitespace_C4_cutoff` | 2 | B20 | Minimum C4 (design-around difficulty) |
| `Anchor_H1_cutoff` | 2 | B22 | Minimum H1 to qualify as anchor (raised to 3 in Critical preset) |
| `HighCitation_H5_min` | 15 | B29 | Minimum H5 forward citations (lowered to 10 in short-lifecycle presets) |
| `PreExpiry_min_years` | 1 | B32 | Lower bound of pre-expiry window |
| `PreExpiry_max_years` | 4 | B32 | Upper bound of pre-expiry window |
| `Salvage_H1_max` | 1 | B31 | Maximum H1 for salvage qualification |
| `Salvage_E4_max` | 5 | B31 | Maximum remaining term for salvage |
| `Salvage_H2_max` | 1 | B31 | Maximum H2 (prior-art exposure) for salvage |
| `Strength_depth_min` | 4 | Scorecard | Min coverage depth for STRONG strength flag |
| `Strength_detect_min` | 2 | Scorecard | Min avg detectability for STRONG flag |
| `Strength_term_min` | 10 | Scorecard | Min avg remaining term for STRONG flag (lowered per short-lifecycle preset) |

---

## Quality gates

The Bundle Quality Scorecard has 5 buyer-facing quality columns that can be individually toggled on Configuration:

| Gate key | Column | What it measures |
|---|---|---|
| `Gate_WeakestH1` | K | Minimum H1 within the bundle — reveals the weakest link |
| `Gate_InvalidityExposure` | L | % of bundle patents with H2 ≤ 1 (high prior-art risk). Above 30% triggers a quality-filter pass |
| `Gate_EoUReady` | M | % of bundle patents with H9 ≠ None. High = litigation-ready |
| `Gate_Survived` | N | % of bundle patents with H7 = "Survived" (passed a validity challenge) |
| `Gate_ContOptionality` | O | % of bundle patents with E3 = Yes (live continuation). High = future claim-shaping room |

When a gate is disabled, its entire column is grayed out, the header gets a `[DISABLED]` marker, and all cells return blanks.

---

## Bundle composition strategies

Documented on the **Bundle Composition Strategy sheet**. These sit one layer above the 33 bundle types — they describe *how to compose, position, and price* bundles once routing rules have produced candidates:

1. **Tiered structure (Premium / Standard / Lot)** — Three quality tiers from the same pool.
2. **Anchor-first bundling** — Every bundle must have a "lead patent" the buyer latches onto.
3. **Storyline bundling** — Group patents that tell a single technical story.
4. **Complementary-weakness bundling** — Pair patents whose weaknesses cancel.
5. **Modular bundling** — Pre-compose modules that buyers can mix-and-match.
6. **Carve-in vs. carve-out logic** — Explicitly mark which patents are sold only as a unit.
7. **Reserve strategy** — Keep the strongest 1–2 patents out of bundles for individual sale.
8. **Conditional / buyer-profile bundling** — Same patents, different bundle constructions for different buyers.
9. **Continuation-optionality overlay** — Bundles with live continuations marketed as "claim-tailorable".
10. **Provenance-coherent bundling** — Patents from same R&D program or sub-business for narrative coherence.

**Pre-offering checklist** (also on the Bundle Composition Strategy sheet): anchor presence, invalidity exposure < 30%, clean chain of title, coherent technical story, EoU charts available for anchors, commercially meaningful size, clear buyer profile.

---

## Workflows / how-to guides

### Workflow 1 — Analyze a new portfolio

1. Open `Patent_Bundling_Template_v6.xlsx`.
2. Go to the **PatSeer Import** sheet. Paste your PatSeer export at cell F40.
3. Review Zone 3 — amber cells require manual scoring. Use `AI_Prompts_for_Attribute_Scoring.md` for the rubric-based attributes.
4. Once Completeness % ≥ 60% for your key patents, copy rows to **Patent Portfolio** sheet.
5. Go to **Bundle Assignment** — review TRUE/FALSE qualifications per patent per bundle.
6. Go to **Bundle Quality Scorecard** — look for STRONG flags and high EoU-ready / Survived percentages.
7. Iterate: refine attribute scores, merge small bundles, split large ones.

### Workflow 2 — Configure for a specific buyer

1. Go to the **Configuration** sheet.
2. Click cell B4 (Active Preset dropdown). Pick the preset matching your buyer:
   - NPE / litigation buyer → "NPE / Counter-Assertion"
   - Operating company defensive → "Operating Company / FTO"
   - Defensive aggregator → "Defensive Aggregator"
   - Standards licensee → "Standards Licensee"
   - EV powertrain portfolio → "EV Powertrain Sale"
   - Short-lifecycle critical window → "Short Lifecycle — Critical (1-2yr)"
   - Short-lifecycle monetization → "Short Lifecycle — Monetization (3-5yr)"
   - Short-lifecycle strategic → "Short Lifecycle — Strategic (6-8yr)"
3. Check the Summary line (B6) to confirm active bundle/gate count.
4. Bundle Assignment and Scorecard recompute automatically.

### Workflow 3 — Fine-tune a preset for one analysis

1. On Configuration, set Edit Mode (B5) to **Yes**.
2. Active Values now read from the Manual Override column (E).
3. Edit any Manual Override cell — flip a toggle to No, change a threshold, etc.
4. Bundle Assignment and Scorecard recompute live.
5. Summary line confirms "EDIT MODE — using Manual Override values".
6. To return to a preset, set Edit Mode back to **No** and pick from the dropdown.

### Workflow 4 — Save a custom configuration as a preset

1. On Configuration, set Edit Mode to Yes; configure all values to your liking.
2. Go to the **Presets** sheet.
3. Pick one of the 4 empty Custom Preset columns (I–L).
4. Enter your values row-by-row matching the Configuration sheet row order.
   (Shortcut: copy column C from Configuration, paste-special "values only" into the custom preset column.)
5. Rename row 1 of that column to your preset name.
6. Back on Configuration: set Edit Mode = No; pick your new preset from the dropdown.

### Workflow 5 — Use PatSeer Import for a real portfolio

1. In PatSeer, run your search or open your project. Export to Excel.
2. In the export dialog, select all field groups available: Bibliographic, Parties, Classifications, Family, Citations, and Scores/Other.
3. Open `Patent_Bundling_Template_v6.xlsx` → go to **PatSeer Import** sheet.
4. Check that row 39 headers match your export column names. If they differ, edit col A of the Mapping Table (rows 12–36).
5. Paste your PatSeer export at cell **F40** (paste area starts col F, row 40).
6. Review amber cells in Zone 3 (cols AE–BT). For each amber attribute:
   - Use `AI_Prompts_for_Attribute_Scoring.md` for A3/A5, B1, C2, D1–D3, H1, H4, I2–I4
   - Enter directly for E3 (col C), E2 override (col D), E5 override (col E) in the manual helper columns
7. Check Completeness % column (BU). Traffic-light: green ≥60%, amber 30–60%, red <30%.
8. Review bundle results (cols BW–DC). The active Configuration preset determines which bundles are evaluated.
9. Copy rows with sufficient completeness to **Patent Portfolio** for full Scorecard analysis.

### Workflow 6 — Write an IAM Market listing from a completed bundle

1. Identify the bundle's primary buyer archetype using `Patent_Buyer_Profiles_v1.docx` (Section 7).
2. Open `Value_Proposition_Framework_v3.md` and apply the archetype's language register settings.
3. Follow the 15-step authoring workflow in that document.
4. Use the signal tier taxonomy (T1–T4) to sequence claims: lead with T1, support with T2, contextualize with T3, frame with T4.
5. Apply the deal-killer pre-emption logic for the archetype — address the buyer's top objections proactively in the listing.

### Workflow 7 — Add a new patent attribute

This requires editing the build scripts and regenerating the workbook:

1. Edit `build_v6.py` (or the most recent build script).
2. Add the new attribute to the relevant group's headers in the attribute list.
3. Update any downstream routing rules in Bundle Assignment that should reference the new attribute.
4. Add corresponding entries to `AI_Prompts_for_Attribute_Scoring.md` if the attribute requires LLM scoring.
5. Update the PatSeer Import sheet mapping table if the attribute can be derived from PatSeer data.
6. Run the build script and `scripts/recalc.py` to verify zero formula errors.
7. Update this documentation.

### Workflow 8 — Add a new bundle type

1. Edit `build_v6.py` — append to the `BUNDLES` list and add routing logic.
2. Add the bundle to the `BA_FORMULAS_ROW3` dict in `build_v6.py` for the PatSeer Import sheet.
3. Update the `BUNDLE_NAMES` dict in `build_v6.py`.
4. Update all presets to mark the new bundle on/off.
5. Rebuild the workbook; run recalc.
6. Update this documentation and `Patent_Buyer_Profiles_v1.docx` buyer-matching table.

---

## Maintenance notes

### Formula conventions

- **Blue text** = user input cells (hardcoded values)
- **Black text** = formula cells / computed values
- **Green text** = cross-sheet links (e.g., Bundle Assignment pulling from Patent Portfolio)
- **Teal background** = PatSeer Import Zone 3 cells auto-filled from PatSeer data
- **Amber background** = cells requiring manual or AI-assisted input
- **Gray italic text with `[DISABLED]`** = parameter or row turned off via Configuration

### Conditional formatting layers (Bundle Assignment and PatSeer Import)

1. **Disabled gray-out** (highest priority): triggered when header contains `[DISABLED]`
2. **TRUE = green / FALSE = light red**: applied to enabled bundle cells
3. **Completeness traffic-light** (PatSeer Import only): green ≥60%, amber 30–60%, red <30%

### Formula counts per version

| Version | Total formulas | Notes |
|---|---|---|
| v5 | 2,413 | 12 original + 30 new patents, 3 new presets |
| v6 | 12,013 | +9,600 from PatSeer Import sheet (200 rows × ~48 formulas per row) |

### Known limitations

- **Bundle 16 (Foundational + Improvement)** sweeps both ends of C2; refine by adding shared A1 filter if needed.
- **Bundle 30 (Adjacent Re-Read)** uses literal thresholds (I3 ≥ 2 AND G3 ≥ 2) — not Configuration-tunable yet.
- **Picket-Fence (Bundle 23)** doesn't enforce minimum bundle size of 4; aggregate count observed on Scorecard only.
- **E3 (Continuation)** is binary Yes/No; real portfolios may need a 3-way split (live / abandoned / none).
- **No carve-in/carve-out marker** on Bundle Assignment yet — strategy documented but not encoded.
- **PatSeer Import** derives E2 (Prosecution) from Legal Status text matching — works for standard PatSeer output but may need adjustment for non-standard legal status strings.
- **PatSeer Import E4 formula** falls back from Remaining Life → Estimated Expiry Date → blank. If neither field is exported by PatSeer, E4 must be entered manually.
- **PatSeer Import bundle formulas** reference Configuration sheet thresholds — they respect the active preset but do NOT update the main Scorecard. Copy rows to Patent Portfolio for Scorecard analysis.

### When formulas don't recalculate

- In Excel: press F9 or Ctrl+Alt+F9 for a full recalc.
- In LibreOffice: Tools → Cell Contents → Recalculate Hard (Ctrl+Shift+F9).
- After Python/openpyxl edits: run `python scripts/recalc.py <file>` to force LibreOffice to compute cached values.

### Rebuilding the workbook from source

The build chain is sequential — each script adds on top of the previous version's output:

```
build_v3.py   →  Patent_Bundling_Template_v3.xlsx
build_v4.py   →  Patent_Bundling_Template_v4.xlsx  (input: v3 output)
build_v5.py   →  Patent_Bundling_Template_v5.xlsx  (input: v4 output)
build_v6.py   →  Patent_Bundling_Template_v6.xlsx  (input: v5 output)
```

After any rebuild, always run `python scripts/recalc.py <output_file> 120` and verify `"status": "success"` with zero errors before distributing.

---

## Glossary

| Term | Definition |
|---|---|
| **Anchor patent** | A high-quality patent (typically H1=3, C2≥2, D2≥2) that serves as the lead asset in a bundle. |
| **Battle-tested** | A patent that has survived a validity challenge (IPR, EPO opposition, re-examination). |
| **Buyer archetype** | One of 12 named stakeholder types active in the patent secondary market, each with a distinct mandate, attribute weighting, and language register. Defined in `Patent_Buyer_Profiles_v1.docx`. |
| **Carve-in** | Patents that must be sold as a unit and cannot be cherry-picked out of a bundle. |
| **Carve-out** | Patents excluded from a bundle — either held for individual sale or encumbered. |
| **Chain of title** | The sequence of ownership transfers from the original inventor(s) to the current owner. A clean chain has all assignments recorded and no gaps. |
| **Claim chart** | A document mapping each element of a patent claim to a corresponding feature in a product or standard. The legal heart of an EoU. |
| **Completeness %** | In the PatSeer Import sheet: number of non-blank attribute cells / 42. A proxy for how ready the record is for Scorecard analysis. |
| **Continuation** | A patent application filed off a parent, claiming priority to it. Keeps the family "alive" for further claim-shaping. |
| **Deal killer** | An attribute value that prevents a sale to a given buyer archetype regardless of all other attributes. |
| **Defensive aggregator** | A buyer (e.g., RPX, OIN) that acquires patents to reduce litigation risk for members. Values volume and broad coverage. |
| **Divided infringement** | When the acts required to practice a claim are split between multiple parties. Makes the patent harder to enforce. |
| **EoU (Evidence of Use)** | A document mapping a patent's claims to a specific commercial product, demonstrating infringement. Significantly increases bundle value. |
| **Estoppel (file wrapper)** | Limitations on claim scope from arguments the patentee made during prosecution. High estoppel reduces enforceability. |
| **FTO (Freedom to Operate)** | The ability to commercialize a product without infringing third-party patents. |
| **Halo patent** | A patent supporting an anchor — typically narrower, closing design-around routes around the anchor. |
| **IAM Market** | Intellectual Asset Management Market — the primary online platform for listing patents for outright sale. The primary commercial channel targeted by this framework. |
| **IPR (Inter Partes Review)** | A US PTAB proceeding to challenge patent validity. Surviving an IPR is a strong validity signal. |
| **Language register** | The vocabulary, framing, and emphasis style appropriate for a specific buyer archetype. Defined per-archetype in `Patent_Buyer_Profiles_v1.docx`. |
| **Maintenance fees** | Periodic fees required to keep a granted patent in force. Missing them causes lapse. |
| **MCL (Market Context Library)** | A curated set of pre-approved market-context statements (T3/T4 signals) defined in `Value_Proposition_Framework_v3.md`. Controls signal mixing discipline. |
| **NPE (Non-Practicing Entity)** | An entity that owns patents but doesn't make products. Often a target buyer for assertion-ready bundles. |
| **PatSeer** | Patent analytics platform used as the primary data export source for the PatSeer Import sheet in v6. |
| **Picket-fence** | A bundle of multiple narrow patents collectively encircling a technology or standard. |
| **Pioneer patent** | A foundational, broad-scope patent opening a new technology area. High value but also high invalidity risk. |
| **PTAB** | Patent Trial and Appeal Board — the US tribunal handling IPRs and other post-grant proceedings. |
| **Remaining term** | Years remaining before a patent expires. The central commercial variable for short-lifecycle portfolios. In v5, this drives the 3-band classification (Critical / Monetization / Strategic). |
| **Salvage lot** | A volume bundle of weak, near-expiry, or otherwise low-individual-value patents priced and marketed honestly as such. |
| **SEP (Standard-Essential Patent)** | A patent that must be infringed to comply with a technical standard (5G NR, Wi-Fi 7, USB-C, etc.). |
| **Signal tier** | T1 = verifiable patent facts; T2 = derived from data; T3 = market-context assertions; T4 = positioning claims. Defined in `Value_Proposition_Framework_v3.md`. |
| **Term band** | One of three lifecycle windows in v5: Critical (1-2yr), Monetization (3-5yr), Strategic (6-8yr). Drives buyer targeting, bundle selection, and preset configuration. |
| **Trilateral coverage** | A patent granted in the major three jurisdictions: US + EP + at least one of CN/JP/KR. |
| **Whitespace** | The technical area surrounding a known core patent where competitors might design around. A whitespace bundle closes those escape routes. |

---

*Document maintained alongside `Patent_Bundling_Template_v6.xlsx`. Last updated for **v6 — PatSeer Import sheet release**.*
