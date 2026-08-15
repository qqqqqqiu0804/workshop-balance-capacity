---
name: workshop-balance-capacity
description: |-
  This skill should be used when a user wants to analyze a production workshop's
  "balance point capacity" (平衡点产能 / 平衡点产量 / 产能基准 / 盈亏平衡产量 / equilibrium monthly
  output / break-even output) from order-detail spreadsheet data, typically to design a
  supervisor's salary commission baseline where output above the balance point earns a tiered
  commission. It generates a boss-facing, single-file HTML report with multiple candidate
  balance-point values, sensitivity analysis, and a dynamic rolling-baseline mechanism.
  Trigger phrases (Chinese or English) include: "做一份车间平衡点产能分析", "生产订单数据算平衡点产能",
  "按产量设计提成基准", "平衡点报告", "产能基准报告", "车间平衡点", "提成基准", "算产能基准",
  "平衡点怎么定", "balance point report", "equilibrium capacity", "workshop capacity baseline",
  "commission baseline". Also trigger when the user references a production-order template (订单明细)
  and asks for a capacity or commission-baseline number without specifying the method, or asks to
  reuse the balance-point workflow for a different workshop or a different template.
agent_created: true
---

# Workshop Balance-Point Capacity (平衡点产能) Analysis

## Overview

Turn a production `订单明细` (order detail) spreadsheet into a decision-ready HTML report that
establishes a workshop's **monthly balance-point capacity `B`** — the output level above which a
supervisor earns commission. The skill deliberately outputs a **candidate range + sensitivity
table**, never a single fixed number: choosing `B` is a management decision that trades off
commission cost against motivation, so the agent must present options, not decide.

## When To Use

- User has monthly production order data (xlsx) and needs a capacity baseline for commission design.
- User asks for "平衡点产能 / 盈亏平衡产量 / 产能基准 / 提成基准" analysis.
- Report must be readable by non-technical management (boss-facing styling).

## Workflow

### Step 1 — Inspect the source template (read-only)
Load the xlsx with `openpyxl` using `data_only=True` (reads computed values, not formulas).
Identify the order-detail sheet and confirm column positions for **order date** and **quantity**.
Reference `references/data_schema.md` for the expected template shape and known data-quality traps.
Report the sheet list, headers, row count, and month coverage to the user before computing.

### Step 2 — Extract monthly output (real-time, no cache)
Iterate order-detail rows, coerce the date column to `YYYY-MM`, and accumulate quantity per month.
Also track the **set of distinct days with records** per month — this denominator drives the
workday-normalization method. No intermediate file is written; everything is computed in memory.

### Step 3 — Apply the complete-month caliber
Hard-code the set of months usable for modeling (typically exclude months with partial data due to
Spring Festival, month-start cutoffs, etc.). Surface the excluded months and each month's record-day
count in the report so the caliber is auditable. Make this a top-of-script config constant.

### Step 4 — Compute two algorithm families (core method)
For the complete months' monthly output series `vals`:

**Family A — Historical raw method** (actual achieved output)
- `mean`, `median`, `mean*0.95` (conservative), `max*0.80` (industry "standard quota ≈ peak×80%").

**Family B — Workday-normalized method** (purer capacity, removes "fewer working days" noise)
- `daily = [v/d for v,d in zip(vals, record_days)]`
- `norm = [d * STD_DAYS for d in daily]`  (STD_DAYS default 26)
- then take the same 4 statistics on `norm`.
This matters because low-output months are often low due to **fewer recorded days**, not lower
capacity; normalizing prevents that from dragging `B` down.

Present all 8 candidates. Recommend **raw median** for "actual achievement" framing and
**normalized median** for "pure capacity" framing; the gap between them is the candidate interval.

### Step 5 — Sensitivity (baseline tightness)
For each candidate `B`:
- `over_months = count(v exceeds B for v in vals)`
- `over_total = sum(max(0, v-B) for v in vals)`
Lower `B` → looser (more months trigger commission). This table lets management invert the problem:
"how much commission do we want to pay?" → pick the `B` that yields it.

### Step 6 — Dynamic rolling mechanism (don't hard-freeze B)
Describe a rolling baseline: first month uses a fixed initial `B`; from month 2, `B = median of the
prior 5 complete months`, auto-drifting with capacity trend. Add: management manual override on
structural change (new equipment/process), anomaly-month exclusion from the window, and a lower-bound
guard so consecutive low months can't collapse `B`.

### Step 7 — Render the boss-facing HTML report
Use `scripts/gen_report.py` as the engine. It builds a **single self-contained HTML file** (inline CSS,
inline SVG bar chart, KPI cards, method cards, timeline) so the file opens by double-click with no
external dependencies. Dark consulting-style theme. The script is the source of truth — editing the
config block and re-running regenerates the report; do not hand-edit the HTML.

### Step 8 — Data-quality caveats (required section)
Always list genuine data risks found (e.g. order-number column empty, order-status barely filled,
anomaly-low month). **Do NOT** list cosmetic template-internal display bugs (wrong dashboard formula,
zero-counted helper cells) unless the user explicitly wants them — they distract a boss-facing report
and were explicitly removed per user preference. See `references/data_schema.md`.

## Critical Rules (learned the hard way)

1. **Never decide `B` for the user.** The user corrected this once: "establish a value" means
   *produce the analysis*, not *pick the number*. Always output candidates + sensitivity; let
   management choose. State this explicitly in the report's executive summary.
2. **Workday-normalize before comparing months.** Raw monthly medians mix "low capacity" with
   "few working days". Skipping normalization biases `B` downward.
3. **Incomplete months must be excluded from modeling**, but shown as excluded for transparency.
4. **Don't over-engineer at design stage.** If the surrounding task uses a staged AI-coding
   template (problem-definition → open-source check → plan → skeleton → MVP → security → audit →
   wrap-up), keep Stage 0–3 text-only and write the generator script only at the MVP stage.

## Customization Points (scripts/gen_report.py)

All knobs are a top-of-script CONFIG block; the logic below stays untouched:
- `src` / `out` — override via CLI: `python gen_report.py 源.xlsx 输出.html`
- `sheet`, `date_col`, `qty_col` — map to the user's template
- `all_months`, `complete_months` —建模口径
- `std_days` — standard working days for normalization (default 26)
- `candidates` list — add/replace algorithms; each feeds the candidate + sensitivity tables automatically

## Resources

### scripts/
- `gen_report.py` — the report engine. Reads the xlsx, computes both algorithm families, renders the
  single-file HTML. Reusable as-is after setting CONFIG or passing CLI args.

### references/
- `data_schema.md` — expected `订单明细` template shape, column mapping, and the known data-quality
  traps to check (and which to omit from a boss-facing report).
