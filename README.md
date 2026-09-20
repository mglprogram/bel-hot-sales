# bel-hot-sales

An [Agent Skill](https://docs.anthropic.com/en/docs/agents-and-tools/agent-skills) that turns a
`十部半托管-库存管理表` workbook into a **two-week hot-selling product table** in a fixed,
established layout ("V3").

Give the agent the workbook and a request like *"筛选出两周都有销量的产品，按 V3 格式输出"* and it
produces an `.xlsx` with the recent week split into 7 daily columns, the prior week summarized,
and a week-over-week 环比 column.

## What it does

Starting from the `总销售表` sheet of the workbook:

1. **Filters** by brand — matches the **MSKU brand segment** (e.g. `TM-BEL-102503-5SetA-S` → `BEL`).
2. **Aggregates** by 款号 (style), not by MSKU.
3. **Detects the date window** automatically — anchors on the last date column that actually
   contains data, then takes the 7 days ending there as the recent week and the preceding 7 as
   the prior week.
4. **Keeps** styles that sold in **both** weeks; single-week-only styles are listed separately
   with the reason (newly ramping vs. sold out).
5. **Writes** the V3 layout: 排名 / 款号 / 上上周合计 / 7 daily columns / 上周合计 / 7天日均 / 环比.

## Install

Copy the skill folder into your skills directory:

```bash
cp -r bel-hot-sales ~/.agents/skills/
```

Or install from the packaged `bel-hot-sales.skill` if you use a client that consumes skill bundles.

Requires Python 3 with `openpyxl`.

## Usage

The bundled script can be used standalone:

```bash
python scripts/bel_pipeline.py "<十部半托管-库存管理表.xlsx>" [output.xlsx] [--brand BEL]
```

Useful flags:

| Flag | Purpose |
|---|---|
| `--brand BEL` | Brand segment to filter on (default `BEL`). Any brand code works, e.g. `BH`, `TH`, `BLA`. |
| `--last-week 9.12-9.18` | Pin the recent week explicitly instead of auto-detecting. |
| `--prev-week 9.05-9.11` | Pin the prior week explicitly. |

The script prints the detected window, row counts, and a self-check, so you can confirm the
window matches what you meant before trusting the output:

```
D0(最后有数据日) = 9.18
上周窗口   = 9.12 .. 9.18
上上周窗口 = 9.05 .. 9.11
BEL 行数=1177  款号=114  保留=35  排除=11
校验: 日列之和 vs 上周合计 不符行数 = 0
```

## Two data traps this skill exists to avoid

Both were hit for real while developing this, and both fail *silently* — you get a
plausible-looking table with wrong numbers.

**1. The MSKU is not prefixed by the brand.**
The format is `platform-brand-style-color-size`. The brand sits in the **second** segment.
Filtering with `startswith("BEL")` returns **zero rows**, not an error.

**2. Date headers repeat across multiple blocks.**
The header contains *two* complete `9.5日 … 9.18日` label sets (around columns 128–141 and
860–873), plus a duplicated `4.24日`. The early block is a sparse legacy artifact; the late
block is live. Because labels repeat, **any lookup keyed on `(month, day)` silently resolves to
the wrong block** — undershooting sales by roughly an order of magnitude.

This was verified empirically: summing columns 860–873 matches the sheet's own `最近14天日均`
for **1709/1709** rows, while columns 128–141 match **0/1709**. The pipeline therefore resolves
columns *by position* and warns when duplicates are present.

## Output structure

Two sheets:

- **两周均有销量** — 13 columns:
  `排名 | 款号 | 上上周销量合计 | 9.12日 … 9.18日 | 上周销量合计 | 最近7天日均 | 上周环比上上周`
  Orange banner over the daily columns, top-3 rows highlighted, 环比 shown as a signed
  percentage (red up / blue down), panes frozen at `C5`.
- **已排除-仅单周有销量** — the single-week-only styles and why they were excluded.

## A note on interpreting the result

The keep-condition is "sold in both weeks", which is deliberately loose. Two consequences
worth knowing before you act on the numbers:

- **Small-base styles distort the 环比 column.** A style going 2 → 22 units shows `+1000%`.
  Strictly correct, but weak evidence of a trend. If you want *stable sellers* rather than
  *touched both weeks*, add a volume threshold.
- **Notable single-week styles are excluded by design.** A style that sold 10 units last week
  but zero the week before is filtered out as a new ramp-up — which is why the skill surfaces
  the excluded list rather than burying it.

## Repository layout

```
bel-hot-sales/
├── SKILL.md                 # skill definition: triggers, rules, traps
├── scripts/
│   └── bel_pipeline.py      # the pipeline
└── evals/
    └── evals.json           # test prompts + assertions
```

The eval fixture workbook (~6 MB) is not committed; see `evals/evals.json` for the expected
input shape.
