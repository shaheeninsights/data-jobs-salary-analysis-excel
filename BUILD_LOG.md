# Project 2 Rebuild — Stage 1: Power Query ETL
## Detailed Step Log

**Workbook:** built and worked on under the working name `1_Project_Analysis_Rebuild.xlsx` throughout this log (later renamed to `data_jobs_salary_analysis.xlsx` for final publication — see note below).
**Source data:** `data_jobs_salary_all.xlsx` (32,672 rows, 16 columns)

> **Note on filenames in this log:** the workbook went through a few working names during development (`1_Project_Analysis_Rebuild.xlsx`, and briefly `2 Luke_Project Updated.xlsx` after a mid-project rename and an autosave recovery — see the "Excel Freeze & Recovery" section). The final published version is `data_jobs_salary_analysis.xlsx`. References to the older names below are left as-is to keep this an accurate record of what actually happened at each step.
**Goal of Stage 1:** Produce two clean queries, both loaded to the Data Model:
- `data_jobs_all` — one row per job posting
- `data_job_skills` — one row per (job, skill) pair

---

## Part A — Extract & Duplicate

1. New blank workbook created (later saved as `1_Project_Analysis_Rebuild.xlsx`).
2. **Data** tab → **Get Data** → **From File** → **From Workbook** → selected `data_jobs_salary_all.xlsx`.
3. In Navigator, selected `Sheet1` → clicked **Transform Data** (opens Power Query Editor, does not load yet).
4. Renamed the query (right-click query name in left pane → Rename) to `data_jobs_all`.
5. Right-click `data_jobs_all` → **Duplicate** → renamed the duplicate to `data_job_skills`.

At this point both queries are identical and row-aligned (same 32,672 rows, same order) — this matters later for `job_id`.

---

## Part B — Clean `data_jobs_all`

Query: `data_jobs_all`

1. Selected `job_skills` column header → **Home** tab → **Manage Columns** group → **Remove Columns** (right-click on header also works if available: right-click → Remove).
2. Fixed data types by clicking the type icon in each column header:
   - `job_posted_date` → **Date/Time**
   - `salary_year_avg` → **Decimal Number**
   - `job_work_from_home`, `job_no_degree_mention`, `job_health_insurance` → **True/False**
3. Selected `job_title_short`, `job_title`, `company_name` (ctrl-click to multi-select) → **Transform** tab → **Format** → **Trim**.
4. Verified: no red error cells, no warning triangles in any column.

---

## Part C — Add `job_id` to both queries (Index Column)

**Why:** Neither query had a native `job_id` column. Since both queries are still row-aligned copies of the same source at this point, an Index column added *now* (before any row-count-changing transformation) works as a valid `job_id` key for joining the two tables later.

**Key rule established:** the index must be added to `data_job_skills` **before** the `job_skills` list is split into rows — otherwise the index would number the already-exploded rows instead of the original jobs, breaking the join.

### Standard method (used first, then replaced — see below)
- **Add Column** tab → **Index Column** → **From 1**
- Rename column: double-click header (or right-click → Rename) → type `job_id`

### Refined method — combined into a single Applied Step (used going forward)
Instead of Add Index → Rename → Reorder as three separate steps, this was collapsed into one step via the formula bar:

1. Clicked the **Added Index** step in Applied Steps.
2. Clicked into the formula bar.
3. Replaced the formula with:
   ```
   = Table.ReorderColumns(Table.AddIndexColumn(#"Previous Step Name", "job_id", 1, 1, Int64.Type), {"job_id"} & List.RemoveItems(Table.ColumnNames(#"Previous Step Name"), {"job_id"}))
   ```
   (`#"Previous Step Name"` replaced with whatever the actual prior step was called.)
4. Pressed Enter.
5. Renamed this Applied Step to **`Added job_id`** (right-click step → Rename).

Applied to **both** `data_jobs_all` and `data_job_skills`.

**Result:** `job_id` appears as the first column on both queries, numbered sequentially from 1.

---

## Part D — Explode `job_skills` in `data_job_skills`

Query: `data_job_skills`
Raw values in `job_skills` looked like: `['sql', 'python', 'aws', 'pyspark']`

### Step D1 — Clean punctuation
Selected `job_skills` column → **Transform** tab → **Replace Values**, run three times:
- Find `[` → Replace with *(blank)*
- Find `]` → Replace with *(blank)*
- Find `'` → Replace with *(blank)*

*(Optional consolidation into one step, using nested `Table.ReplaceValue` in the formula bar — not required, three-step version works identically. Left as three separate steps in the final build.)*

**⚠️ Issue hit & fixed:** Initially the **Split Column** step (Part D2 below) was accidentally positioned/run *before* this cleaning step in the Applied Steps order. This caused leftover `]` and `'` characters to remain on the *last* skill in each list (e.g. `git']`, `tensorflow']`, `hadoop']`), because the split ran on still-dirty text and the cleanup no longer applied afterward.
**Fix:** Re-ordered so punctuation cleanup happens *before* the split. Re-verified clean output afterward.

### Step D2 — Split into rows
Selected `job_skills` column → **Transform** tab → **Split Column** → **By Delimiter**:
- Delimiter: **Comma**
- Split at: **Each occurrence of the delimiter**
- Advanced options → **Split into: Rows** (critical — NOT Columns)
- Quote Character: left as default
- Clicked **OK**

**Result:** each job's skill list exploded into multiple rows, one skill per row, with `job_id` and all other original columns repeated across each new row belonging to the same job.

### Step D3 — Trim final skill text
Selected `job_skills` column → **Transform** tab → **Format** → **Trim** (removes leftover leading spaces from the split, e.g. `" python"` → `"python"`).

**Verified:**
- `job_skills` values are clean lowercase text with no brackets/quotes/spaces (e.g. `sql`, `python`, `aws`).
- `job_id` correctly repeats across the multiple rows generated from the same original job.

---

## Part E — Close & Load to Data Model

**Intended method:**
1. **Home** tab → dropdown arrow under **Close & Load** → **Close & Load To...**
2. Select **Only Create Connection** + check **Add this data to the Data Model** → OK.

**⚠️ Issue hit & fixed:** The dropdown didn't register as expected — clicking **Close & Load** loaded both queries straight to worksheets instead of prompting the load-destination dialog.
**Fix — corrected after the fact without redoing transformations:**
1. In the main Excel window, **Queries & Connections** pane (right-hand side) → right-click `data_jobs_all` → **Load To...**
2. Chose **Only Create Connection** + checked **Add this data to the Data Model** → OK.
3. Confirmed the "possible data loss" warning (safe to confirm — no PivotTables/formulas had been built on the worksheet versions yet).
4. Repeated for `data_job_skills`.

**Result:** Both worksheet tabs greyed out/emptied (expected behavior for connection-only queries) — data now lives only in the Data Model.

---

## Part F — Save

**⚠️ Issue hit:** First **Ctrl+S** of the session triggered a **Save As** dialog (expected — file had never been named/saved before this point). Dialog was cancelled once by mistake, meaning work briefly existed only in memory.
**Fix:** Re-ran Ctrl+S, completed Save As properly this time — named and saved as `1_Project_Analysis_Rebuild.xlsx`.

**Lesson adopted:** Save (Ctrl+S) after each meaningful step/checkpoint from now on, and consider enabling AutoSave if the file is stored in OneDrive/SharePoint, to avoid losing unsaved work if the laptop closes/sleeps.

---

## Final State — Confirmed via Power Pivot → Manage

Opened **Power Pivot** tab → **Manage**. Confirmed both tables present in the Data Model:
- `data_jobs_all`
- `data_job_skills`

**Stage 1 (Power Query ETL) — Complete.**

---

## Key Lessons From This Stage

1. **Duplicate the source query before any transformation** — enables two independent analytical paths (job-level vs. skill-level) from one raw dataset.
2. **When no natural key exists, an Index column can serve as `job_id`** — but only safely if added while both queries are still row-aligned, and *before* any row-multiplying transformation (like splitting a list into rows).
3. **Order of Applied Steps matters** — a cleaning step must run before a step that depends on clean input (punctuation cleanup before delimiter split).
4. **"Split into Rows" vs "Split into Columns"** — Rows is correct whenever the number of items per record varies (like skills-per-job); Columns is only appropriate for a fixed, known number of parts per value.
5. **Close & Load To... → Only Create Connection + Add to Data Model** is the correct load target for tables destined for Power Pivot/DAX work — avoids cluttering the workbook with raw worksheet data.
6. **Save early, save often** — Applied Steps and query definitions are only persisted once the workbook itself is saved.

---

## Next: Stage 2 — Power Pivot

Building the relationship between `data_jobs_all` and `data_job_skills` via `job_id` in the Data Model.

---
---

# Stage 2 — Power Pivot: Building the Relationship

**Goal:** Connect `data_jobs_all` and `data_job_skills` in the Data Model via `job_id`, so job-level and skill-level data can be analyzed together.

## Part A — Open Diagram View

1. **Power Pivot** tab → **Manage** (opens the Power Pivot window, showing the Data Model).
2. Confirmed both `data_jobs_all` and `data_job_skills` present as tabs.
3. Inside Power Pivot's own **Home** tab → **View** group → **Diagram View** (switches from grid view to a visual table-relationship layout).

## Part B — Create the relationship

1. In Diagram View, clicked and dragged from `job_id` in `data_jobs_all` → to `job_id` in `data_job_skills`.
2. Released to draw the connecting line.

**Verified relationship type** by reading the line's endpoints directly (no need to click/hover):
- `1` on the `data_jobs_all` side — confirms `job_id` is unique per row (one row per job).
- `*` on the `data_job_skills` side — confirms multiple rows share the same `job_id` (many skill-rows per job).
- Arrow pointing toward `data_job_skills` — correct filter direction (filtering jobs flows down to filter their skills).

Result: correctly formed **One-to-Many relationship**, `data_jobs_all` (1) → `data_job_skills` (*).

## Part C — Sanity-check with a PivotTable

1. **Insert** tab → **PivotTable** → **Use this workbook's Data Model** → **New Worksheet** → OK.
2. In PivotTable Fields pane: expanded `data_jobs_all` → dragged `job_title_short` into **Rows**.
3. Expanded `data_job_skills` → dragged `job_skills` into **Rows**, nested below `job_title_short`.
4. Dragged `job_id` into **Values**.

**⚠️ Issue hit & fixed:** `job_id` in Values defaulted to **Sum**, producing meaningless large numbers (e.g. 19,299,833) instead of row counts.
**Fix:** Clicked the `job_id` field dropdown in the Values area → **Value Field Settings...** → changed "Summarize value field by" from **Sum** to **Count** → OK. Produced sensible counts (e.g. 86 for Cloud Engineer + a given skill, 1001 for Business Analyst + a given skill — plausible given relative frequency of those job titles in the dataset).

**Note on the yellow "Relationships between tables may be needed" banner:** This appears automatically any time a PivotTable pulls fields from multiple Data Model tables — it does not mean the relationship is missing or broken. Since the PivotTable already produced correct, sensible counts across both tables, the relationship was confirmed working. The banner was dismissed manually via the **X** on the banner; it can otherwise be ignored.

**Stage 2 (Power Pivot relationship) — Complete.**

## Key Lessons From This Stage

1. Diagram View's `1` / `*` notation on the relationship line tells you the cardinality directly — no need to click or hover for a tooltip.
2. Default aggregation for a numeric-looking ID field in a PivotTable's Values area is **Sum**, which is meaningless for an ID column — always check and switch to **Count** (or Distinct Count) when the field is a key, not a measure.
3. The "Relationships between tables may be needed" banner is a generic prompt tied to using multiple tables in one PivotTable — not an error indicator. Trust the actual data output (sensible counts) over the banner's presence.

---

## Next: Stage 3 — DAX Measures

Building DAX measures to answer the four README analysis questions (skills vs. pay, regional salary, top skills, pay by top 10 skills).

---
---

# Stage 3 — DAX Measures (in progress)

**Goal:** Build reusable DAX measures in the Data Model to answer the four README questions.

## Part A — Median Salary (overall)

Table: `data_jobs_all`

1. **Power Pivot** tab → **Manage** → clicked `data_jobs_all` table tab at the bottom.
2. Clicked into an empty cell in the **Calculation Area** (blank row below the data grid).
3. Typed:
   ```
   Median Salary := MEDIAN(data_jobs_all[salary_year_avg])
   ```
4. Pressed Enter.

**⚠️ Issue hit & fixed:** The measure cell appeared to show only truncated text (`Median Salary:...`) with no visible number — looked like the result wasn't calculating. This was simply a **column width issue** — the calculation area cell was too narrow to display the full label + value.
**Fix:** Double-clicked the column border in the calculation area to auto-fit the width. Full formula and result then displayed correctly. (Note: measure cells in the calculation area are greyed out by design — that's normal, distinguishing a measure from raw column data.)

**Result:** `Median Salary` = **115,000**

## Part B — Median Salary, United States only

Table: `data_jobs_all`

1. Clicked into another empty cell in the calculation area.
2. Typed:
   ```
   Median Salary US := CALCULATE(MEDIAN(data_jobs_all[salary_year_avg]), data_jobs_all[job_country] = "United States")
   ```
3. Pressed Enter.

**Result:** `Median Salary US` = **118,940** (sensibly higher than the overall median, consistent with README's US salary premium insight).

## Part C — Skill Count (total skill mentions)

Table: `data_job_skills`

1. Clicked the `data_job_skills` table tab.
2. Clicked into an empty calculation area cell.
3. Typed:
   ```
   Skill Count := COUNTROWS(data_job_skills)
   ```
4. Pressed Enter.

**Result:** `Skill Count` = **170,578** (total skill mentions across all postings, not distinct skills).

## Part D — Job Count (distinct jobs represented in skills table)

Table: `data_job_skills`

1. Clicked into another empty calculation area cell.
2. Typed:
   ```
   Job Count := DISTINCTCOUNT(data_job_skills[job_id])
   ```
3. Pressed Enter.

**Result:** `Job Count` = **32,672** — matches the original source row count exactly, confirming no jobs were lost during the `job_skills` explode step in Stage 1, and `job_id` integrity holds across the relationship.

## Part E — Average Skills per Job

Table: `data_job_skills`

1. Clicked into another empty calculation area cell.
2. Typed:
   ```
   Avg Skills per Job := DIVIDE([Skill Count], [Job Count])
   ```
3. Pressed Enter.

`DIVIDE` used instead of the `/` operator — safely handles division-by-zero by returning blank rather than an error.

**Result:** `Avg Skills per Job` = **5.2**

## Measures Summary (so far)

| Measure | Table | Formula | Result |
|---|---|---|---|
| Median Salary | data_jobs_all | `MEDIAN(data_jobs_all[salary_year_avg])` | 115,000 |
| Median Salary US | data_jobs_all | `CALCULATE(MEDIAN(...), job_country = "United States")` | 118,940 |
| Skill Count | data_job_skills | `COUNTROWS(data_job_skills)` | 170,578 |
| Job Count | data_job_skills | `DISTINCTCOUNT(data_job_skills[job_id])` | 32,672 |
| Avg Skills per Job | data_job_skills | `DIVIDE([Skill Count], [Job Count])` | 5.2 |

## Key Lessons From This Part

1. Calculation area cells can look "empty" or truncated purely due to column width — always widen before assuming a measure failed.
2. `DISTINCTCOUNT` on the "many" side of a relationship is the correct way to get a true entity count (e.g. jobs) when that table has repeated keys.
3. `DIVIDE(measure_a, measure_b)` is the safe DAX pattern for ratios — avoids divide-by-zero errors that a plain `/` would throw.
4. Cross-checking a derived measure (Job Count = 32,672) against a known source fact (original row count) is a good sanity check for catching data-loss issues early.

---

## Next: Stage 3 continued — remaining measures

Still to build: skill-level median salary (for README Q4, top 10 skills by pay), skill likelihood/frequency %, and any region/country breakdowns needed for README Q2 beyond the US figure already built.

---
---

# Stage 3 — DAX Measures (continued): Skill-Level Salary & the Filter Direction Problem

**Goal:** Build a measure that returns median salary correctly filtered by skill (needed for README Q3/Q4), and understand why a naive approach fails.

## Part F — First attempt: plain MEDIAN measure (failed)

Table: `data_jobs_all`

1. Clicked into an empty calculation area cell.
2. Typed:
   ```
   Median Salary by Skill := MEDIAN(data_jobs_all[salary_year_avg])
   ```
3. Pressed Enter.

**⚠️ Issue hit:** Tested by dragging into a PivotTable with `job_skills` in Rows — every skill under every job title showed the exact same value (115,000), identical to the unfiltered overall median. The measure was not responding to the skill filter at all.

**Diagnosis:** Formula alone can't create filtering — filtering happens via *filter context*, which is supposed to flow through the relationship when a field like `job_skills` is used in a PivotTable row. The relationship built in Stage 2 is **single-direction**: `data_jobs_all` (1) → `data_job_skills` (*). Filtering FROM `data_jobs_all` correctly filters `data_job_skills` (proven in Stage 2's test), but the reverse — filtering by something on the "*" side (`job_skills`) to affect the "1" side (`salary_year_avg`) — does **not** happen automatically with a one-directional relationship.

**Note on where a measure is typed:** which table's calculation area a measure is created in is purely organizational — it does not affect what data the measure can access or how filtering works. What actually determines behavior is the DAX formula itself (and, for plain dragged fields, which table's column was used — see Part H below).

## Part G — Fix: CROSSFILTER

1. Edited the same measure:
   ```
   Median Salary by Skill := CALCULATE(MEDIAN(data_jobs_all[salary_year_avg]), CROSSFILTER(data_jobs_all[job_id], data_job_skills[job_id], BOTH))
   ```
2. Pressed Enter.

**What this does:** `CROSSFILTER` temporarily overrides the relationship's default one-way behavior, forcing it to filter in `BOTH` directions just for this calculation — allowing a filter on `job_skills` to flow back and correctly narrow `salary_year_avg`.

**⚠️ Side-note debugging trap hit:** After editing the formula, checked the value directly in the Power Pivot calculation area and still saw one flat number (115,000) — looked like the fix hadn't worked. **This was a false alarm**: the calculation area always shows a single unfiltered value for any measure, regardless of the formula, because there is no row/column filter context there (it's just a formula-verification space, not a live preview per category). The correct way to test is to view the measure inside a PivotTable with `job_skills` in Rows, where real filter context exists.

**Result after correct testing in a PivotTable:** Median salary now varies correctly by skill (e.g., under Business Analyst: different values per skill instead of one repeated number).

## Part H — Second bug: the Count of job_id field also frozen

While verifying Part G, noticed **`Count of job_id`** in the same PivotTable was *also* stuck at one repeated value per job title (e.g., 86 for every skill under Cloud Engineer) — including skills that clearly weren't used for that role (confirmed because `Median Salary by Skill` correctly showed blank for those same rows, while the count still showed 86).

**Diagnosis:** Same root cause as Part F — `job_id` had been dragged into Values from the **`data_jobs_all`** field list (the "1" side), which doesn't receive a filter from `job_skills` without the same one-directional relationship limitation. This was a plain aggregated field (not a custom measure), so `CROSSFILTER` wasn't the fix here.

**Fix:**
1. Removed `Count of job_id` (sourced from `data_jobs_all`) from the PivotTable's Values area.
2. In the Fields pane, expanded **`data_job_skills`** instead → dragged its `job_id` field into Values (defaults to Count).

**Why this works without CROSSFILTER:** `job_id` and `job_skills` both live in the *same table* (`data_job_skills`), so counting `job_id` from that table is filtered by `job_skills` automatically via ordinary row-level table filtering — no relationship-crossing required at all.

## Part I — Full verification

Checked across multiple job titles (Business Analyst, Cloud Engineer, Data Analyst) with both fixed fields in the PivotTable:
- Skills not actually used for a given job title correctly show blank/very low count and blank median salary.
- Skills genuinely used show sensible, varying counts and medians per skill (e.g., under Data Analyst: `aws` = 477 postings @ 100,500 median; `airtable` = 9 postings @ 90,000 median).
- Row-level totals per job title (e.g., Business Analyst = 3,474 total skill mentions) look like plausible aggregate figures.

**Both fields now confirmed working correctly across the full dataset, not just a couple of spot-checked rows.**

## Key Lessons From This Part

1. **A DAX formula alone doesn't create filtering — filter context does.** The same `MEDIAN(...)` formula behaves completely differently depending on what's filtering it at the time it's evaluated (a PivotTable row, a CALCULATE filter, etc.).
2. **One-directional relationships only propagate filters one way.** Filtering from the "1" side correctly filters the "*" side automatically; the reverse requires an explicit override like `CROSSFILTER(..., BOTH)`.
3. **The Power Pivot calculation area is not a live per-category preview** — it always shows one unfiltered value, regardless of the formula. Real testing must happen inside a PivotTable (or similar) where genuine row/column filter context exists.
4. **Plain dragged fields (not measures) inherit the same one-directional relationship limitation** as un-fixed measures — but the fix differs: if the field and its filtering column live in the *same* table, no CROSSFILTER is needed at all; simply source the field from the correct table.
5. **Cross-checking two independent fields against each other** (median salary blank vs. count still showing a value) was what actually exposed the second bug — a good general debugging habit: verify multiple related fields together, not just one in isolation.

---

## Next: Stage 3 continued — remaining measures

Still to build: Skill Likelihood % (for the combo chart in README Q4 — bars = median salary, line = likelihood), and any additional region/country breakdowns for README Q2.

---
---

# Interlude — Excel Freeze & Recovery

**What happened:** Workbook was left open in Power Pivot overnight/across a break. On return, Excel's status bar showed "Excel PivotTables and related objects are being updated from the data model" and the app stopped responding to any input (window wouldn't move, no clicks registered in either the Excel or Power Pivot windows).

## Diagnosis before acting

1. Checked **Task Manager** — Excel showed active (non-zero) CPU and ~985 MB memory usage, initially suggesting it might still be genuinely processing rather than hard-frozen (large model + several CROSSFILTER measures can make a full Data Model refresh slow).
2. Waited a further period to rule out a slow-but-working recalculation.
3. Re-tested responsiveness directly (tried moving/clicking the window) — confirmed **fully unresponsive**, not just slow — safe to conclude it was genuinely hung, not merely busy.

## Recovery steps taken

1. Attempted **End Task** via Task Manager on both the Excel process and the separate Power Pivot process — did not close the application (fully hung, past the point Task Manager could gracefully terminate it).
2. Performed a **full system restart** as the next safe option.
3. On reopening Excel, the **Document Recovery** pane appeared automatically, offering two versions:
   - `[Original]` — last manually saved version, timestamped 11:44
   - Autosaved version, timestamped **12:18** (34 minutes newer)
4. Opened the newer autosaved version first (rather than assuming the manual save was the safest bet) and **verified it against known-correct behavior** before trusting it: checked that `Count of job_id` and `Median Salary by Skill` both varied correctly per skill (the exact fix from Stage 3, Steps 7–8) across multiple job titles, not just spot-checked once.
5. Confirmed the autosaved version was fully intact and up to date → saved it immediately (Ctrl+S) to lock it in as the working file. (Excel saved it as a new file at the same location rather than overwriting the original filename — acceptable outcome, no data lost, just a naming change to be aware of going forward.)

**Result:** No work lost. All Stage 1–3 progress (including both CROSSFILTER-related fixes) recovered intact.

## Key Lessons From This Episode

1. **Don't assume "not responding" means frozen** — check Task Manager for active CPU/memory movement first; large Data Model recalculations can legitimately take several minutes and look identical to a freeze at a glance.
2. **Confirm true unresponsiveness before force-closing** (e.g., try moving the window) — force-closing too early risks losing work that was actually just slow, not stuck.
3. **When Document Recovery offers multiple versions, don't default to the manually-saved one** — compare timestamps, and prefer the newer autosave if it exists, since it may contain more recent work.
4. **Never trust a recovered file blindly — verify it against known-correct behavior** before treating it as safe to keep working on. Re-checking the specific fixes made earlier (not just "does it open OK") is what actually confirms integrity.
5. **A recovered file may save under a new filename rather than overwriting the original** — this is normal and not data loss, but it's worth noting so future sessions don't accidentally reopen an outdated file from Recent Files.

---

## Next: Stage 3 continued — remaining measures

Still to build: Skill Likelihood % (for the combo chart in README Q4), and any additional region/country breakdowns for README Q2.

---
---

# Stage 3 — DAX Measures (final piece): Skill Likelihood %

**Goal:** Build a measure showing what % of postings for a given job title mention each skill — powers the line series in the README's Q4 combo chart (bars = median salary, line = likelihood).

## Part J — First attempt: over-applied CROSSFILTER (failed)

Table: `data_jobs_all`

1. Typed:
   ```
   Skill Likelihood % := DIVIDE(DISTINCTCOUNT(data_job_skills[job_id]), CALCULATE(DISTINCTCOUNT(data_jobs_all[job_id]), CROSSFILTER(data_jobs_all[job_id], data_job_skills[job_id], BOTH)))
   ```
2. Pressed Enter.

**⚠️ Issue hit:** Tested in a PivotTable (Job Title → Skill in Rows) — every row showed **1 (100%)**, regardless of skill.

**Diagnosis:** This was the opposite mistake from the earlier `Median Salary by Skill` fix — `CROSSFILTER(..., BOTH)` was applied where it wasn't needed, and it actively broke the calculation. Forcing bidirectional filtering meant the **denominator** (meant to represent the *total* jobs for a job title, ignoring the skill filter) ended up **also being filtered by skill** — making numerator and denominator nearly identical, hence a ratio of ~1 everywhere.

**Key insight:** the relationship (`data_jobs_all` → `data_job_skills`) is one-directional by design — a skill filter on `data_job_skills` was never going to reach `data_jobs_all` in the first place. No override was needed here at all; `CROSSFILTER` was only necessary for `Median Salary by Skill` because that measure needed the filter to flow in the *reverse* direction (which doesn't happen by default). This measure needed the *opposite* — the default one-directional behavior left alone.

## Part K — Corrected formula

1. Edited the measure:
   ```
   Skill Likelihood % := DIVIDE(DISTINCTCOUNT(data_job_skills[job_id]), DISTINCTCOUNT(data_jobs_all[job_id]))
   ```
2. Pressed Enter.

**Why this works:**
- Numerator (`data_job_skills[job_id]`) — naturally filtered by both job title *and* skill, since it's on the "many" side where both filters apply directly.
- Denominator (`data_jobs_all[job_id]`) — filtered only by job title; the skill filter simply doesn't reach it (by the relationship's default one-way design), giving the true total job count for that title.

**⚠️ Re-confirmed false alarm:** Checked the value in the Power Pivot calculation area first and saw **1** again — this was expected and not a bug (same reason as with `Median Salary by Skill`): with no filter context, both distinct counts cover the entire table and are equal by definition. Verified correctly instead inside a PivotTable.

**Result in PivotTable:** Values varied correctly, but initially appeared as `0.0` for most rows — traced to **decimal formatting**, not a calculation problem (differences existed but were too small to show at 1 decimal place as a plain number).

## Part L — Format as percentage

1. In Power Pivot, clicked the `Skill Likelihood %` measure in the calculation area.
2. **Home** tab (Power Pivot ribbon) → **Formatting** group → set **Format** to **Percentage**, decimals to 1.

**Result:** Values now display meaningfully, e.g. 17.1% for one skill category, 0.7% for a niche skill like `airflow` — clear, sensible spread across skills.

## Measures Summary — Complete Set (Stage 3)

| Measure | Table | Formula | Result / Behavior |
|---|---|---|---|
| Median Salary | data_jobs_all | `MEDIAN(data_jobs_all[salary_year_avg])` | 115,000 |
| Median Salary US | data_jobs_all | `CALCULATE(MEDIAN(...), job_country = "United States")` | 118,940 |
| Skill Count | data_job_skills | `COUNTROWS(data_job_skills)` | 170,578 |
| Job Count | data_job_skills | `DISTINCTCOUNT(data_job_skills[job_id])` | 32,672 |
| Avg Skills per Job | data_job_skills | `DIVIDE([Skill Count], [Job Count])` | 5.2 |
| Median Salary by Skill | data_jobs_all | `CALCULATE(MEDIAN(...), CROSSFILTER(..., BOTH))` | Varies correctly per skill |
| Skill Likelihood % | data_jobs_all | `DIVIDE(DISTINCTCOUNT(data_job_skills[job_id]), DISTINCTCOUNT(data_jobs_all[job_id]))` | Varies correctly per skill, formatted as % |

**Plus one corrected plain field:** `job_id` Count sourced from `data_job_skills` (not `data_jobs_all`) for accurate per-skill counts in PivotTables.

## Key Lessons From This Part

1. **CROSSFILTER is not a general-purpose fix — apply it only when the specific filter direction actually needs reversing.** Using it reflexively (because it worked for a similar-looking problem earlier) broke a measure that didn't need it.
2. **Diagnose by asking "which direction does this specific calculation need filtered?"** rather than pattern-matching to a previous fix. `Median Salary by Skill` needed skill→salary (reverse of default); `Skill Likelihood %`'s denominator needed to explicitly NOT be affected by the skill filter (the default behavior already correct).
3. **The calculation area's single flat value is expected for any ratio-of-two-DISTINCTCOUNTs measure when both counts cover the same unfiltered universe** — not a diagnostic signal on its own; always verify in a real filter context.
4. **A measure can be calculating correctly but *look* broken due to display formatting alone** — small percentage differences rounded to 1 decimal as a plain number can appear identical; setting the measure's format to Percentage in Power Pivot (so it's correct everywhere it's used) resolved this cleanly.

**Stage 3 (DAX Measures) — Complete.** Full measure set now covers all four README analysis questions.

---

## Next: Stage 4 — PivotTables, PivotCharts & Analysis

Building the actual PivotTables and PivotCharts for each of the four README questions, using the measures built in Stage 3.

---
---

# Stage 4 — PivotTables & Charts

## Part A — README Q1: "Do more skills get you better pay?"

**Goal:** Compare `Avg Skills per Job` against `Median Salary` per job title, to test the README's claim that roles requiring more skills tend to pay more.

### Build the PivotTable

1. **Insert** tab → **PivotTable** → **Use this workbook's Data Model** → **New Worksheet** → OK.
2. Renamed the sheet: right-click sheet tab → **Rename** → `Salary_Vs_Skills` (matches original README naming convention).
3. In PivotTable Fields: dragged `job_title_short` into **Rows**; `Median Salary` and `Avg Skills per Job` into **Values**.

**⚠️ Issue hit & fixed:** `Median Salary` did not appear in the Fields pane at all, even though it calculated correctly (115,000) in the Power Pivot calculation area.
**Diagnosis:** The measure had been accidentally set to **"Hide from Client Tools"** — a Power Pivot setting that keeps a measure fully functional in the Data Model but hides it from PivotTable field lists. The greyed-out appearance of the measure in the calculation area was the visual clue.
**Fix:** Right-clicked the `Median Salary` measure in the calculation area → unchecked **"Hide from Client Tools"**. Measure then appeared correctly in the Fields pane.

**Result — PivotTable output (10 job titles):**

| Job Title | Median Salary | Avg Skills per Job |
|---|---|---|
| Business Analyst | 85,000 | 3.5 |
| Cloud Engineer | 90,000 | 4.9 |
| Data Analyst | 90,000 | 3.7 |
| Data Engineer | 125,000 | 7.0 |
| Data Scientist | 127,500 | 5.1 |
| Machine Learning Engineer | 107,550 | 5.4 |
| Senior Data Analyst | 111,175 | 4.4 |
| Senior Data Engineer | 147,500 | 8.2 |
| Senior Data Scientist | 155,000 | 5.3 |
| Software Engineer | 99,150 | 5.6 |
| **Grand Total** | **115,000** | **5.2** |

Clear pattern: Business Analyst/Data Analyst cluster low on both dimensions; Senior Data Engineer/Senior Data Scientist cluster high on both — supports the "more skills → better pay" insight.

### Build the scatter chart

**⚠️ Issue hit:** Attempted to build a PivotChart directly (Scatter type) → Excel error: *"You cannot create this chart type with data inside a PivotTable... select a different chart type or copy the data outside the PivotTable."*
**Diagnosis:** Excel PivotCharts don't support Scatter (XY) charts natively, since PivotCharts are built around category-based axes, while scatter plots need two independent numeric axes.
**Fix (chosen approach):** Copied the PivotTable's values out to a plain range using **Copy → Paste Special → Values Only**, breaking the live PivotTable link, then built a standard **Insert → Charts → Scatter** chart from the plain data.

**⚠️ Second issue hit:** The first scatter attempt plotted both `Median Salary` and `Avg Skills per Job` as two separate Y-series against row number (1–10) as a shared X-axis — not what we wanted. `Avg Skills per Job` appeared as a flat line near zero, since its small values (3.5–8.2) were squashed on the same 0–180,000 scale as salary.
**Fix:** Right-click chart → **Select Data** → removed the `Avg Skills per Job` series entirely → edited the remaining `Median Salary` series to set:
- **Series X values** → the `Avg Skills per Job` column of numbers
- **Series Y values** → the `Median Salary` column of numbers
- **Series name** → typed directly as plain text (`Job Titles by Skill & Pay`) — no `=` or quotes needed when typing a literal name; `=` syntax only appears automatically when referencing a cell instead.

**⚠️ Third issue hit:** After the above fix, the chart still looked wrong — points spread evenly across x = 0–12 instead of clustering realistically between ~3.5 and ~8.2.
**Diagnosis:** Checked **Select Data → Edit Series** and found the X/Y cell references were pointing to mismatched, incorrect row ranges (X range accidentally spanned two columns; X and Y ranges covered different row spans entirely), so points were being plotted from misaligned data.
**Fix:** Re-selected both ranges carefully by dragging directly over the correct single-column ranges, covering the **exact same 10 rows** (the 10 job titles, excluding the Grand Total row) for both Series X values and Series Y values.

**Result:** Correct scatter — X-axis ~3.0–8.5 (skills), Y-axis ~80,000–155,000 (salary), 10 points, visible upward trend from bottom-left to top-right.

### Add job title labels

1. Selected the data point series → right-click → **Add Data Labels** → **Add Data Labels** (defaults to showing Y-value).
2. Right-click labels → **Format Data Labels** → checked **"Value From Cells"** → selected the job title text column as the label source.
3. Unchecked "Y Value" (and X Value) so only the job title text displays per point.

**Result:** Each point now labeled with its job title (e.g., "Senior Data Scientist" at the top-right, "Business Analyst" at the bottom-left). Some overlap remains in the mid-cluster (Software Engineer / Cloud Engineer / Machine Learning Engineer labels crowd together) — noted as a polish item, not a data issue.

**Status: Q1 scatter chart is functionally complete and correctly plotted.**

## Pending polish (deferred, to revisit after all four README charts are built)

- Add a proper chart title (currently shows the series name placeholder)
- Add axis titles (X = "Avg Skills per Job", Y = "Median Salary")
- Manually nudge/reposition overlapping labels in the mid-cluster for readability
- Consider removing/hiding the default legend entry ("Job Titles by Skill & Pay") since it's redundant with the chart title

## Key Lessons From This Part

1. **"Hide from Client Tools" makes a measure invisible in PivotTable field lists while still fully functional in the model** — a greyed-out calculation area entry is the visual tell; toggle it off via right-click if a working measure won't appear where expected.
2. **Excel PivotCharts cannot create Scatter (XY) charts directly** — copy PivotTable output to plain values (Paste Special → Values Only) first, then build a standard chart from that.
3. **A scatter chart needs one series with explicit X and Y cell ranges, not two separate Y-series plotted against row number** — check Select Data → Edit Series to confirm X/Y are correctly assigned before assuming a chart is "just messy."
4. **Always verify Select Data's X/Y ranges cover the same rows and single columns** — mismatched or multi-column ranges silently misplot points in a way that can look like "random" scatter rather than an obvious error.
5. **Data labels default to showing the Y-value; "Value From Cells" is required to label points with a separate text field** (like job title) instead of the plotted number.

## Part B — Q1 Polish Pass (using Luke's original chart as a style reference)

**Reference used:** the original tutorial's version of this chart — axes swapped (salary on X), trendline added, callout-style labels, $K axis formatting, descriptive title matching the README question.

### Swapped axes
Right-click chart → **Select Data** → Edit series → swapped **Series X values** to `Median Salary`, **Series Y values** to `Avg Skills per Job` (previously the reverse).

### Added trendline
Selected the data point series → right-click → **Add Trendline** → **Linear**.

### Updated chart title
Changed to: `Do more Skills get you better Pay?` — matches the actual README question.

### Y-axis: start at 3, whole numbers only
Right-click Y-axis → **Format Axis** → **Minimum** set to `3.0` (manual, not Auto) → **Number** format code set to `0` (no decimals).

**⚠️ Issue hit:** After setting Y-axis minimum to 3, the X-axis line (and its salary labels) jumped to the **top** of the chart instead of staying at the bottom.
**Diagnosis:** With Y-axis minimum no longer at 0, Excel had nowhere natural to draw the X-axis (which normally sits at Y=0) and defaulted to placing it at the top of the visible range instead.
**Fix — first attempt (wrong axis):** Tried setting "Vertical axis crosses" on the **X-axis's** Format Axis pane — this did not fix it. This setting controls where the *Y-axis line* crosses the *X-axis*, not the reverse — edited the wrong property.
**Fix — correct approach:** Right-click the **Y-axis** instead → Format Axis → **"Horizontal axis crosses"** → selected **"Axis value"** and typed `3`. This is the setting that actually controls where the X-axis line sits vertically.
**⚠️ Sub-issue:** First attempt at this still left the X-axis line sitting between Y=8–9 instead of at Y=3 — traced to the Axis value box not genuinely containing `3` (likely a stray/incorrect value left over from a previous edit). Fixed by clicking into the box, clearing it fully, and re-entering `3` directly, then pressing Enter to confirm before closing the pane.
**Result:** X-axis line and labels correctly sit at the bottom of the chart, at the Y=3 gridline.

### X-axis number formatting
Right-click X-axis → **Format Axis** → **Number** → custom format code applied to display values like `$80K` instead of `80,000`.

### Data labels — callout style
**⚠️ Issue hit:** Data labels didn't carry over after the axis swap (series was edited/recreated) — needed to be added fresh.
**Fix:**
1. Selected the data point series → right-click → **Add Data Labels**.
2. Right-click labels → **Format Data Labels** → checked **"Value From Cells"** → selected the job title column as the label source → unchecked "Y Value" so only job title text shows.
3. Applied **callout-style label shape** (speech-bubble with connector line to each point) via the Label Shape option in the Format Data Labels pane — this automatically added connector lines and gave Excel room to place text away from crowded points, resolving the earlier mid-cluster overlap without needing to manually drag every label.

**Final result:** Clean, fully-labeled scatter chart — correct axes and scale, $K-formatted salary axis, linear trendline, descriptive title, and readable callout labels for all 10 job titles. Matches (and in some respects exceeds) the polish level of the original reference chart.

**Q1 — chart building AND polish, fully complete.**

---

## Next: Stage 4 continued — README Q2, Q3, Q4

---
---

# Stage 4 — README Q2: "What's the salary for data jobs in different regions?"

**Goal:** Compare overall median salary against US-only median salary, per job title, using the existing `Median Salary` and `Median Salary US` measures from Stage 3.

## Part A — Scoping decision: UK-specific comparison considered and dropped

Given the project owner is UK-based, considered building a UK-specific salary measure (mirroring `Median Salary US`) for a more personally relevant comparison.

**Investigated data support first, before building anything:**
- Checked `job_location` (city-level) filter dropdown in `data_jobs_all` — found the field is highly fragmented (hundreds of individual cities worldwide, e.g. "Aberdeen, UK" mixed in among thousands of other city entries), meaning a clean UK rollup would require additional aggregation work.
- Also noted: `salary_year_avg` is denominated in **USD regardless of job location** — a "UK median salary" measure would still be a USD figure for UK-based postings, not a true GBP figure. Converting to £ would require applying an exchange rate, introducing a value that goes stale over time and doesn't reflect real paid currency.

**Decision:** Dropped the UK-specific measure. Kept the project scoped to the original README structure (overall vs. US median), which is simpler, better-supported by the data, and avoids introducing a conversion/aggregation caveat that would weaken the analysis's credibility. Noted as a documented scoping decision rather than an oversight.

## Part B — Build the PivotTable

1. **Insert** tab → **PivotTable** → **Use this workbook's Data Model** → **New Worksheet** → OK.
2. Renamed the sheet: `Salary_Analysis`.
3. Dragged `job_title_short` into **Rows**; `Median Salary` and `Median Salary US` into **Values**.

**Result — PivotTable output (10 job titles):**

| Job Title | Median Salary | Median Salary US |
|---|---|---|
| Business Analyst | 85,000 | 90,000 |
| Cloud Engineer | 90,000 | 115,000 |
| Data Analyst | 90,000 | 90,000 |
| Data Engineer | 125,000 | 125,000 |
| Data Scientist | 127,500 | 130,000 |
| Machine Learning Engineer | 107,550 | 150,000 |
| Senior Data Analyst | 111,175 | 110,000 |
| Senior Data Engineer | 147,500 | 150,000 |
| Senior Data Scientist | 155,000 | 155,000 |
| Software Engineer | 99,150 | 125,000 |
| **Grand Total** | **115,000** | **118,940** |

**Insight confirmed:** US salaries are equal to or higher than overall median for every job title. Largest gaps: Machine Learning Engineer (~40% higher in the US), Cloud Engineer (~28% higher), Software Engineer (~26% higher). Data Engineer and Senior Data Scientist show no gap. This directly supports the README's stated insight about salary disparity being most notable in high-tech roles.

## Part C — Currency formatting

**Goal:** Display salary values as `$115,000` instead of plain `115,000`.

**Step 1 — formatted at the measure level (Power Pivot):**
1. Power Pivot → Manage → `data_jobs_all` table tab.
2. Selected `Median Salary` measure → Home tab (Power Pivot ribbon) → Formatting group → Category = **Currency**, Symbol = **$**, Decimal places = **0**.
3. Repeated for `Median Salary US`.
4. **Verified directly in the calculation area** — both measures correctly displayed with `$` there (e.g. `Median Salary: $115,000`), confirming the measure-level format had genuinely saved.

**⚠️ Issue hit:** Despite the confirmed measure-level formatting, the PivotTable itself displayed **inconsistent formatting** — the Grand Total row showed no `$` and neither did the individual job title rows (i.e., formatting appeared not to be applying anywhere in the PivotTable, despite working correctly in the calculation area).
**Attempted fixes that did NOT resolve it:**
- Refreshing the PivotTable (right-click → Refresh) — no change.
- Removing and re-adding both fields to the Values area — no change.

**Root cause found:** the PivotTable has its own **per-field Number Format** setting (accessed via **Value Field Settings → Number Format**) which sits on top of and can override the measure's own formatting. This field had defaulted to General/Number rather than inheriting Currency from the measure.

**Fix:**
1. Right-clicked a value in the `Median Salary` column → **Value Field Settings** → **Number Format** button.
2. Set to **Currency**, **$**, **0** decimal places → OK → OK.
3. Repeated for `Median Salary US`.

**Result:** All rows, including Grand Total, now display `$` consistently.

## Key Lessons From This Part

1. **Before building a new regional/segment measure, check the data actually supports it** — a quick look at field granularity (city-level vs. country-level) and underlying currency/unit assumptions can save building something that looks precise but is actually shaky or misleading.
2. **A measure's format set in Power Pivot and a PivotTable field's own Number Format are two separate settings** — the PivotTable-level one can silently override the measure-level one. If a confirmed-correct measure format isn't showing in a PivotTable, check **Value Field Settings → Number Format** specifically, not just the measure itself.
3. **Verify formatting fixes at the source first** (the calculation area, in this case) before assuming a fix failed — this correctly separated "is the measure formatted right?" (yes) from "is the PivotTable displaying it right?" (no, separate issue), avoiding wasted effort re-doing the wrong fix repeatedly.

---

## Next: Stage 4, Q2 — build the clustered column chart

---
---

# Stage 4, Q2 — Chart Building & Polish

## Part D — Third measure: Median Salary Non-US

**Motivation:** Reviewed the original tutorial's (Luke's) version of this same chart as a style reference — it included a third "Non-US" comparison column. However, cross-checked his actual numbers and found his "Median Salary" (overall) column was implausible — showing values *higher* than both his US and Non-US figures (e.g. Cloud Engineer: $197,500 overall vs $115,000 US and $89,100 Non-US), which is mathematically impossible for a true blended median. A country slicer was visibly left active/filtered (on "Argentina") in that screenshot, strongly suggesting his figure was accidentally filtered rather than a genuine global median.

**Decision:** Build the third measure independently and correctly, rather than copying the reference's structure — a good example of using a reference for inspiration on presentation while independently verifying the underlying numbers.

1. Power Pivot → Manage → `data_jobs_all` tab.
2. Typed:
   ```
   Median Salary Non-US := CALCULATE(MEDIAN(data_jobs_all[salary_year_avg]), data_jobs_all[job_country] <> "United States")
   ```
3. Formatted as Currency ($, 0 decimals) at the measure level.
4. Added to the PivotTable's Values area alongside the existing two measures.
5. **Also had to set Value Field Settings → Number Format → Currency on this new field specifically** (same override issue as Part C) — confirmed same fix required for every newly-added field, not just a one-time occurrence.

**Result — verified table (all 10 job titles):**

| Job Title | Median Salary | Median Salary US | Median Salary Non-US |
|---|---|---|---|
| Business Analyst | 85,000 | 90,000 | 75,000 |
| Cloud Engineer | 90,000 | 115,000 | 89,100 |
| Data Analyst | 90,000 | 90,000 | 90,000 |
| Data Engineer | 125,000 | 125,000 | 123,500 |
| Data Scientist | 127,500 | 130,000 | 119,550 |
| Machine Learning Engineer | 107,550 | 150,000 | 101,029 |
| Senior Data Analyst | 111,175 | 110,000 | 111,175 |
| Senior Data Engineer | 147,500 | 150,000 | 147,500 |
| Senior Data Scientist | 155,000 | 155,000 | 155,000 |
| Software Engineer | 99,150 | 125,000 | 89,100 |
| **Grand Total** | **115,000** | **118,940** | **111,175** |

**Sanity check passed:** Non-US values sit consistently at or below Overall; US values sit consistently at or above Overall — mathematically correct blended-median behavior across every row, unlike the flawed reference.

**Notable findings for README write-up:**
- Machine Learning Engineer: largest US premium (~48% higher than Non-US)
- Cloud Engineer: ~29% US premium
- Senior Data Scientist: identical across all three columns — no US premium for this role, a genuine outlier worth calling out
- Data Analyst: also identical across Overall and US — small/no gap

## Part E — Build the chart

1. Clicked inside the PivotTable → **PivotTable Analyze** → **PivotChart** → **Clustered Column** → OK.
   (Native PivotChart support — no workaround needed here, unlike Q1's scatter.)
2. Chart automatically picked up all three measures as three bars per job title once the third field was added to the table.

## Part F — Polish pass

1. **Title:** set to `What's the salary for data jobs in different regions?` (matches README question, same convention as Q1).
2. **Color palette:** Applied via **Chart Design → Change Colors → Palette 4 (Monochromatic)**. Deliberately tested legibility before committing — checked whether all three series remained distinguishable without relying on the legend ("squint test") before finalizing the monochrome choice over a more hue-varied "Colorful" palette.
3. **X-axis labels tilted:** Right-click X-axis → **Format Axis** → **Text Options** → **Custom angle: -45°** — fixed the awkward two-line wrapping of longer job titles (e.g. "Business Analyst" splitting across two lines) so each label reads on a single diagonal line.
4. **Y-axis formatting:** Applied the same `$#,##0,"K"` custom number format used in Q1, displaying $20K, $40K etc. instead of $20,000, $40,000.
5. **Y-axis title** added: "Median Salary (USD)".
6. **Data labels deliberately omitted** — with 3 series × 10 categories (30 bars total), individual data labels would clutter badly; the formatted Y-axis and legend were judged sufficient for readability. Noted as a deliberate choice, not an oversight (unlike Q1's scatter, where only 10 total points made per-point labels appropriate).

## Key Lessons From This Part

1. **Cross-check a reference/tutorial's numbers before adopting its structure** — a visually convincing chart can still contain a data error (here, an active slicer left filtering the "overall" figure to a single country). Independently verifying the logic (Non-US and US should bracket the Overall) caught this before it could be copied into the rebuild.
2. **The Value Field Settings → Number Format override (from Part C) applies per-field, not globally** — each newly added measure needs the same fix reapplied, even if the underlying Power Pivot measure format is already correct.
3. **Monochromatic palettes are a legitimate, professional choice for multi-series charts, but warrant a legibility check** (can the series be told apart without the legend?) before committing, especially with 3+ series — a Colorful palette is the safer default when in doubt, but not required if monochrome tests out fine.
4. **Not every chart needs data labels** — the right amount of on-chart detail depends on series/category count; a clean axis + legend can be the more professional choice than cluttering a chart with a label on every single bar. This is a deliberate design decision, not a shortcut.

**Q2 — chart building AND polish, fully complete.**

---

## Next: Stage 4 continued — README Q3, Q4

---
---

# Stage 4, Q3 — "What are the top skills of data professionals?"

## Part A — Build the PivotTable

1. **Insert** tab → **PivotTable** → **Use this workbook's Data Model** → **New Worksheet** → OK.
2. Renamed sheet: `Skill_Job_Analysis`.
3. Dragged `job_skills` (from `data_job_skills`) into **Rows**; `job_id` (from `data_job_skills` — the correctly-sourced count field established back in Stage 4/Q1) into **Values**.

**⚠️ Issue hit:** Needed to sort skills by count (descending) to see the top skills first. Clicking the **Row Labels** dropdown only offered alphabetical A-Z/Z-A sorting — no option to sort by the value column.
**Fix:** Right-clicked directly on a **value cell** in the `Count of job_id` column (not the header) → **Sort → Sort Largest to Smallest**. This sorts the whole table by that column's values, which the Row Labels dropdown cannot do.

**Result — top skills, verified against README's stated insight:**

| Rank | Skill | Count |
|---|---|---|
| 1 | sql | 18,500 |
| 2 | python | 17,689 |
| 3 | tableau | 7,046 |
| 4 | r | 6,929 |
| 5 | aws | 6,844 |
| 6 | excel | 6,264 |
| 7 | spark | 5,294 |
| 8 | sas | 4,806 |
| 9 | azure | 4,760 |
| 10 | java | 3,827 |

**Confirmed:** SQL and Python dominate by a wide margin (both far ahead of 3rd place Tableau — less than half of Python's count), matching the README's stated insight. AWS and Azure both appear in the top 10, supporting the README's point about cloud technology demand.

## Part B — Limit to Top 10 and exclude blank

**Note:** a `(blank)` row (count 3,187, ~rank 14) exists in the full skill list — representing postings with no skills listed (same category investigated back in Stage 3). Not a real skill, so excluded from the "top skills" ranking.

**Method used — Value Filters (cleaner than manually unchecking blank):**
1. Row Labels filter dropdown → **Value Filters** → **Top 10...**
2. Confirmed default: "Top" "10" "Items" by "Count of job_id" → OK.

**Why Top 10 specifically (not Top 15):** chosen deliberately because blank's rank (~14) falls outside Top 10, so this filter automatically excludes it without needing a separate manual step — Top 15 would have required also manually unchecking blank afterward.

## Part C — Build the chart

Chose a **horizontal Bar chart** (rather than matching Q1/Q2's vertical column style) — deliberate deviation, since horizontal bars read skill names cleanly at full length without needing tilted labels, and the top-to-bottom descending order reads naturally as a ranking for a "Top 10" question.

1. Clicked inside table → **PivotTable Analyze** → **PivotChart** → **Bar** → OK.

## Part D — Polish

1. **Title:** changed from default "Total" to `What are the top skills of data professionals?`
2. **Legend removed** — only one series ("Total"/Count), so the legend box was redundant; selected and deleted.
3. **X-axis number formatting (comma separators):**
   **⚠️ Issue hit:** Standard custom format code approach (`#,##0`) required typing `#`, which wasn't straightforwardly available on this UK Dell keyboard layout (not on the visible keycap; Alt Gr method didn't work either).
   **Fix — avoided the character entirely:** Right-click X-axis → **Format Axis** → **Number** → selected built-in **Category: Number** (not Custom) → checked **"Use 1000 separator (,)"** → **Decimal places: 0**. Achieves identical result (18,500 instead of 18500) through menu options only, no special character typing needed.

**Result:** Clean horizontal bar chart, correctly titled, no redundant legend, comma-formatted axis, clear visual dominance of SQL/Python over the rest of the top 10.

**Q3 — chart building AND polish, fully complete.**

## Key Lessons From This Part

1. **Sorting a PivotTable by a value column (not the row field) requires right-clicking a value cell directly** — the Row Labels dropdown only ever offers alphabetical sorting of the row field itself, never the aggregated values next to it.
2. **Value Filters → Top N is a clean way to both limit and implicitly exclude unwanted categories** in one step, if the unwanted category's rank happens to fall outside the chosen N — worth checking the excluded item's actual rank first to confirm this shortcut will work before relying on it.
3. **Chart orientation doesn't need to be uniform across every chart in a project** — a horizontal bar suited this specific "ranked top 10" question better than forcing consistency with the vertical columns used elsewhere; matching data-to-chart-type sensibly matters more than visual uniformity alone.
4. **Keyboard layout differences (UK vs US) can block typing certain special characters needed for custom format codes** — when a needed symbol isn't easily accessible, check whether Excel's built-in format categories (Number, Currency, etc.) can achieve the same visual result via checkboxes/dropdowns instead of typing a custom code.

---

## Next: Stage 4, Q4 — "What's the pay for the top 10 skills?" (the combo chart)

---
---

# Stage 4, Q4 — "What's the pay for the top 10 skills?" (Combo Chart)

## Part A — Build the PivotTable

1. **Insert** tab → **PivotTable** → **Use this workbook's Data Model** → **New Worksheet** → OK.
2. Renamed sheet: `Skill_Salary_Analysis`.
3. Dragged `job_skills` (from `data_job_skills`) into **Rows**; `Median Salary by Skill` and `Skill Likelihood %` (both from `data_jobs_all`, built in Stage 3) into **Values**. Confirmed a measure's home table doesn't restrict where it can be used — dragging fields from a different table's row context works correctly via the relationship.

## Part B — Filter to Top 10, matching Q3's skill list

**Decision:** deliberately kept the same Top 10 skills as Q3 (by frequency) rather than a different Top 10 by salary, so Q3→Q4 tell a connected story: "here are the most in-demand skills — here's what they pay and how likely each is to be requested."

1. Temporarily added `Count of job_id` to Values as a verification aid.
2. Applied **Value Filters → Top 10** by `Skill Likelihood %` (mathematically equivalent ranking to raw frequency, since both derive from the same postings) → confirmed the resulting 10 skills matched Q3's list exactly (sql, python, tableau, r, aws, excel, spark, sas, azure, java).
3. Removed `Count of job_id` from Values once confirmed (table reverted to just the two needed measures; row order reset as a side effect — re-sorted via right-click → Sort Largest to Smallest on `Skill Likelihood %`).

## Part C — Build the combo chart

**⚠️ Issue hit (first attempt):** Built the combo chart while `Count of job_id` was still present in Values (left over from the filtering step). The Custom Combination dialog then showed **three** series to configure instead of two, and the default assignment was wrong — `Skill Likelihood %` defaulted to Clustered Column (invisible next to salary's much larger scale) while `Count of job_id` was set to Line + Secondary Axis (not the series intended for the line).
**Fix:** Cancelled the dialog, removed `Count of job_id` from the PivotTable entirely (see Part B), then reopened **Insert → Charts → Combo Chart** with only the two intended series present — far simpler to configure correctly with two series rather than three.

**Correct chart build:**
1. Clicked inside PivotTable → **Insert** tab → **Charts** → **Combo Chart** icon → **Create Custom Combo Chart**.
2. Set:
   - `Median Salary by Skill` → **Clustered Column**, Secondary Axis unchecked
   - `Skill Likelihood %` → **Line**, Secondary Axis **checked**
3. Clicked OK.

**Result:** Correct combo chart — 10 salary columns on the primary ($) axis, one likelihood line on the secondary (%) axis, both scaled appropriately and clearly readable.

**Insight surfaced:** demand (likelihood %) varies far more dramatically across the top 10 skills than pay does — the likelihood line drops sharply after Python/SQL, while salary bars stay relatively similar across the whole top 10. Worth highlighting in the README write-up as a notable finding distinct from the original tutorial's framing.

## Part D — Polish

1. **Title:** `What's the pay for the top 10 skills?`
2. **Primary Y-axis:** formatted to $K style (e.g. $160K instead of 160000), matching Q1/Q2 convention.
3. **Axis titles added:** "Median Salary (USD)" on primary axis; secondary axis retains its "Skill Likelihood %" label from the field name.
4. **Legend:** left as-is, already clearly distinguishing the two series.
5. X-axis skill-name labels were already angled/readable by default — no further adjustment needed.

**Q4 — chart building AND polish, fully complete.**

## Part E — Reference check against the original tutorial's Q4 chart

Reviewed Luke's version of this same chart before finalizing. Found two red flags suggesting unreliable underlying data:
1. A final row showing **225% skill likelihood** — mathematically impossible (likelihood is bounded 0-100%), most likely caused by an included Grand Total row or an aggregation error, similar in nature to the Argentina-slicer bug found in his Q2 chart.
2. All ten salary bars appeared suspiciously uniform (~$90-100K each), lacking the visible variation this rebuild's version correctly shows (aws $135,000 down to excel $92,500) — suggesting a possible calculation or scope error in his version.

**Decision:** did not adopt his chart's structure or data as a reference for this question, unlike Q1/Q2 where his general *styling* (not data) was useful inspiration. Relied entirely on this project's own independently-verified numbers.

## Key Lessons From This Part

1. **Clean up a PivotTable's Values area to only the fields actually needed before opening a combo chart dialog** — leftover fields (like a verification-only count) get pulled into the chart type picker automatically and default to the wrong chart type/axis, creating avoidable rework.
2. **A combo chart with 2 series is far simpler to configure correctly than one with 3** — if a third field snuck in unintentionally, removing it first (rather than trying to exclude/hide it within the combo dialog) is the more reliable fix.
3. **Consistently sanity-checking reference/tutorial charts before adopting them paid off a third time** — an impossible percentage value (225%) and suspiciously flat bar heights were both catchable red flags before ever touching this project's own build.
4. **Axis titles are more valuable when a chart has two genuinely different numeric measures on two axes** (as in this combo chart and the Q1 scatter) than on a single-measure chart like Q3's bar chart, where the row labels and chart title already make the axis self-explanatory.

---

# PROJECT STATUS: All four README questions (Q1–Q4) — data, charts, and polish — complete.

## Final project structure
- **Stage 1:** Power Query ETL — `data_jobs_all` and `data_job_skills` queries, cleaned and loaded to Data Model
- **Stage 2:** Power Pivot — One-to-Many relationship via `job_id`
- **Stage 3:** DAX Measures — 7 measures covering all four analysis questions
- **Stage 4:** Four PivotTables + PivotCharts, one per README question, each independently verified against (and in some cases improved upon) the original tutorial's reference version:
  - `Salary_Vs_Skills` — Q1 scatter chart with trendline
  - `Salary_Analysis` — Q2 clustered column (3-way regional comparison)
  - `Skill_Job_Analysis` — Q3 horizontal bar (top 10 skills by frequency)
  - `Skill_Salary_Analysis` — Q4 combo chart (salary + likelihood for top 10 skills)

## Notable improvements over the original tutorial version
- Caught and avoided a slicer-contamination bug in the reference Q2 chart (Argentina-filtered "overall" median)
- Caught and avoided an impossible 225% likelihood value and flat/suspicious salary bars in the reference Q4 chart
- Added a genuinely useful third measure (`Median Salary Non-US`) not present in the original, enabling a more complete regional comparison
- Made a documented, reasoned scoping decision to exclude a UK-specific salary comparison given data limitations, rather than building something unreliable just because it seemed personally relevant

## Suggested remaining/optional work
- Final proofread pass across all four sheets/charts for consistency
- Consider a summary/dashboard tab combining highlights from all four questions (potentially with a slicer) as a stretch goal
- Write up the README narrative sections (insights, "so what") for the final portfolio piece, incorporating the notable findings and honest methodology/debugging narrative captured throughout this log

---
---

# Stage 5 — Final Consistency Pass

Reviewed all four charts (Q1–Q4) against a single checklist: title style/phrasing, color palette cohesion, axis number formatting, font consistency, chart sizing, gridline style.

**Issue found and fixed:** Q1's chart title read "Do more Skills get you better Pay?" (capitalized "Skills" and "Pay"), inconsistent with the other three titles' sentence-case convention ("What's the salary...", "What are the top skills...", "What's the pay..."). Corrected to `Do more skills get you better pay?` for consistency.

Also checked font consistency across all four chart titles/axis labels — standardized where needed.

**Consistency pass — complete.**

---
---

# Stage 6 — Adding Slicers (with rigorous verification)

## Part A — Reconsidering slicers after reviewing the original tutorial's approach

Initially decided against slicers (reasoning: each chart already shows its full comparison at a glance, and a slicer might work against that). After reviewing the original tutorial's version — which added a `job_country` slicer to its Q2 equivalent, and `job_title_short` slicers to its Q3/Q4 equivalents — reconsidered: a slicer doesn't remove the default full view, it just adds an *optional* way to drill in further. Decided this is a genuine value-add for a dashboard-style deliverable, not a contradiction of each chart's purpose.

**Important caveat identified up front:** reviewing the tutorial's slicer-filtered screenshots also re-confirmed (a third time) that its underlying measures are unreliable — with a country slicer active, its "overall" median salary figures were clearly being affected by the slicer selection (e.g., showing $197,500 for "Argentina" filtered view) rather than correctly representing a slicer-independent benchmark, and its skill-likelihood measure showed an impossible 238% total under a job-title slicer. This reinforced the need to **explicitly verify** this project's own measures behave correctly under slicer interaction, rather than assuming they would just because the underlying logic had tested fine in isolation.

## Part B — Added `job_country` slicer to Salary_Analysis (Q2)

1. Clicked inside the `Salary_Analysis` PivotTable → **PivotTable Analyze** → **Insert Slicer** → checked `job_country` → OK.

## Part C — First stress test revealed a genuine DAX behavior (not a bug, but an important gotcha)

**Test:** Selected "India" in the new slicer, checked all three salary measures.
**Result:** `Median Salary` correctly updated to reflect India-only data. However, `Median Salary US` and `Median Salary Non-US` both remained frozen at their original global values ($118,940 and $111,175 respectively) regardless of which country was selected in the slicer — they did not respond to the slicer at all.

**Root cause — understood, not just patched:** Both measures were originally written as:
```
Median Salary US := CALCULATE(MEDIAN(...), job_country = "United States")
```
In DAX, when `CALCULATE` sets an explicit filter on a column that **already has an active filter from elsewhere** (here, the slicer's filter on that same `job_country` column), the new filter **replaces** the existing one entirely, rather than combining with it. So no matter what the slicer selected, these measures always recalculated as if filtered to "all US postings, full stop," ignoring the slicer completely.

**Decision:** since these measures are intended for dashboard-style interactivity (where a viewer expects every visible number to respond consistently to a slicer), fixed both measures using `KEEPFILTERS`, which changes CALCULATE's behavior from "replace the existing filter" to "combine (AND) with the existing filter."

**Fix applied:**
```
Median Salary US := CALCULATE(MEDIAN(data_jobs_all[salary_year_avg]), KEEPFILTERS(data_jobs_all[job_country] = "United States"))

Median Salary Non-US := CALCULATE(MEDIAN(data_jobs_all[salary_year_avg]), KEEPFILTERS(data_jobs_all[job_country] <> "United States"))
```

## Part D — Verification: three independent test cases

Re-tested with three different slicer selections to confirm the fix, checking not just "did the number change" but "does the resulting logic make mathematical sense":

| Slicer selection | Median Salary | Median Salary US | Median Salary Non-US | Logic check |
|---|---|---|---|---|
| United States | $118,940 | $118,940 (matches) | *(blank)* | ✅ Overall = US when filtered to US only; Non-US correctly impossible/blank |
| India | e.g. $149,653 | *(blank)* | $149,653 (matches Overall) | ✅ India is already non-US, so "India AND Non-US" = "India"; US correctly blank (no US rows within an India filter) |
| Germany | $108,413 | *(blank)* | $108,413 (matches Overall) | ✅ Same logic as India — Non-US measure equals Overall; US correctly blank |

All three cases confirmed internally consistent, correct behavior — not just "a number changed," but the *relationships between the three measures* held up exactly as they mathematically should in each scenario.

## Key Lessons From This Part

1. **`CALCULATE`'s filter behavior on a column is "replace," not "combine," by default** — if a column already has an active filter (from a slicer, a row/column field, or another CALCULATE), a plain `column = value` filter inside CALCULATE overrides it entirely rather than intersecting with it.
2. **`KEEPFILTERS` changes this to combine (AND) with existing filters** — essential for any measure that needs to respect external filter context (like a slicer) rather than enforcing a fixed, independent condition.
3. **This is not something that shows up in isolated testing** — the original measures worked perfectly well in every PivotTable test throughout Stages 3–4, because no slicer was present to conflict with their filters. The issue only surfaced once a slicer was introduced, reinforcing the value of testing measures under every way they might actually be used, not just the way they were originally built and checked.
4. **Verifying a fix means checking the relationships between multiple measures, not just one number** — confirming `Median Salary US` went blank was necessary but not sufficient; checking that `Median Salary Non-US` correctly equaled `Median Salary` in the India/Germany cases (a non-obvious mathematical consequence of the fix) provided much stronger confidence the logic was genuinely correct, not just superficially "changed."
5. **Reviewing a flawed reference implementation's slicer behavior was directly useful** — the tutorial's clearly-broken slicer interaction (impossible percentages, filtered "overall" figures) motivated the decision to rigorously verify this project's own measures rather than assume correctness, catching a real (if subtle) issue before it reached the final deliverable.

---

## Next: consider adding slicers (with the same KEEPFILTERS pattern where relevant) to Skill_Job_Analysis (Q3) and Skill_Salary_Analysis (Q4), matching the tutorial's job_title_short slicer placement — then final project wrap-up.

---
---

# Interlude — Understanding *why* measures matched under slicer filtering (conceptual clarification)

After verifying the `KEEPFILTERS` fix (Part D above), a follow-up question arose: since `Median Salary` and `Median Salary Non-US` showed identical values when a single non-US country (India, Germany) was selected, does this mean `Median Salary US` is redundant and could be removed?

**Clarified: no — the equality was a special case specific to that test scenario, not a sign of redundancy.**

In the **default, unfiltered view** of the Q2 chart (no slicer applied — the normal way the chart will be viewed), all three measures show genuinely different values:
- `Median Salary` = $115,000 (all countries blended)
- `Median Salary US` = $118,940 (US only)
- `Median Salary Non-US` = $111,175 (non-US only)

This three-way difference is the entire point of the Q2 chart. The apparent equality seen during slicer testing only occurred because the slicer had already narrowed the whole table down to one specific non-US country — in that narrow context, "median for Germany" and "median for Germany, further restricted to Non-US" mathematically describe the identical set of rows, so of course the results matched. This is a coincidental consequence of that specific test filter, not evidence the measure does nothing in general.

**Also clarified:** a chart's legend always lists all series names the chart is built to display, regardless of whether a given series currently has data for the active filter — so seeing "3 items in the legend" does not by itself indicate whether 2 or 3 bars are actually rendering per category. Confirmed the reliable way to check this is clicking directly on an individual bar to see its exact series/value via tooltip, rather than judging bar count visually (especially from a photo, where two equal-height bars sitting adjacent can visually blend into what looks like one).

**Key takeaway:** always distinguish between a measure's behavior in its *default/most common* use case versus behavior observed under a specific, narrow test filter — a coincidental match under one test condition doesn't imply the measure is unnecessary in general use.

---

## Next: Stage 6 continued — add `job_title_short` slicer to Skill_Job_Analysis (Q3), stress-test (expect plain Count of job_id field to respond correctly by default, no CALCULATE/KEEPFILTERS issue anticipated since this field carries no hardcoded filter), then repeat for Skill_Salary_Analysis (Q4).

---
---

# Stage 6 continued — Slicers on Q3 and Q4, and a permanent fix for the recurring blank-skill nuisance

## Part E — Permanently removing "(blank)" skill rows at the source

**Problem:** the `(blank)` skill category (postings with no skills listed) kept resurfacing across different job-title slicer selections in Q3, and combining a manual "(blank) unchecked" filter with a **Top 10 Value Filter** proved unreliable — reapplying Top 10 would silently reset the manual blank exclusion, causing blank to reappear in the list. This whack-a-mole behavior (fix it, apply another filter, it comes back) was the recurring source of confusion.

**Root cause of the conflict:** a Value Filter (Top N) and a manual Label filter (unchecking a specific item) are two separate filter mechanisms in a PivotTable field, and applying one can silently override or reset the other rather than combining cleanly.

**Permanent fix — removed at the data source instead of the PivotTable layer:**
1. **Data** tab → **Queries & Connections** → double-clicked `data_job_skills` to reopen **Power Query Editor**.
2. Clicked the filter dropdown on the `job_skills` column header → unchecked **(blank)**/**(null)** → OK. This adds a permanent Applied Step.
3. **Home** → **Close & Load**.
4. Refreshed all PivotTables (**Data** → **Refresh All**) to pick up the change.

**Result:** blank skill rows are now excluded from the underlying `data_job_skills` table entirely — they cannot reappear in any PivotTable, regardless of which combination of slicers, Top N filters, or manual filters is subsequently applied, since the rows simply no longer exist in the source data feeding every PivotTable.

**Lesson:** when a "nuisance" category needs excluding and it's fighting with other filters at the PivotTable level, fixing it once at the Power Query/source level is more robust than repeatedly reapplying PivotTable-level filters that can conflict with each other.

## Part F — Added `job_title_short` slicer to Skill_Job_Analysis (Q3)

1. Clicked inside the `Skill_Job_Analysis` PivotTable → **PivotTable Analyze** → **Insert Slicer** → checked `job_title_short` → OK.
2. Reapplied **Value Filters → Top 10** (now reliable, since blank can no longer resurface after the Part E fix).

**Stress test — sample results across multiple job titles, verified for sensible role-specific variation:**
- **Data Analyst:** sql (5,033), excel (3,839), python (2,765), tableau (2,725) — Excel ranking unusually high specifically for this role (vs. lower for other roles) matches real-world expectations.
- **Cloud Engineer:** aws (29), python (38), sql/azure (22 each) — cloud/infrastructure skills correctly dominate over analyst-style tools for this role.
- **Business Analyst:** sql (466), excel (393), tableau (288) — Excel/Tableau-heavy profile, consistent with a reporting-focused role.

**Confirmed:** `Count of job_id` (a plain field, not a CALCULATE-based measure) responded correctly to the slicer by default, with no CALCULATE-replaces-filter conflict — consistent with the reasoning that only measures with an explicit `CALCULATE(..., column = value)` filter on the *same column* as an active slicer are at risk of the Part C/D issue found in Q2.

## Part G — Added `job_title_short` slicer to Skill_Salary_Analysis (Q4)

1. Clicked inside the `Skill_Salary_Analysis` PivotTable → **PivotTable Analyze** → **Insert Slicer** → checked `job_title_short` → OK.

**⚠️ Apparent issue during first test (resolved as a false alarm):** Selected "Data Scientist" and initially found `Median Salary by Skill` / `Skill Likelihood %` showing values **identical** to the unfiltered view (sql: $120,000/57% in both cases) — looked like a repeat of the Q2 frozen-measure bug.

**Investigation before assuming a bug:**
1. Confirmed via **Report Connections** that the slicer was correctly linked to the relevant PivotTable (it was — ruled out a connection issue).
2. Confirmed the underlying PivotTable *and* chart both updated (not just a chart-caching quirk) for several other job titles tested (Cloud Engineer, Data Analyst, Business Analyst) — all showed correctly varying, sensible values.
3. Re-selected "Data Scientist" explicitly and re-checked the SQL row directly: this time showed **$132,500 / 13%** — genuinely different from the unfiltered $120,000/57%.

**Conclusion:** the original "identical" result was a **screenshot/testing mix-up** (likely comparing against a stale or mismatched screenshot), not a genuine bug. Re-verified result makes logical sense: SQL is somewhat less central specifically to Data Scientist roles (who lean more heavily on Python/statistics/ML) compared to its dominance across the *entire* data job market (which includes SQL-heavy roles like Data Analyst and Business Analyst) — a lower likelihood % paired with similar-or-higher pay for a comparatively less-required skill is a believable pattern.

**Confirmed:** no `KEEPFILTERS` fix was needed for Q4's measures — `Median Salary by Skill`'s existing `CROSSFILTER(..., BOTH)` already handles bidirectional filtering by design, and `job_title_short` is a normal row-level field with no hardcoded CALCULATE filter to conflict with the slicer, unlike Q2's `job_country`-filtered measures.

## Key Lessons From This Part

1. **When two different PivotTable filter mechanisms (Value Filters vs. manual Label filters) conflict, fixing the underlying issue once in Power Query is more robust than repeatedly managing filter interactions at the PivotTable layer.**
2. **Not every "identical result" under a new test condition is a bug** — before assuming a measure is frozen/broken, rule out simpler explanations first: wrong PivotTable connection, stale/mismatched screenshots, or (as very rarely but possibly) a genuine coincidental match. Testing multiple different filter values (not just one) before concluding "it's broken" is what caught this was a false alarm rather than a real issue.
3. **Only measures with an explicit CALCULATE filter on the *same column* as an active slicer are at risk of the "filter replaces instead of combines" issue** — plain fields (like `Count of job_id`) and measures using CROSSFILTER on a *different* filtering mechanism (like `Median Salary by Skill`'s `job_id`-based CROSSFILTER, unrelated to the `job_title_short` slicer column) are not subject to the same risk, and testing confirmed this distinction held up correctly in practice.

**Stage 6 (Slicers) — Complete.** All three applicable sheets (Salary_Analysis, Skill_Job_Analysis, Skill_Salary_Analysis) now have working, rigorously-verified slicers.

---

# PROJECT STATUS: Fully complete

All four README questions built, charted, polished, and now enhanced with verified interactive slicers. The project includes a well-documented trail of genuine debugging, conceptual understanding (not just fixes), and independent verification against a flawed reference implementation throughout — strong foundation for a portfolio write-up.
