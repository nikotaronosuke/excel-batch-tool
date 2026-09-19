# Owner Decision Log

[日本語](OWNER_DECISIONS.md) | English

Five decisions are worth keeping because measurement or failure changed the implementation direction.

## 1. Built a read-only analyzer before any editing feature

The first risk was not "can this be automated?" but "can automation damage the source workbook?"

Phase 0 started with a safety analyzer opened via `FileAccess.Read`. Tests verify SHA-256, file size, and mtime remain unchanged.

**Evidence:** [Phase 0](https://github.com/nikotaronosuke/excel-batch-tool/commit/da461404d10fdde32ec54c061bf50a8369107900)

## 2. Reordered the roadmap after studying real jobs

CrowdWorks / Lancers job patterns pushed the roadmap toward:

- repeated fixed-cell entry
- table-to-form transfer
- key-based matching
- CSV transformation

CSV transformation moved ahead of generic search / replace. VBA editing was not pulled into the product just because VBA jobs existed.

**Evidence:** [market research](research/market-summary.md) / [CSV transform](https://github.com/nikotaronosuke/excel-batch-tool/commit/ce15518b09babd84589b9de92886f6f7592b1529)

## 3. Changed validation strategy after measuring a 100k-row source

Measured on a large source:

- full validator: **~26 s / ~1.4 GB**
- streaming scan: **~5.4 s / ~1 MB**

Read-only sources and rewritten targets no longer use the same validation strategy. Sources use streaming checks for the semantics needed; targets keep full validation.

**Evidence:** [Phase 2C1](https://github.com/nikotaronosuke/excel-batch-tool/commit/265631084e781d7c02e4d5b55f76676c75d6606e)

## 4. Measured PP-OCRv5 and rejected it

Comparing the existing v4 dual-read path with v5 + japan v4 showed almost no meaningful difference across 480 fields, while the v5 configuration was **2.4–2.9× slower**.

The newer model was not adopted.

**Evidence:** [Phase 2F-B3](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)

## 5. Found that deskew angle estimation was right while the correction direction was wrong

The measured skew angle looked correct, so the intermediate metric suggested the feature worked.

The final image showed the rotation direction was reversed and the table grid was damaged.

The tests were changed to validate the corrected output itself, not only the angle estimate.

**Evidence:** [Phase 2F-B3 deskew fix](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)
