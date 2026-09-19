# Owner Decision Log

[日本語](OWNER_DECISIONS.md) | English

Excel Batch Tool was not designed around the goal of automating every possible spreadsheet operation.

The project has repeatedly prioritized:

> **repeatable work, local processing, source-file safety, and clear refusal over silent corruption**

This document highlights the owner-level decisions behind that boundary.

---

## 1. Start with read-only safety analysis before building editing features

### Problem

The first risk in batch Excel automation was not "can this be automated?"

It was:

> **Can the tool change something it does not fully understand?**

Real workbooks may contain formulas, merged cells, drawings, charts, images, pivot tables,
external references, protection, validation, names, and other parts that simple cell-editing code can damage.

### Decision

Phase 0 was a **read-only analyzer**, not an editor.

Input files are opened with `FileAccess.Read`, so write access is blocked at the OS level.

Tests compare SHA-256, file size, and modification time before and after analysis.

The development order was deliberate:

> detect risk first, then add write features on top of that safety model.

**Evidence:** [Phase 0: read-only workbook safety analyzer](https://github.com/nikotaronosuke/excel-batch-tool/commit/da461404d10fdde32ec54c061bf50a8369107900)

---

## 2. Choose fully local Open XML processing instead of Excel automation or cloud processing

### Problem

The target use cases include repeated office work and crowdsourcing jobs.

That means the tool may be used where:

- Microsoft Excel is not available
- business data cannot be uploaded to a cloud service
- the user cannot maintain macros or scripts
- external AI APIs are inappropriate

### Decision

The product baseline became:

- Windows desktop
- no Excel requirement
- fully local
- no login
- no application network path
- Microsoft Open XML SDK as the main workbook engine

The central value is not "maximum API convenience."
It is the ability to inspect and change workbook packages while minimizing unintended rewrites.

**Evidence:** [v1 overview](spec/v1-overview.md) / [design decisions](decisions.md)

---

## 3. Block unsupported content instead of silently dropping it

### Problem

When aggregating worksheets into a new workbook, it is easy to copy only the parts the implementation understands.

That can produce a file that opens successfully while silently losing:

- formulas
- charts / drawings / images
- tables / pivot tables
- comments
- validation
- conditional formatting
- hyperlinks
- external references

A successful open is not enough if the workbook's meaning changed.

### Decision

If the current phase cannot safely preserve required content, the preview **blocks with a reason**.

The system does not copy the supported 80% and quietly discard the unsupported 20%.

This same principle continued as support expanded:
when a sheet contains a mixture of supported and unsupported rule variants, partial copying is generally rejected.

**Evidence:** [Phase 1B.1](https://github.com/nikotaronosuke/excel-batch-tool/commit/827f2e4c9ea8b0b0454408e30c7b2ee8fc0d6e95)

---

## 4. Never edit the source file in place

### Problem

Opening the original workbook for editing is simpler.

But if the process crashes, a file is locked, or validation fails mid-write,
the original data itself may be damaged.

### Decision

Write operations use this shape:

```text
source (read-only)
   ↓ byte copy
temporary file
   ↓ edit
re-open + validate
   ↓
new final output
```

Additional rules include:

- input path cannot be the output path
- existing output is not silently overwritten
- generated files are reopened and validated before final move
- rollback is best-effort and the UI must not claim cleanup succeeded unless it can verify that

The source workbook remains the safety anchor.

**Evidence:** [Phase 1A](https://github.com/nikotaronosuke/excel-batch-tool/commit/b846d359680cc3b5f8edf47118ff3ee0ee46ca24) / decisions D-023–D-024

---

## 5. Prefer an all-or-nothing batch over silent partial application

### Problem

Suppose 99 edits are safe and one is unsafe.

Applying the 99 safe edits appears convenient, but the user may believe the full request succeeded.

Partial success can create a result that is harder to audit than a clear failure.

### Decision

For batch cell updates:

> **one unsafe target blocks the entire batch**

Duplicate target cells also block rather than choosing the first or last value.

Pasted operation lists follow the same philosophy:
one invalid row rejects the batch instead of quietly adding the rest.

The goal is to make the result of an operation unambiguous.

**Evidence:** [Phase 2B](https://github.com/nikotaronosuke/excel-batch-tool/commit/cba1b9ac0bd7a18fa039dcf738a560f0f5ab0e51)

---

## 6. Do not perform "smart" key normalization in business data

### Problem

It is tempting to automatically treat these as equivalent:

- `00123` and `123`
- `ABC` and `abc`
- values with surrounding spaces
- full-width and half-width forms

But business identifiers can be intentionally distinct.

### Decision

Key matching uses **exact strings**.

The tool does not silently:

- trim
- case-fold
- normalize width
- normalize Unicode
- coerce strings to numbers
- repair zero-padding

Zero matches or multiple matches are surfaced instead of guessed.

The project prefers a visible mismatch over writing data to the wrong row.

**Evidence:** [Phase 2C1](https://github.com/nikotaronosuke/excel-batch-tool/commit/265631084e781d7c02e4d5b55f76676c75d6606e) / [Phase 2C2](https://github.com/nikotaronosuke/excel-batch-tool/commit/72ffe33ab5181e7114af74d3a74257b020bb4687)

---

## 7. Let real job patterns change the roadmap

### Problem

The roadmap could have expanded by adding generally useful spreadsheet features in an arbitrary order.

Instead, crowdsourcing-market research showed recurring demand around:

- multi-cell repetitive entry
- source-table to form transfer
- key-based table updates
- system-to-system CSV transformation
- repeated monthly workbook operations

### Decision

Implementation order was changed to follow those recurring tasks.

That moved features such as:

- input sets
- table-to-cell mapping
- table-to-table updates
- CSV transformation

ahead of generic search / replace.

VBA editing was also not pulled into the product simply because VBA jobs exist.

The decision criterion was:

> does this fit the product's safe, repeatable workflow — not merely "can it be implemented?"

**Evidence:** [market research summary](research/market-summary.md) / [Phase 2E CSV transform](https://github.com/nikotaronosuke/excel-batch-tool/commit/ce15518b09babd84589b9de92886f6f7592b1529)

---

## 8. Change validation scope based on what the file is used for

### Problem

"Maximum safety" can sound like "run the full validator on everything."

But measured on a 100k-row source sheet:

- full workbook validation: about **26 s / 1.4 GB**
- streaming scan: about **5.4 s / 1 MB**

The source file is only being read.
The target file is being rewritten.

Those operations do not have the same risk model.

### Decision

For read-only transfer sources, validation focuses on the parts needed to interpret values safely,
and unsupported cells block during reading.

For workbooks that will be rewritten, full validation remains.

This separates:

> **what is being protected, from what failure**

instead of applying the same expensive check everywhere.

**Evidence:** [Phase 2C1](https://github.com/nikotaronosuke/excel-batch-tool/commit/265631084e781d7c02e4d5b55f76676c75d6606e)

---

## 9. Add PDF / OCR only after a delivery-oriented feasibility benchmark

### Problem

"Text can be extracted from this PDF" is a weak success criterion for client work.

A useful PDF-to-Excel tool needs to preserve fields and cells accurately enough to produce a deliverable result.

### Decision

PDF / OCR work started with a separate benchmark harness.

It measured generated fictional fixtures across:

- born-digital text
- born-digital tables
- clean scans
- degraded scans
- fixed forms
- scanned tables
- checkboxes / marks
- handwriting-like characterization

Ground truth was created together with fixtures and compared mechanically.

Synthetic handwriting-like fixtures reaching high accuracy were **not** treated as proof that real handwriting was solved.
Real handwriting remained out of scope.

Product integration followed only after the benchmark supported a conditional GO decision.

**Evidence:** [Phase 2F-R](https://github.com/nikotaronosuke/excel-batch-tool/commit/3994ab4f7dbce92ac2f531bf886b4a3f361c5d99)

---

## 10. Prefer zero false auto-accepts over a higher automatic acceptance rate

### Problem

OCR confidence can be high and still wrong.

Development found cases where two recognizers agreed on the same incorrect value with high confidence.

Examples included code / shape confusions and upside-down table content.

Raising a confidence threshold alone did not solve that class of error.

### Decision

OCR output is divided into:

- auto-accepted
- needs review
- unreadable / not found

Additional structural gates consider things such as:

- field type
- character-shape pattern
- column shape consistency
- suspicious upside-down numeric forms
- merged / oversized detected bands

These guards do not rewrite a recognized value into a guessed correction.
They only decide whether automation is allowed.

Across the final fictional job-like fixture set, the target remained:

> **false auto-accept = 0**

even when that meant more human review.

**Evidence:** [Phase 2F-B3](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)

---

## 11. Reject a newer OCR model when the measured trade-off was worse

### Problem

A newer OCR model looked like an obvious upgrade candidate.

### Measurement

On the same product path, the newer configuration produced almost the same exact-match result,
while taking approximately **2.4–2.9×** as long and increasing the pack size.

The major accuracy gain came from the runtime update, not from replacing the recognition model.

### Decision

**PP-OCRv5 was not adopted.**

"Newer" was not treated as a product benefit on its own.

The decision was based on:

> useful accuracy improvement relative to latency / distribution cost

not model version number.

**Evidence:** [Phase 2F-B3](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)

---

## 12. Validate the final result, not only an intermediate metric

### Problem

Deskew angle estimation looked accurate.

That suggested deskew itself was working.

In reality, the correction direction was reversed.

The system had measured roughly the right skew angle but rotated the page the wrong way.

### Decision

The final corrected table / form result became part of the acceptance check.

After fixing the direction, measured tilted-fixture results improved substantially while false auto-accepts returned to zero.

The lesson was simple:

> an intermediate metric being correct does not prove the product output is correct.

**Evidence:** [Phase 2F-B3 deskew correction](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)

---

## What this project prioritizes

Excel Batch Tool prioritizes:

- never editing source workbooks in place
- refusing unsupported semantics instead of silently dropping them
- clear batch failure over ambiguous partial success
- exact keys over guessed identity
- real job demand over feature-list expansion
- validation that matches the actual risk of the operation
- exact-match / false-accept measurement for OCR
- human review as a safety mechanism, not a failure
- measured trade-offs over "newer technology"
- final output validation over reassuring intermediate metrics

AI-assisted implementation is part of the development process.

The part I wanted this repository to preserve is **where automation deliberately stops**.
