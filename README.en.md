# Excel Batch Tool

English | [日本語](README.md)

> **Excel Batch Tool is a development codename. The final product name has not been decided.**

A Windows desktop app for **safe, fully local batch processing of Excel workbooks** without requiring Microsoft Excel itself.

The project is still in early development. It is not presented as a finished stable product.

<p align="center">
  <img src="docs/images/app-overview.webp" width="900" alt="Excel Batch Tool showing workbook analysis, safety checks, and eight processing tabs" />
</p>

<p align="center"><em>Real application screen — workbook safety analysis, batch processing, CSV conversion, and PDF/OCR live in one desktop UI.</em></p>

## Core idea

The main value is not "support every Excel feature."

It is:

- process many workbooks repeatedly
- work without Excel installed
- stay fully offline
- preview before changing anything
- never modify the source file
- refuse cases the tool cannot preserve safely
- validate generated output before finalizing it
- keep an audit trail for changes

## Current feature set

### 1. Read-only workbook analysis

Inspect multiple `.xlsx` files before doing any write operation.

The analyzer detects workbook / worksheet characteristics that may make later editing unsafe, including:

- formulas
- merged cells
- drawings / charts / images
- pivot tables
- external references
- protection
- tables
- data validation
- conditional formatting
- comments
- defined names
- hyperlinks
- OLE / ActiveX
- macro-related content

Results are classified as normal / needs attention / unsupported.

Input files are opened read-only and tests verify that SHA-256, size, and modification time are unchanged after analysis.

### 2. Vertical table merge

Combine similarly shaped tables from multiple workbooks into one **new** workbook.

Features include:

- header-based column matching even when column order differs
- explicit base-sheet selection for output column order
- preview of row count, columns, and blocking problems
- optional source-file / source-sheet metadata columns
- no overwrite of existing files

Formula-bearing or ambiguous table structures block rather than being silently flattened.

### 3. Worksheet aggregation

Collect selected sheets from multiple workbooks into a new workbook while preserving a deliberately limited set of spreadsheet semantics.

Supported content is copied only when its meaning can be preserved.
Unsupported elements block in preview instead of silently disappearing.

### 4. Batch cell updates

Apply a set of fixed text / numeric / blank values to selected sheets across many workbooks.

Important safety rules:

- source workbooks remain unchanged
- output uses a new filename
- duplicate target cells block
- one unsafe target blocks the whole batch
- only intended worksheet parts are changed
- an `.audit.json` sidecar records the actual changes

### 5. Table row → fixed-cell mapping

Read one row from an Excel / CSV table by exact key match and write selected fields into fixed target cells.

This is useful for cases such as:

- store-by-store forms
- monthly report templates
- one source table feeding many individual files

Keys are matched exactly. The tool does not silently trim, case-fold, normalize width, coerce numbers, or repair zero-padding.

### 6. Table-to-table matching and update

Match source and target table rows by key and update selected columns in existing target rows.

Only the intersection is updated:

- source-only keys do not create rows
- target-only keys are not deleted
- ambiguous / duplicate keys block when they would affect the update
- the key column itself cannot be changed by the same operation

### 7. CSV transformation

Transform Excel / CSV input into a new CSV shape required by another system.

You can:

- rename output columns
- reorder columns
- omit internal columns
- duplicate a source column into multiple output columns
- add fixed-value columns
- choose UTF-8 with / without BOM or Shift_JIS
- choose normal quoting or quote-every-field

The generated CSV is read back and checked before it is finalized.

### 8. PDF → Excel / CSV

Born-digital PDFs can be parsed into structured text / table output.

The project also supports an optional **Offline OCR Pack** for scanned PDFs.

The OCR flow distinguishes:

- auto-accepted
- needs review
- unreadable / not found

Suspicious OCR is not silently rewritten into a guessed value.

The original PDF page is shown during review so the user can compare the source and recognition result before export.

## OCR philosophy

The OCR feature was added only after a separate feasibility benchmark measured deliverable-quality exact matches.

The project intentionally values:

> **zero false auto-accepts over a higher automatic acceptance rate**

Examples observed during development included both recognizers agreeing on a wrong high-confidence result.
Because of that, agreement and confidence alone are not considered sufficient evidence.

The current system adds structural / field-shape guards and sends uncertain cases to human review.

Handwriting remains outside the supported scope because synthetic handwriting-like fixtures cannot justify a claim about real handwriting.

## Safe write pipeline

Write operations follow the same basic shape:

```text
source file (read-only)
        ↓ byte copy
temporary output
        ↓ edit intended parts only
re-open + validate
        ↓
final output (new file)
```

The source workbook is never opened for in-place editing.

Existing output paths are not silently overwritten.

Rollback is best-effort and the UI does not claim files were removed unless it can verify that result.

## All-or-nothing behavior

For batch edits, the project generally prefers:

> one unsafe operation blocks the batch

over:

> apply the 99 safe rows and silently skip the 1 unsafe row

Partial success can be harder to understand than a clear stop.
The user should know exactly whether the requested batch was applied.

## Processing recipes

Frequently repeated settings can be saved locally as recipes.

Recipes can store things such as:

- key column choice
- header row
- source-to-target mappings
- output suffix
- PDF reading mode
- fixed-form field definitions

Recipes do **not** store the source Excel / CSV / PDF file itself.

For PDF recipes, the original file name, path, recognized text, and page contents are not stored.

Loading a recipe never auto-executes it; preview and safety checks must run again for the current files.

Recipes are stored locally under the user's application data directory and are not cloud-synced.

## Offline design

- Windows only
- .NET 8 / WPF
- Microsoft Excel is not required
- Office Interop is not used
- workbook processing is based on Microsoft Open XML SDK
- no login
- no cloud service
- no external AI API
- the application itself has no network communication path

The primary workbook format is `.xlsx`; legacy `.xls` is out of scope.

## Build

.NET 8 SDK is required for development:

```bash
dotnet build
dotnet test
dotnet run --project src/ExcelBatchTool.App
```

### Self-contained Windows build

```bash
dotnet publish src/ExcelBatchTool.App -c Release -p:PublishProfile=win-x64-self-contained
```

The self-contained publish does not require the user to separately install the .NET Desktop Runtime.

## Offline OCR Pack

The scanned-PDF OCR runtime is built and distributed separately from the main app.

The pack contains its models, native inference runtime, app-local runtime dependencies,
license notices, and a manifest with file sizes / SHA-256 hashes.

The finished pack is self-contained and does not download models at runtime.

The main app verifies the pack manifest before loading native OCR components.
If the pack is incomplete or mismatched, it stops with a user-facing reason instead of entering native loading with an unknown state.

The main app can run normally without the OCR Pack; born-digital PDF support does not depend on it.

## Repository layout

```text
src/ExcelBatchTool.Core          core analysis and batch-processing logic
src/ExcelBatchTool.App           WPF desktop application
src/ExcelBatchTool.Ocr           contents used to assemble the Offline OCR Pack
tools/OcrPackBuilder             OCR Pack assembly tool
tests/                           generated fictional workbook tests
docs/                            specification, roadmap, decisions, research, implementation notes
research/PdfFeasibility          benchmark code used before product OCR integration
```

## Validation

The test suite uses fictional / generated documents.

It verifies areas such as:

- source-file immutability
- output validation
- stale-preview / snapshot guards
- exact-key behavior
- partial-application prevention
- package-part integrity
- CSV round-trip correctness
- OCR review / auto-accept boundaries
- PDF extraction behavior

The project also records known limitations instead of silently claiming complete Excel compatibility.

## Known boundaries

This is **not** a universal Excel automation engine.

Examples of intentionally unsupported or constrained areas include:

- legacy `.xls`
- VBA editing
- in-place overwrite
- arbitrary formula-aware editing
- every possible conditional-format / data-validation variant
- severe perspective distortion in photographed documents
- real handwriting
- arbitrary free-layout document understanding

Unsupported cases should be detected and surfaced rather than silently damaged.

## Documentation

The full Japanese documentation contains the detailed feature-level specification and benchmark history:

- [Japanese README](README.md)
- [Decision history](docs/decisions.md) *(Japanese, append-only; migrated into this repository on 2026-09-19 — see D-038)*
- [Documentation index](docs/README.md) *(Japanese)*
- [PDF feasibility benchmark](docs/research/pdf-feasibility-research.md) *(Japanese)*

## AI-assisted development

The project was developed with AI-assisted coding and research.

## License

MIT. See [LICENSE](LICENSE).