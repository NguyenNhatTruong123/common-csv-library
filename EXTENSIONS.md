# Extension proposal

The current module keeps one predictable calculation contract and exposes `TabularData`, `ColumnMapping`, and `OutputLabels` so callers can supply their own parsing and metadata. To grow it into a reusable multi-organization component:

1. **Introduce format adapters.** Keep parsing outside the arithmetic service and define adapters for CSV, XLSX, JSON, and other requested formats. Have each adapter produce the same normalized tabular model. Enforce upload size, row count, encoding, and formula/macro policies at the adapter boundary.
2. **Define a versioned schema.** Accept a JSON/YAML schema containing field role, source header aliases or index, requiredness, numeric format, unit, and output labels. Validate schema documents and reject ambiguous field matches. Keep the simple CSV method as a convenience API.
3. **Make locale and money policy explicit.** Add configurable decimal/grouping separators, currency code, precision, and rounding policy. Do not guess locale from a handful of headers for financial data. The current API assumes one currency and two decimal places.
4. **Separate computation from file output.** Return typed calculation results first, then add CSV/XLSX renderers. This enables callers to use results in services or other formats and makes arithmetic tests independent of spreadsheet libraries.
5. **Harden operational behavior.** Add limits for rows/columns/cell lengths, structured validation errors, logging hooks without sensitive row contents, and streaming implementations for large uploads.
6. **Release cleanly.** Choose a stable organization-owned group ID, semantic versions, license, changelog, dependency policy, and CI build. Publish only from a reviewed clean source revision to the organization's Maven repository; keep credentials and signing keys in CI secrets.
7. **Test interoperability.** Keep unit tests for validation/calculation and add adapter tests for encodings, quoted delimiters, large files, locale-specific numbers, and representative customer schemas. Add compatibility tests before changing the public API.
