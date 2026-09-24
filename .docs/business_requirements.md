# Common CSV Invoice Calculator Requirements

## 1. Scope

Java library for parsing supported tabular input, validating required invoice fields, calculating per-item totals before and after VAT, and returning an XLSX workbook containing the original data and calculations.

The primary convenience flow accepts CSV. Callers that read other formats can adapt the data to the library's `TabularData` model and provide a `ColumnMapping`. This repository does not include readers for arbitrary file formats or a user interface.

## 2. Technology baseline

- Java 17 or later; Maven.
- Maven coordinates: `io.github.commoncsv:common-csv-library:1.0.0-SNAPSHOT`.
- Root package: `io.github.commoncsv`.
- Apache Commons CSV for delimited input parsing.
- Apache POI OOXML for XLSX output.
- JUnit 5 for unit tests.
- Public API: `CsvInvoiceCalculator`, `TabularData`, `ColumnMapping`, and `OutputLabels`.
- No web server, persistence layer, CLI application, or UI is included in the library.

## 3. Input data contract

Every input table has a header row and zero or more product rows. Required fields are:

| Field | Required | Rules |
| --- | ---: | --- |
| Product name | Yes | Non-blank text |
| Quantity | Yes | Numeric and greater than zero |
| Unit price | Yes | Numeric and zero or greater; before tax |
| VAT percent | Yes | Numeric percentage from 0 through 100 inclusive; `%` suffix is accepted |

Other columns are optional and are copied to the result without interpretation. All unit prices are assumed to use the same currency. The library does not convert currency or infer units.

### CSV convenience input

- UTF-8 text; an initial UTF-8 BOM is accepted.
- Comma, semicolon, and tab delimiters are supported.
- Standard quoted fields are supported, including quoted values containing delimiters.
- Header names are matched case-insensitively and normalized for accents and punctuation.
- Common header aliases are detected in English, Vietnamese, Japanese, and Korean. Detection is a convenience feature, not a universal translation service; callers should pass explicit `ColumnMapping` and `OutputLabels` for other names or ambiguous schemas.

| Field | Current common aliases |
| --- | --- |
| Product name | `Product`, `Product Name`, `Item`, `Item Name`, `Name`; Vietnamese `Tên mặt hàng`, `Tên sản phẩm`, `Sản phẩm`; Japanese `商品名`, `品名`; Korean `제품명`, `상품명`, `품명`, `품목명` |
| Quantity | `Quantity`, `Qty`; Vietnamese `Số lượng`; Japanese `数量`; Korean `수량` |
| Unit price | `Unit Price`, `Price per Unit`, `Price`; Vietnamese `Đơn giá`; Japanese `単価`; Korean `단가`, `개당 가격`, `단위 가격` |
| VAT percent | `VAT`, `VAT %`, `VAT Percent`, `Tax`, `Tax Rate`; Vietnamese `Thuế`, `Thuế suất`; Japanese `税率`; Korean `부가세`, `부가가치세`, `부가세율`, `세율` |

Other input formats must be parsed by the caller into `TabularData`, which requires a non-empty header list and rows with the same width. `ColumnMapping` uses zero-based indexes and requires four distinct, non-negative indexes.

## 4. Functional requirements

- Calculate each item's net total as `quantity × unit price`.
- Calculate each item's gross total as `net total × (1 + VAT percent / 100)`.
- Round each item's net and gross totals to two decimal places using `HALF_UP`.
- Grand totals sum the already-rounded item totals.
- Return an XLSX workbook as a `byte[]`.
- Preserve all input headers and source cell text in the result workbook.
- Append two calculated columns for each item's before-tax and after-tax totals.
- Add a final row containing grand totals before and after tax.
- Use numeric XLSX cells with two-decimal display formatting for calculated amounts.
- Use localized labels for the calculated columns and grand totals. Built-in output labels are available in English, Vietnamese, Japanese, and Korean. Callers can provide custom labels through `OutputLabels`.
- For CSV, infer output label language from detected Vietnamese, Japanese, or Korean header characters; otherwise use English. For predictable localization, callers should use `process(TabularData, ColumnMapping, OutputLabels)` with explicit labels.
- An input with headers but no product rows is valid and produces a workbook with the headers and zero grand totals.

## 5. Validation and quality requirements

- Reject empty CSV input and CSV files missing any required header.
- Reject empty product names, missing numeric values, malformed numeric strings, non-positive quantity, negative unit price, and VAT outside the allowed range.
- Accept non-negative integer or decimal numeric input; decimal point is supported, and decimal comma is supported when it remains within a single parsed CSV cell. In comma-delimited CSV, quote a decimal-comma value such as `"1,25"`.
- Accept conventional thousands grouping such as `1,234.56`.
- Reject rows whose number of cells differs from the header width.
- Fail fast with an `IllegalArgumentException` that identifies the CSV row and field for invalid row values.
- Reject invalid mapping indexes and duplicate required-field mappings.
- Unit tests should cover calculations, rounding, CSV parsing, multilingual headers and labels, preservation of extra columns, and validation failures. Tests are in `src/test/java` and can be run with `mvn clean test`.
- Build and test execution are performed by the project user or CI; this requirements document does not imply publishing or deployment.

## 6. Non-goals

Web UI, REST service, arbitrary-format file readers, universal language detection, locale-specific currency conversion, configurable currency precision, multiple currencies in one file, exchange-rate lookup, formula evaluation, macro preservation, streaming large-file processing, database storage, authentication, deployment, and publication to a remote Maven repository are outside the current library scope.
