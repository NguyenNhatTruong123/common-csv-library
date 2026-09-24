# Common CSV invoice calculator

A small Java 17+ library that validates tabular invoice data, calculates each item's total before and after VAT, and returns an Excel `.xlsx` workbook with the original columns, calculated columns, and grand totals. XLSX is used as the output because it keeps text and numeric cells separate and supports Unicode labels and multiple columns without delimiter/quoting ambiguity.

## Requirements

- JDK 17 or later
- Maven 3.8+

## Quick use: CSV with recognizable headers

The convenience method accepts UTF-8 CSV input (including UTF-8 BOM) separated by comma, semicolon, or tab. It detects common English, Vietnamese, and Japanese names for the required columns and returns XLSX bytes.

```java
import io.github.commoncsv.CsvInvoiceCalculator;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

try (InputStream input = Files.newInputStream(Path.of("invoice.csv"))) {
    byte[] xlsx = CsvInvoiceCalculator.processCsv(input);
    Files.write(Path.of("invoice-results.xlsx"), xlsx);
}
```

For an HTTP endpoint, return the byte array with content type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` and a filename ending in `.xlsx`.

## Explicit metadata and other input formats

When headers use other languages, are ambiguous, or metadata identifies fields by another rule, parse the source into `TabularData`, then provide zero-based column indexes and output labels. This entry point can be used by an XLSX, JSON, database, or other format adapter without changing the calculation library.

```java
import io.github.commoncsv.*;
import java.util.List;

var table = new TabularData(
    List.of("商品", "数", "価格", "税率", "備考"),
    List.of(List.of("ペン", "2", "120", "10%", "青")));
var columns = new ColumnMapping(0, 1, 2, 3); // item name, quantity, unit price, VAT percent
var labels = OutputLabels.japanese();
byte[] xlsx = CsvInvoiceCalculator.process(table, columns, labels);
```

`OutputLabels.english()`, `.vietnamese()`, and `.japanese()` are included. Supply a custom `new OutputLabels(...)` for any other language or customer terminology. Input columns are copied as written; only the two appended calculation headers and the grand-total labels use `OutputLabels`.

## Calculation and validation rules

- Net item total = quantity × unit price.
- Gross item total = net item total × (1 + VAT percent / 100).
- Each item total is rounded to two decimal places using `HALF_UP`; grand totals add the rounded item totals.
- Quantity must be a positive number. Unit price must be zero or greater. VAT must be between 0 and 100 inclusive.
- Product name and all numeric fields are required. Numeric fields accept digits and an optional decimal point or decimal comma (e.g. `1.25` or `1,25`); VAT may also end in `%`. Thousands grouping is supported in the conventional `1,234.56` form.
- Invalid rows fail with `IllegalArgumentException` identifying the CSV row number (header is row 1) and field. Uneven row widths and invalid metadata indexes also fail fast.
- All prices are assumed to use the same currency and unit price is before tax. The library does not perform currency conversion or infer units.

## Develop and distribute

- [TESTING.md](GUIDELINE/TESTING.md): run the included tests, inspect reports, and add your own test cases.
- [PACKAGING.md](GUIDELINE/PACKAGING.md): build the JAR and install it into a local Maven repository for use in another project.

## Extension proposal

See [EXTENSIONS.md](EXTENSIONS.md) for a practical path toward organization-wide reuse, schema metadata, more file formats, currencies, and localization.
