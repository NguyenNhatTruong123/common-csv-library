# Testing the Library

## Try a real CSV file and inspect the output

This section runs the library as a user would: provide a CSV file, run a small Java program, then open the generated XLSX workbook to review its data, labels, and layout. The sample files live in `manual-test/`, outside `target/`, so you can edit the CSV and run it again.

### 1. Create the input CSV

Create `manual-test/invoice.csv` with the following UTF-8 content. Save it as a comma-separated CSV file:

```csv
SKU,Product,Category,Brand,Quantity,Unit Price,VAT,Currency,Supplier,Warehouse,Order Date,Batch,Note
SKU-1001,Notebook A5,Stationery,PaperCo,3,2.50,10%,USD,SUP-01,WH-NORTH,2026-09-01,B-260901,School supplies
SKU-1002,Blue pen,Stationery,WriteWell,2,1.25,5%,USD,SUP-02,WH-NORTH,2026-09-01,B-260901,Blue ink
SKU-1003,Desk lamp,Electronics,BrightHome,1,18.99,8%,USD,SUP-03,WH-CENTRAL,2026-09-02,B-260902,"LED, warm white"
SKU-1004,Coffee mug,Kitchen,Ceramica,4,6.75,10%,USD,SUP-04,WH-SOUTH,2026-09-02,B-260902,350 ml
SKU-1005,USB cable,Electronics,ConnectX,5,3.40,8%,USD,SUP-03,WH-CENTRAL,2026-09-03,B-260903,1 meter
SKU-1006,Water bottle,Outdoor,TrailGear,2,12.00,5%,USD,SUP-05,WH-SOUTH,2026-09-03,B-260903,Steel
SKU-1007,Sticky notes,Stationery,PaperCo,10,0.85,10%,USD,SUP-01,WH-NORTH,2026-09-04,B-260904,Pastel colors
SKU-1008,Wireless mouse,Electronics,ConnectX,1,24.50,8%,USD,SUP-03,WH-CENTRAL,2026-09-04,B-260904,Bluetooth
SKU-1009,Cotton towel,Home,SoftLiving,3,8.25,5%,USD,SUP-06,WH-SOUTH,2026-09-05,B-260905,Large
SKU-1010,Lunch box,Kitchen,DailyUse,2,9.99,10%,USD,SUP-04,WH-SOUTH,2026-09-05,B-260905,BPA-free
SKU-1011,Power bank,Electronics,ConnectX,1,32.00,8%,USD,SUP-03,WH-CENTRAL,2026-09-06,B-260906,10000 mAh
SKU-1012,Pencil set,Stationery,WriteWell,6,2.10,5%,USD,SUP-02,WH-NORTH,2026-09-06,B-260906,"HB, 12 pieces"
SKU-1013,Plant pot,Garden,GreenCorner,2,7.50,10%,USD,SUP-07,WH-SOUTH,2026-09-07,B-260907,Ceramic
SKU-1014,Phone stand,Accessories,ConnectX,3,5.25,8%,USD,SUP-03,WH-CENTRAL,2026-09-07,B-260907,Adjustable
SKU-1015,Tea tin,Kitchen,DailyUse,1,14.75,5%,USD,SUP-04,WH-SOUTH,2026-09-08,B-260908,Earl Grey
SKU-1016,File folder,Stationery,PaperCo,8,1.15,10%,USD,SUP-01,WH-NORTH,2026-09-08,B-260908,A4
SKU-1017,Bluetooth speaker,Electronics,BrightHome,1,45.00,8%,USD,SUP-03,WH-CENTRAL,2026-09-09,B-260909,Compact
SKU-1018,Sunglasses,Outdoor,TrailGear,2,16.50,5%,USD,SUP-05,WH-SOUTH,2026-09-09,B-260909,UV400
SKU-1019,Hand soap,Home,SoftLiving,5,2.35,10%,USD,SUP-06,WH-SOUTH,2026-09-10,B-260910,Citrus
SKU-1020,Keyboard,Electronics,ConnectX,1,39.99,8%,USD,SUP-03,WH-CENTRAL,2026-09-10,B-260910,US layout
SKU-1021,Reusable bag,Home,DailyUse,4,3.00,5%,USD,SUP-06,WH-SOUTH,2026-09-11,B-260911,Canvas
SKU-1022,Marker set,Stationery,WriteWell,3,4.60,10%,USD,SUP-02,WH-NORTH,2026-09-11,B-260911,8 colors
SKU-1023,Desk organizer,Office,PaperCo,1,11.20,10%,USD,SUP-01,WH-NORTH,2026-09-12,B-260912,Bamboo
SKU-1024,Travel adapter,Electronics,ConnectX,2,13.45,8%,USD,SUP-03,WH-CENTRAL,2026-09-12,B-260912,International
```

You can replace this data with your own. The convenience method recognizes common English, Vietnamese, and Japanese header names. If your headers use a different language or naming scheme, see the metadata example below.

### Optional: Ask an AI assistant to create the manual test files

Instead of creating the files by hand, give an AI coding assistant access to the project and ask it to create the manual test setup. Review the generated files before running them. This prompt asks the assistant to create only files under `manual-test/` and not change the library implementation:

```text
In this Java 17+ Maven repository, create a runnable end-to-end manual test setup under the manual-test/ directory for the CSV invoice calculator library.

Create these files:
1. manual-test/pom.xml: a small Maven project that depends on io.github.commoncsv:common-csv-library:1.0.0-SNAPSHOT and uses exec-maven-plugin 3.3.0. Configure ManualCsvRun as the main class and configure manual-test/src as the source directory.
2. manual-test/src/ManualCsvRun.java: a Java main class that calls CsvInvoiceCalculator.processCsv(InputStream), reads manual-test/invoice.csv when launched from the repository root or invoice.csv when launched from manual-test/, and writes invoice-results.xlsx next to the input CSV. If neither input path exists, print a helpful error showing both paths checked.
3. manual-test/invoice.csv: a valid UTF-8 comma-separated CSV with at least 20 product rows and several extra informational columns (for example SKU, category, brand, currency, supplier, warehouse, order date, batch, and note), in addition to recognizable Product, Quantity, Unit Price, and VAT headers. Include at least one correctly quoted field containing a comma.

Do not edit the library's production code, unit tests, README, or other files. Make sure the paths and Maven source-directory configuration match the files you create. Summarize the files created and give the exact Maven command to run from the repository root. Do not run build or tests unless I ask.
```

You can compare the generated files with the examples below and review the paths, headers, row count, and expected totals before running them.

### 2. Create the runner program

Create `manual-test/src/ManualCsvRun.java`:

```java
import io.github.commoncsv.CsvInvoiceCalculator;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

public class ManualCsvRun {
    public static void main(String[] args) throws Exception {
        Path inputPath = Path.of("manual-test", "invoice.csv");
        if (!Files.exists(inputPath)) {
            inputPath = Path.of("invoice.csv");
        }
        if (!Files.exists(inputPath)) {
            throw new java.nio.file.NoSuchFileException(
                "CSV not found. Looked for manual-test/invoice.csv and invoice.csv from "
                    + Path.of("").toAbsolutePath());
        }
        Path outputPath = inputPath.resolveSibling("invoice-results.xlsx");

        try (InputStream input = Files.newInputStream(inputPath)) {
            byte[] result = CsvInvoiceCalculator.processCsv(input);
            Files.write(outputPath, result);
        }
        System.out.println("Created: " + outputPath.toAbsolutePath());
    }
}
```

### 3. Build the library and run the program

Open a terminal in the library's root directory (the directory containing its `pom.xml`) and install the library to your local Maven repository:

```shell
mvn clean install
```

Create `manual-test/pom.xml` so the runner can use the library and its dependencies:

```xml
<project xmlns="http://maven.apache.org/POM/4.0.0"
         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
         xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 https://maven.apache.org/xsd/maven-4.0.0.xsd">
  <modelVersion>4.0.0</modelVersion>
  <groupId>manual.test</groupId>
  <artifactId>manual-csv-run</artifactId>
  <version>1.0-SNAPSHOT</version>
  <properties>
    <maven.compiler.release>17</maven.compiler.release>
    <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
  </properties>
  <dependencies>
    <dependency>
      <groupId>io.github.commoncsv</groupId>
      <artifactId>common-csv-library</artifactId>
      <version>1.0.0-SNAPSHOT</version>
    </dependency>
  </dependencies>
  <build>
    <plugins>
      <plugin>
        <groupId>org.codehaus.mojo</groupId>
        <artifactId>exec-maven-plugin</artifactId>
        <version>3.3.0</version>
        <configuration>
          <mainClass>ManualCsvRun</mainClass>
        </configuration>
      </plugin>
    </plugins>
    <sourceDirectory>${project.basedir}/src</sourceDirectory>
  </build>
</project>
```

Run the program from the library root. No `-Dexec.mainClass` argument is needed, avoiding shell-specific parsing issues:

```shell
mvn -f manual-test/pom.xml compile exec:java
```

On success, Maven prints the output path. Open `manual-test/invoice-results.xlsx` with Excel, LibreOffice Calc, or another XLSX-compatible application.

### 4. Review the output

For the sample CSV, the workbook should contain 24 product rows and 15 columns: the original 13 input columns plus `Total before tax` and `Total after tax`. The final grand-total row should show `467.66` before tax and `503.86` after tax. Check that `SKU`, `Category`, `Brand`, `Currency`, `Supplier`, `Warehouse`, `Order Date`, `Batch`, and `Note` are still present. Some notes contain quoted commas in the CSV; each should remain in a single cell without shifting columns.

To check validation, change a quantity to `0`, a price to `abc`, or VAT to `150`. The program should report the CSV row and invalid field, and should not produce a successful workbook.

## Use headers the convenience method cannot detect

If the CSV uses another language or header names, parse it in a way that fits the file and call `process(TabularData, ColumnMapping, OutputLabels)` with explicit metadata. For a quick check, replace the body of `main` in the runner above with:

```java
var table = new io.github.commoncsv.TabularData(
    java.util.List.of("Articulo", "Unidades", "Precio", "Impuesto"),
    java.util.List.of(java.util.List.of("Cuaderno", "3", "2.50", "10%")));
byte[] result = io.github.commoncsv.CsvInvoiceCalculator.process(
    table,
    new io.github.commoncsv.ColumnMapping(0, 1, 2, 3),
    new io.github.commoncsv.OutputLabels("Total sin impuesto", "Total con impuesto", "Suma sin impuesto", "Suma con impuesto"));
java.nio.file.Files.write(java.nio.file.Path.of("manual-test/invoice-results.xlsx"), result);
```

After changing the library source, run `mvn clean install` again before rerunning the manual project so it picks up the latest local build.

## Prerequisites

Install JDK 17+ and Maven 3.8+, then open a terminal in the directory containing `pom.xml`. Confirm the tools are available:

```shell
java -version
mvn -version
```

## Run the included unit tests

Run all tests and compile the project:

```shell
mvn clean test
```

Maven reports the result in the terminal. Detailed per-test reports are written to `target/surefire-reports/` as `.txt` and `.xml` files. To run one test class:

```shell
mvn -Dtest=CsvInvoiceCalculatorTest test
```

To verify packaging as well as compilation and tests, run:

```shell
mvn clean verify
```

The current project does not configure a coverage-report plugin. The unit tests cover normal calculations, rounding, CSV parsing, preserving source columns, and invalid input. If a numeric coverage percentage is required, configure JaCoCo in `pom.xml` and run `mvn clean verify`; Maven Surefire test results alone do not measure coverage.

## Add a test for your own case

Create a test class under `src/test/java/io/github/commoncsv/`. The following example demonstrates the public tabular API and checks the generated XLSX values:

```java
package io.github.commoncsv;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class MyInvoiceCaseTest {
    @Test
    void calculatesTotalsForMyInput() throws Exception {
        var input = new TabularData(
            List.of("Item", "Qty", "Price", "VAT"),
            List.of(List.of("Notebook", "3", "2.50", "10%")));

        byte[] output = CsvInvoiceCalculator.process(
            input, new ColumnMapping(0, 1, 2, 3), OutputLabels.english());

        try (var workbook = WorkbookFactory.create(new ByteArrayInputStream(output))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals(7.50, sheet.getRow(1).getCell(4).getNumericCellValue(), 0.001);
            assertEquals(8.25, sheet.getRow(1).getCell(5).getNumericCellValue(), 0.001);
        }
    }
}
```

Run only this test with:

```shell
mvn -Dtest=MyInvoiceCaseTest test
```

For validation cases, use `assertThrows(IllegalArgumentException.class, () -> ...)`. For CSV-specific cases, pass UTF-8 bytes to `CsvInvoiceCalculator.processCsv(...)`, then open the returned bytes with Apache POI as in the example to assert values and labels. You can also run tests from IntelliJ IDEA or Eclipse by right-clicking the test class and selecting **Run**; Maven is the standard way to run the same suite in automation.
