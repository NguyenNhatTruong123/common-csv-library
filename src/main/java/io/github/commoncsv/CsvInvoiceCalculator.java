package io.github.commoncsv;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.charset.StandardCharsets;
import java.text.Normalizer;
import java.util.*;

/**
 * Calculates item totals and returns a downloadable XLSX workbook.
 * Money is rounded to two decimal places using {@link RoundingMode#HALF_UP}.
 */
public final class CsvInvoiceCalculator {

    private CsvInvoiceCalculator() {
    }

    /**
     * Reads a UTF-8 CSV file, detects its delimiter and required headers, and creates an XLSX result.
     * Comma, semicolon, and tab-separated input is supported.
     *
     * @param input the CSV input stream
     * @return the generated XLSX workbook as a byte array
     * @throws IOException if the input cannot be read or the workbook cannot be written
     * @throws IllegalArgumentException if the CSV is empty, malformed, or lacks required columns
     * @throws NullPointerException if {@code input} is {@code null}
     */
    public static byte[] processCsv(InputStream input) throws IOException {
        Objects.requireNonNull(input, "input");

        byte[] bytes = input.readAllBytes();
        String content = new String(bytes, StandardCharsets.UTF_8);

        if (!content.isEmpty() && content.charAt(0) == '\uFEFF') {
            content = content.substring(1);
        }

        char delimiter = detectDelimiter(content);

        try (CSVParser parser = CSVParser.parse(content, CSVFormat.DEFAULT.builder().setDelimiter(delimiter).setTrim(true).build())) {
            Iterator<CSVRecord> it = parser.iterator();

            if (!it.hasNext()) {
                throw new IllegalArgumentException("CSV file is empty");
            }

            CSVRecord header = it.next();
            List<String> headers = new ArrayList<>();
            header.forEach(headers::add);

            List<List<String>> rows = new ArrayList<>();

            while (it.hasNext()) {
                List<String> row = new ArrayList<>();
                it.next().forEach(row::add);
                rows.add(row);
            }

            TabularData data = new TabularData(headers, rows);
            ColumnMapping mapping = detectMapping(headers);

            return process(data, mapping, detectLabels(headers));
        }
    }

    /**
     * Processes parsed tabular input and creates an XLSX result.
     * This entry point can be used for XLSX, JSON, database, or custom format adapters.
     *
     * @param data the parsed input table
     * @param mapping the columns containing the required invoice fields
     * @param labels the labels to write for calculated values and totals
     * @return the generated XLSX workbook as a byte array
     * @throws IOException if the workbook cannot be written
     * @throws IllegalArgumentException if mapped columns or row values are invalid
     * @throws NullPointerException if an argument is {@code null}
     */
    public static byte[] process(TabularData data, ColumnMapping mapping, OutputLabels labels) throws IOException {
        Objects.requireNonNull(data, "data");
        Objects.requireNonNull(mapping, "mapping");
        Objects.requireNonNull(labels, "labels");

        int width = data.headers().size();

        for (int index : new int[]{mapping.productName(), mapping.quantity(), mapping.unitPrice(), mapping.vatPercent()}) {
            if (index >= width) {
                throw new IllegalArgumentException("Mapped column index " + index + " is outside the " + width + " input columns");
            }
        }

        List<BigDecimal> before = new ArrayList<>();
        List<BigDecimal> after = new ArrayList<>();
        BigDecimal sumBefore = BigDecimal.ZERO, sumAfter = BigDecimal.ZERO;

        for (int r = 0; r < data.rows().size(); r++) {
            List<String> row = data.rows().get(r);
            int line = r + 2;

            if (row.get(mapping.productName()).isBlank()) {
                throw invalid(line, "product name", "must not be blank");
            }

            BigDecimal qty = number(row.get(mapping.quantity()), line, "quantity");
            BigDecimal price = number(row.get(mapping.unitPrice()), line, "unit price");
            BigDecimal vat = percent(row.get(mapping.vatPercent()), line);

            if (qty.signum() <= 0) {
                throw invalid(line, "quantity", "must be greater than zero");
            }

            if (price.signum() < 0) {
                throw invalid(line, "unit price", "must not be negative");
            }

            if (vat.signum() < 0 || vat.compareTo(new BigDecimal("100")) > 0) {
                throw invalid(line, "VAT", "must be between 0 and 100 percent");
            }

            BigDecimal net = qty.multiply(price).setScale(2, RoundingMode.HALF_UP);
            BigDecimal gross = net.multiply(BigDecimal.ONE.add(vat.movePointLeft(2))).setScale(2, RoundingMode.HALF_UP);

            before.add(net);
            after.add(gross);
            sumBefore = sumBefore.add(net);
            sumAfter = sumAfter.add(gross);
        }

        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Results");
            Row header = sheet.createRow(0);

            for (int c = 0; c < width; c++) {
                header.createCell(c).setCellValue(data.headers().get(c));
            }

            header.createCell(width).setCellValue(labels.beforeTax());
            header.createCell(width + 1).setCellValue(labels.afterTax());

            for (int r = 0; r < data.rows().size(); r++) {
                Row row = sheet.createRow(r + 1);
                List<String> source = data.rows().get(r);

                for (int c = 0; c < width; c++) {
                    row.createCell(c).setCellValue(source.get(c));
                }

                row.createCell(width).setCellValue(before.get(r).doubleValue());
                row.createCell(width + 1).setCellValue(after.get(r).doubleValue());
            }

            int totalRowIndex = data.rows().size() + 1;
            Row total = sheet.createRow(totalRowIndex);

            total.createCell(0).setCellValue(labels.totalBeforeTax());
            total.createCell(1).setCellValue(sumBefore.setScale(2, RoundingMode.HALF_UP).doubleValue());
            total.createCell(2).setCellValue(labels.totalAfterTax());
            total.createCell(3).setCellValue(sumAfter.setScale(2, RoundingMode.HALF_UP).doubleValue());

            CellStyle money = workbook.createCellStyle();
            money.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));

            for (int r = 1; r <= data.rows().size(); r++) {
                sheet.getRow(r).getCell(width).setCellStyle(money);
                sheet.getRow(r).getCell(width + 1).setCellStyle(money);
            }

            total.getCell(1).setCellStyle(money);
            total.getCell(3).setCellStyle(money);

            for (int c = 0; c < width + 2; c++) {
                sheet.autoSizeColumn(c);
            }

            workbook.write(out);

            return out.toByteArray();
        }
    }

    private static ColumnMapping detectMapping(List<String> headers) {
        int name = find(headers, List.of("product", "product name", "item", "item name", "name", "tên mặt hàng", "tên sản phẩm", "sản phẩm", "商品名", "品名"));
        int quantity = find(headers, List.of("quantity", "qty", "数量", "số lượng", "so luong"));
        int price = find(headers, List.of("unit price", "price per unit", "price", "単価", "đơn giá", "don gia"));
        int vat = find(headers, List.of("vat", "vat %", "vat percent", "tax", "tax rate", "税率", "thuế", "thuế suất", "thue", "thue suat"));

        if (name < 0 || quantity < 0 || price < 0 || vat < 0) {
            throw new IllegalArgumentException("Could not detect all required columns from headers. Use process(data, mapping, labels) with explicit metadata.");
        }

        return new ColumnMapping(name, quantity, price, vat);
    }

    private static int find(List<String> headers, List<String> aliases) {
        Set<String> keys = new HashSet<>();
        for (String alias : aliases) {
            keys.add(normalize(alias));
        }
        for (int i = 0; i < headers.size(); i++) {
            if (keys.contains(normalize(headers.get(i)))) {
                return i;
            }
        }

        return -1;
    }

    private static String normalize(String s) {
        return Normalizer.normalize(s, Normalizer.Form.NFD)
                .replaceAll("\\p{M}", "")
                .toLowerCase(Locale.ROOT)
                .replaceAll("[^\\p{L}\\p{N}%]+", " ")
                .trim();
    }

    private static OutputLabels detectLabels(List<String> headers) {
        String joined = String.join(" ", headers).toLowerCase(Locale.ROOT);
        if (joined.matches(".*[\\p{IsHiragana}\\p{IsKatakana}\\p{IsHan}].*")) {
            return OutputLabels.japanese();
        }

        if (joined.contains("số lượng") || joined.contains("đơn giá") || joined.contains("thuế") || joined.contains("tên")) {
            return OutputLabels.vietnamese();
        }

        return OutputLabels.english();
    }

    private static char detectDelimiter(String csv) {
        String first = csv.lines().findFirst().orElse("");
        char best = ',';
        int max = -1;
        for (char candidate : new char[]{',', ';', '\t'}) {
            int count = 0;
            boolean quoted = false;
            for (int i = 0; i < first.length(); i++) {
                char ch = first.charAt(i);
                if (ch == '"') {
                    quoted = !quoted;
                } else if (ch == candidate && !quoted) {
                    count++;
                }
            }
            if (count > max) {
                max = count;
                best = candidate;
            }
        }

        return best;
    }

    private static BigDecimal number(String raw, int row, String field) {
        if (raw == null || raw.isBlank()) {
            throw invalid(row, field, "is required");
        }

        String value = raw.trim().replace(" ", "");

        if (value.matches("[+-]?\\d+,\\d+") && !value.contains(".")) {
            value = value.replace(',', '.');
        } else if (value.matches("[+-]?\\d{1,3}(,\\d{3})+(\\.\\d+)?")) {
            value = value.replace(",", "");
        }

        if (!value.matches("[+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)")) {
            throw invalid(row, field, "must be a non-negative number using digits and an optional decimal separator");
        }

        try {
            return new BigDecimal(value);
        } catch (NumberFormatException e) {
            throw invalid(row, field, "is not a valid number");
        }
    }

    private static BigDecimal percent(String raw, int row) {
        String v = raw == null ? "" : raw.trim();
        if (v.endsWith("%")) {
            v = v.substring(0, v.length() - 1).trim();
        }

        return number(v, row, "VAT");
    }

    private static IllegalArgumentException invalid(int row, String field, String message) {
        return new IllegalArgumentException("CSV row " + row + ": " + field + " " + message);
    }
}
