package io.github.commoncsv;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class CsvInvoiceCalculatorTest {
    private static final ColumnMapping MAP = new ColumnMapping(0, 1, 2, 3);
    private static TabularData data(String name, String qty, String price, String vat) {
        return new TabularData(List.of("Product", "Quantity", "Unit Price", "VAT"), List.of(List.of(name, qty, price, vat)));
    }
    private static double cell(byte[] xlsx, int row, int col) throws Exception {
        try (var wb = WorkbookFactory.create(new ByteArrayInputStream(xlsx))) { return wb.getSheetAt(0).getRow(row).getCell(col).getNumericCellValue(); }
    }

    @Test void calculatesItemAndGrandTotals() throws Exception {
        var input = new TabularData(List.of("Item", "Qty", "Unit Price", "VAT", "Note"), List.of(
                List.of("A", "2", "10.25", "10", "kept"), List.of("B", "1", "5", "20%", "also kept")));
        byte[] result = CsvInvoiceCalculator.process(input, MAP, OutputLabels.english());
        assertEquals(20.5, cell(result, 1, 5), 0.001); assertEquals(22.55, cell(result, 1, 6), 0.001);
        assertEquals(5, cell(result, 2, 5), 0.001); assertEquals(6, cell(result, 2, 6), 0.001);
        try (var wb = WorkbookFactory.create(new ByteArrayInputStream(result))) {
            assertEquals("Grand total before tax", wb.getSheetAt(0).getRow(3).getCell(0).getStringCellValue());
        }
        assertEquals(25.5, cell(result, 3, 1), 0.001); assertEquals(28.55, cell(result, 3, 3), 0.001);
    }
    @Test void parsesVietnameseSemicolonCsvAndUsesVietnameseLabels() throws Exception {
        String csv = "Tên mặt hàng;Số lượng;Đơn giá;Thuế suất\nBút;2;1,25;10%";
        byte[] result = CsvInvoiceCalculator.processCsv(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)));
        assertEquals(2.5, cell(result, 1, 4), 0.001); assertEquals(2.75, cell(result, 1, 5), 0.001);
        try (var wb = WorkbookFactory.create(new ByteArrayInputStream(result))) { assertEquals("Tổng trước thuế", wb.getSheetAt(0).getRow(0).getCell(4).getStringCellValue()); }
    }
    @Test void parsesQuotedCommaCsvAndPreservesEmbeddedComma() throws Exception {
        String csv = "Product,Quantity,Unit Price,VAT,Note\n\"Pen, blue\",1,2.50,0,\"a, b\"";
        byte[] result = CsvInvoiceCalculator.processCsv(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)));
        try (var wb = WorkbookFactory.create(new ByteArrayInputStream(result))) {
            assertEquals("Pen, blue", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            assertEquals("a, b", wb.getSheetAt(0).getRow(1).getCell(4).getStringCellValue());
        }
    }
    @Test void acceptsZeroPriceAndMaximumVat() throws Exception {
        byte[] result = CsvInvoiceCalculator.process(data("Free", "1", "0", "100"), MAP, OutputLabels.english());
        assertEquals(0d, cell(result, 1, 4), 0.001); assertEquals(0d, cell(result, 1, 5), 0.001);
    }
    @Test void preservesAllSourceColumnsAndAddsCalculatedHeaders() throws Exception {
        var input = new TabularData(List.of("SKU", "Product", "Qty", "Price", "VAT", "Memo"), List.of(List.of("x", "Pen", "1", "1", "0", "custom")));
        byte[] result = CsvInvoiceCalculator.process(input, new ColumnMapping(1, 2, 3, 4), OutputLabels.vietnamese());
        try (var wb = WorkbookFactory.create(new ByteArrayInputStream(result))) {
            var sheet = wb.getSheetAt(0); assertEquals("custom", sheet.getRow(1).getCell(5).getStringCellValue());
            assertEquals("Tổng trước thuế", sheet.getRow(0).getCell(6).getStringCellValue());
        }
    }
    @Test void roundsMoneyHalfUp() throws Exception { assertEquals(1.01, cell(CsvInvoiceCalculator.process(data("x", "1", "1.005", "0"), MAP, OutputLabels.english()), 1, 4), 0.001); }
    @Test void rejectsNonPositiveQuantity() { assertInvalid(data("x", "0", "1", "0"), "quantity"); assertInvalid(data("x", "-1", "1", "0"), "quantity"); }
    @Test void rejectsNegativePrice() { assertInvalid(data("x", "1", "-1", "0"), "unit price"); }
    @Test void rejectsVatOutsidePercentageRange() { assertInvalid(data("x", "1", "1", "101"), "VAT"); assertInvalid(data("x", "1", "1", "-1"), "VAT"); }
    @Test void rejectsNonNumericRequiredValues() { assertInvalid(data("x", "one", "1", "0"), "quantity"); assertInvalid(data("x", "1", "1x", "0"), "unit price"); assertInvalid(data("x", "1", "1", "ten"), "VAT"); }
    @Test void rejectsBlankNameAndMissingNumbers() { assertInvalid(data(" ", "1", "1", "0"), "product name"); assertInvalid(data("x", "", "1", "0"), "quantity"); }
    @Test void rejectsBadMappingAndUnevenRows() {
        assertThrows(IllegalArgumentException.class, () -> new ColumnMapping(0, 0, 2, 3));
        assertThrows(IllegalArgumentException.class, () -> new TabularData(List.of("a", "b"), List.of(List.of("x"))));
        assertThrows(IllegalArgumentException.class, () -> CsvInvoiceCalculator.process(data("x", "1", "1", "0"), new ColumnMapping(0, 1, 2, 4), OutputLabels.english()));
    }
    @Test void rejectsEmptyOrUndetectableCsv() {
        assertThrows(IllegalArgumentException.class, () -> CsvInvoiceCalculator.processCsv(new ByteArrayInputStream(new byte[0])));
        assertThrows(IllegalArgumentException.class, () -> CsvInvoiceCalculator.processCsv(new ByteArrayInputStream("a,b\nx,y".getBytes(StandardCharsets.UTF_8))));
    }
    @Test void rejectsInvalidVatPercentSyntax() { assertInvalid(data("x", "1", "1", "10%%"), "VAT"); }
    private static void assertInvalid(TabularData data, String message) {
        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class, () -> CsvInvoiceCalculator.process(data, MAP, OutputLabels.english()));
        assertTrue(ex.getMessage().contains(message));
    }
}
