package io.github.commoncsv;

/** Zero-based input column indexes for the required product fields. */
public record ColumnMapping(int productName, int quantity, int unitPrice, int vatPercent) {
    public ColumnMapping {
        if (productName < 0 || quantity < 0 || unitPrice < 0 || vatPercent < 0) throw new IllegalArgumentException("Column indexes must be zero or greater");
        if (productName == quantity || productName == unitPrice || productName == vatPercent || quantity == unitPrice || quantity == vatPercent || unitPrice == vatPercent)
            throw new IllegalArgumentException("Required fields must map to different columns");
    }
}
