package io.github.commoncsv;

/**
 * Defines the zero-based input columns used for invoice calculations.
 *
 * @param productName the column containing the product name
 * @param quantity the column containing the quantity
 * @param unitPrice the column containing the unit price
 * @param vatPercent the column containing the VAT percentage
 */
public record ColumnMapping(int productName, int quantity, int unitPrice, int vatPercent) {

    /**
     * Validates that all required fields use non-negative and distinct columns.
     *
     * @param productName the column containing the product name
     * @param quantity the column containing the quantity
     * @param unitPrice the column containing the unit price
     * @param vatPercent the column containing the VAT percentage
     * @throws IllegalArgumentException if an index is negative or two fields use the same column
     */
    public ColumnMapping {
        if (productName < 0 || quantity < 0 || unitPrice < 0 || vatPercent < 0) {
            throw new IllegalArgumentException("Column indexes must be zero or greater");
        }

        if (productName == quantity || productName == unitPrice || productName == vatPercent
                || quantity == unitPrice || quantity == vatPercent || unitPrice == vatPercent) {
            throw new IllegalArgumentException("Required fields must map to different columns");
        }
    }
}
