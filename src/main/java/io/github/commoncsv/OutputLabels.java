package io.github.commoncsv;

import java.util.Objects;

/**
 * Provides labels written for calculated columns and grand totals.
 *
 * @param beforeTax the label for each calculated amount before tax
 * @param afterTax the label for each calculated amount after tax
 * @param totalBeforeTax the label for the grand total before tax
 * @param totalAfterTax the label for the grand total after tax
 */
public record OutputLabels(String beforeTax, String afterTax, String totalBeforeTax, String totalAfterTax) {

    /**
     * Ensures that every output label is available to the workbook writer.
     *
     * @param beforeTax the label for each calculated amount before tax
     * @param afterTax the label for each calculated amount after tax
     * @param totalBeforeTax the label for the grand total before tax
     * @param totalAfterTax the label for the grand total after tax
     * @throws NullPointerException if any label is {@code null}
     */
    public OutputLabels {
        Objects.requireNonNull(beforeTax);
        Objects.requireNonNull(afterTax);
        Objects.requireNonNull(totalBeforeTax);
        Objects.requireNonNull(totalAfterTax);
    }

    /**
     * Creates the standard English output labels.
     *
     * @return labels written in English
     */
    public static OutputLabels english() {
        return new OutputLabels("Total before tax", "Total after tax", "Grand total before tax", "Grand total after tax");
    }

    /**
     * Creates the standard Vietnamese output labels.
     *
     * @return labels written in Vietnamese
     */
    public static OutputLabels vietnamese() {
        return new OutputLabels("Tổng trước thuế", "Tổng sau thuế", "Tổng cộng trước thuế", "Tổng cộng sau thuế");
    }

    /**
     * Creates the standard Japanese output labels.
     *
     * @return labels written in Japanese
     */
    public static OutputLabels japanese() {
        return new OutputLabels("税抜合計", "税込合計", "税抜総合計", "税込総合計");
    }

    /**
     * Creates the standard Korean output labels.
     *
     * @return labels written in Korean
     */
    public static OutputLabels korean() {
        return new OutputLabels("세전 금액", "세후 금액", "세전 총액", "세후 총액");
    }
}
