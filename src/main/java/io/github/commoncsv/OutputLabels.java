package io.github.commoncsv;

import java.util.Objects;

/** Text written for calculated columns and grand totals. */
public record OutputLabels(String beforeTax, String afterTax, String totalBeforeTax, String totalAfterTax) {
    public OutputLabels { Objects.requireNonNull(beforeTax); Objects.requireNonNull(afterTax); Objects.requireNonNull(totalBeforeTax); Objects.requireNonNull(totalAfterTax); }
    public static OutputLabels english() { return new OutputLabels("Total before tax", "Total after tax", "Grand total before tax", "Grand total after tax"); }
    public static OutputLabels vietnamese() { return new OutputLabels("Tổng trước thuế", "Tổng sau thuế", "Tổng cộng trước thuế", "Tổng cộng sau thuế"); }
    public static OutputLabels japanese() { return new OutputLabels("税抜合計", "税込合計", "税抜総合計", "税込総合計"); }
}
