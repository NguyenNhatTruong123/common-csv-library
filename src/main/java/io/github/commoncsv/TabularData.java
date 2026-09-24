package io.github.commoncsv;

import java.util.List;
import java.util.Objects;

/**
 * Represents a header row and zero or more data rows with a consistent width.
 *
 * @param headers the input column names
 * @param rows the input data rows
 */
public record TabularData(List<String> headers, List<List<String>> rows) {

    /**
     * Copies the supplied table and validates that every row matches the header width.
     *
     * @param headers the input column names
     * @param rows the input data rows
     * @throws NullPointerException if the headers or rows are {@code null}
     * @throws IllegalArgumentException if there are no headers or a row has a different width
     */
    public TabularData {
        headers = List.copyOf(Objects.requireNonNull(headers, "headers"));
        rows = Objects.requireNonNull(rows, "rows").stream().map(List::copyOf).toList();

        if (headers.isEmpty()) {
            throw new IllegalArgumentException("At least one column is required");
        }

        for (int i = 0; i < rows.size(); i++) {
            if (rows.get(i).size() != headers.size()) {
                throw new IllegalArgumentException("Row " + (i + 1) + " has " + rows.get(i).size()
                        + " cells; expected " + headers.size());
            }
        }
    }
}
