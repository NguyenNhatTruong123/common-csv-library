package io.github.commoncsv;

import java.util.List;
import java.util.Objects;

/** A header row and zero or more data rows. Every row must have the same number of cells. */
public record TabularData(List<String> headers, List<List<String>> rows) {
    public TabularData {
        headers = List.copyOf(Objects.requireNonNull(headers, "headers"));
        rows = Objects.requireNonNull(rows, "rows").stream().map(List::copyOf).toList();
        if (headers.isEmpty()) throw new IllegalArgumentException("At least one column is required");
        for (int i = 0; i < rows.size(); i++) {
            if (rows.get(i).size() != headers.size()) throw new IllegalArgumentException("Row " + (i + 1) + " has " + rows.get(i).size() + " cells; expected " + headers.size());
        }
    }
}
