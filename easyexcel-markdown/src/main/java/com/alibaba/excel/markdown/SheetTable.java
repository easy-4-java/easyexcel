package com.alibaba.excel.markdown;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * A single sheet extracted from a workbook, kept as plain string cells so that it can be rendered
 * to Markdown without any Excel specific object model.
 * <p>
 * Multi-row headers are flattened into one row when rendering: distinct non-empty values of the
 * same column are joined with a single space, which keeps hierarchy information such as
 * {@code 基本信息 姓名} readable.
 *
 * @author wandl
 */
public final class SheetTable {

    /**
     * Sheet name, empty string when unknown.
     */
    public String sheetName = "";

    /**
     * Header rows, outer list is the header row index, inner list is the cell value per column.
     */
    public List<List<String>> headers = new ArrayList<>();

    /**
     * Data rows, outer list is the row index, inner list is the cell value per column.
     */
    public List<List<String>> rows = new ArrayList<>();

    /**
     * Per-column alignment derived from the header row's cell styles: each entry is
     * {@code "center"}, {@code "right"}, or {@code null} (left / general, the default).
     * Empty list means no alignment information is available.
     */
    public List<String> columnAlignments = Collections.emptyList();

    /**
     * Render this sheet as a Markdown fragment: a {@code ##} heading followed by a GitHub
     * Flavored Markdown table. When no header was captured, the first data row is promoted to the
     * header row (same behaviour as markitdown). All rows are padded to the widest row so the
     * table stays rectangular.
     */
    public String toMarkdown() {
        StringBuilder builder = new StringBuilder();
        builder.append("## ").append(sheetName == null ? "" : sheetName).append("\n\n");

        List<String> headerRow = flattenedHeaders();
        List<List<String>> dataRows = this.rows;
        if (headerRow.isEmpty() && !rows.isEmpty()) {
            headerRow = rows.get(0);
            dataRows = rows.subList(1, rows.size());
        }
        if (headerRow.isEmpty() && dataRows.isEmpty()) {
            return builder.append("(empty sheet)").append("\n").toString();
        }

        int columnCount = columnCount(headerRow, dataRows);
        appendTableRow(builder, headerRow, columnCount);
        appendTableSeparator(builder, columnCount);
        for (List<String> row : dataRows) {
            appendTableRow(builder, row, columnCount);
        }
        return builder.toString();
    }

    private List<String> flattenedHeaders() {
        List<String> flattened = new ArrayList<>();
        for (int columnIndex = 0; columnIndex < headerColumnCount(); columnIndex++) {
            StringBuilder cell = new StringBuilder();
            for (List<String> header : headers) {
                if (header == null || columnIndex >= header.size()) {
                    continue;
                }
                String value = header.get(columnIndex);
                if (value == null || value.isEmpty()) {
                    continue;
                }
                if (cell.length() > 0 && !cell.toString().endsWith(" ")) {
                    cell.append(' ');
                }
                cell.append(value.trim());
            }
            flattened.add(cell.toString());
        }
        return flattened;
    }

    private int headerColumnCount() {
        int count = 0;
        for (List<String> header : headers) {
            if (header != null) {
                count = Math.max(count, header.size());
            }
        }
        return count;
    }

    private int columnCount(List<String> headerRow, List<List<String>> dataRows) {
        int count = headerRow == null ? 0 : headerRow.size();
        if (dataRows != null) {
            for (List<String> row : dataRows) {
                if (row != null) {
                    count = Math.max(count, row.size());
                }
            }
        }
        return count;
    }

    private void appendTableRow(StringBuilder builder, List<String> row, int columnCount) {
        builder.append('|');
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            builder.append(' ').append(escape(cellAt(row, columnIndex))).append(" |");
        }
        builder.append('\n');
    }

    private void appendTableSeparator(StringBuilder builder, int columnCount) {
        builder.append('|');
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            String align = columnIndex < columnAlignments.size() ? columnAlignments.get(columnIndex) : null;
            if ("center".equals(align)) {
                builder.append(" :---: |");
            } else if ("right".equals(align)) {
                builder.append(" ---: |");
            } else {
                builder.append(" --- |");
            }
        }
        builder.append('\n');
    }

    private String cellAt(List<String> row, int columnIndex) {
        if (row == null || columnIndex >= row.size()) {
            return "";
        }
        String value = row.get(columnIndex);
        return value == null ? "" : value;
    }

    private String escape(String value) {
        return value.replace("|", "\\|").replace("\r\n", "\n").replace('\r', '\n').replace("\n", "<br>");
    }
}
