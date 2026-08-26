package com.alibaba.excel.markdown;

import java.util.ArrayList;
import java.util.List;

/**
 * A whole workbook document: the file title plus every sheet extracted as a {@link SheetTable}.
 * The Markdown rendering keeps all sheets, each one behind its own {@code ##} heading, which is
 * the main difference with markitdown's single-sheet converter.
 *
 * @author wandl
 */
public final class SheetDocument {

    /**
     * Document title, usually the source file name.
     */
    public String title;

    /**
     * All sheets of the workbook in reading order.
     */
    public List<SheetTable> sheets = new ArrayList<>();

    /**
     * Render the whole document as Markdown: a top level {@code #} title followed by every
     * sheet's Markdown fragment separated by a blank line.
     */
    public String toMarkdown() {
        StringBuilder builder = new StringBuilder();
        builder.append("# ").append(title == null ? "" : title).append("\n\n");
        for (SheetTable sheet : sheets) {
            if (sheet == null) {
                continue;
            }
            builder.append(sheet.toMarkdown());
            builder.append('\n');
        }
        return builder.toString();
    }
}
