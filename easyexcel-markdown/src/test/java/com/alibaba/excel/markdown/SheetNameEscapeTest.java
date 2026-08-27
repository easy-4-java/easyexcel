package com.alibaba.excel.markdown;

import java.util.Arrays;
import java.util.List;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test that sheet names containing Markdown-significant characters (newlines, tabs, {@code #})
 * are escaped before being rendered into the {@code ##} heading line.
 * <p>
 * Escape rules (defined in {@link SheetTable#escapeSheetName(String)}):
 * <ul>
 *   <li>CR, LF, CRLF and TAB are collapsed into a single space.</li>
 *   <li>Leading/trailing whitespace is trimmed.</li>
 *   <li>A leading {@code #} is escaped as {@code \#} to prevent it from being interpreted as a
 *       nested ATX heading.</li>
 * </ul>
 *
 * @author wandl
 */
public class SheetNameEscapeTest {

    private static List<String> row(String... cells) {
        return Arrays.asList(cells);
    }

    /**
     * A sheet name containing embedded newlines and a leading {@code #} must produce a single-line
     * heading whose line count is exactly one (plus the blank line separator).
     */
    @Test
    public void newlineAndHashInSheetName() {
        SheetTable table = new SheetTable();
        table.sheetName = "#Title\nWith\nNewlines";
        table.headers.add(row("A"));
        table.rows.add(row("1"));

        String md = table.toMarkdown();

        // The heading line must be a single line starting with "## "
        String[] lines = md.split("\n", -1);
        assertThat(lines[0]).isEqualTo("## \\#Title With Newlines");

        // The heading block (## + blank) is exactly 2 lines before the table
        assertThat(lines[1]).isEmpty();
        // The table starts on line 2
        assertThat(lines[2]).startsWith("| A |");
    }

    /**
     * A sheet name with CRLF must not produce any extra lines in the heading region.
     */
    @Test
    public void crlfCollapsedToSpace() {
        SheetTable table = new SheetTable();
        table.sheetName = "Hello\r\nWorld";
        table.headers.add(row("X"));
        table.rows.add(row("v"));

        String md = table.toMarkdown();
        assertThat(md).startsWith("## Hello World\n\n");
    }

    /**
     * A sheet name that is only {@code #} characters still produces a valid heading.
     */
    @Test
    public void pureHashSheetName() {
        SheetTable table = new SheetTable();
        table.sheetName = "###";
        table.headers.add(row("A"));
        table.rows.add(row("1"));

        String md = table.toMarkdown();
        assertThat(md).startsWith("## \\###\n\n");
    }

    /**
     * A sheet name with tabs is collapsed to spaces.
     */
    @Test
    public void tabInSheetName() {
        SheetTable table = new SheetTable();
        table.sheetName = "A\tB";
        table.headers.add(row("A"));
        table.rows.add(row("1"));

        String md = table.toMarkdown();
        assertThat(md).startsWith("## A B\n\n");
    }

    /**
     * Whitespace-only sheet name collapses to empty heading.
     */
    @Test
    public void whitespaceOnlySheetName() {
        SheetTable table = new SheetTable();
        table.sheetName = "   \n  \t  ";
        table.headers.add(row("A"));
        table.rows.add(row("1"));

        String md = table.toMarkdown();
        assertThat(md).startsWith("## \n\n");
    }

    /**
     * A normal sheet name passes through unchanged.
     */
    @Test
    public void normalSheetNameUnchanged() {
        SheetTable table = new SheetTable();
        table.sheetName = "Sales Report Q4";
        table.headers.add(row("A"));
        table.rows.add(row("1"));

        String md = table.toMarkdown();
        assertThat(md).startsWith("## Sales Report Q4\n\n");
    }

    /**
     * Null sheet name produces an empty heading.
     */
    @Test
    public void nullSheetNameBecomesEmpty() {
        SheetTable table = new SheetTable();
        table.sheetName = null;
        table.headers.add(row("A"));
        table.rows.add(row("1"));

        String md = table.toMarkdown();
        assertThat(md).startsWith("## \n\n");
    }
}
