package com.alibaba.excel.markdown;

import java.util.Arrays;
import java.util.List;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test {@link SheetDocument} and {@link SheetTable} Markdown rendering.
 *
 * @author wandl
 */
public class SheetDocumentTest {

    private static List<String> row(String... cells) {
        return Arrays.asList(cells);
    }

    @Test
    public void renderSingleSheetWithHeaderAndRows() {
        SheetTable table = new SheetTable();
        table.sheetName = "Users";
        table.headers.add(row("Name", "Age"));
        table.rows.add(row("Zhang San", "20"));
        table.rows.add(row("Li Si", "30"));

        SheetDocument document = new SheetDocument();
        document.title = "users.xlsx";
        document.sheets.add(table);

        assertThat(document.toMarkdown()).isEqualTo(""
            + "# users.xlsx\n"
            + "\n"
            + "## Users\n"
            + "\n"
            + "| Name | Age |\n"
            + "| --- | --- |\n"
            + "| Zhang San | 20 |\n"
            + "| Li Si | 30 |\n"
            + "\n");
    }

    @Test
    public void renderMultipleSheetsInOrder() {
        SheetDocument document = new SheetDocument();
        document.title = "book.xlsx";

        SheetTable first = new SheetTable();
        first.sheetName = "Sheet1";
        first.headers.add(row("A"));
        first.rows.add(row("1"));
        document.sheets.add(first);

        SheetTable second = new SheetTable();
        second.sheetName = "Sheet2";
        second.headers.add(row("B"));
        second.rows.add(row("2"));
        document.sheets.add(second);

        assertThat(document.toMarkdown()).isEqualTo(""
            + "# book.xlsx\n"
            + "\n"
            + "## Sheet1\n"
            + "\n"
            + "| A |\n"
            + "| --- |\n"
            + "| 1 |\n"
            + "\n"
            + "## Sheet2\n"
            + "\n"
            + "| B |\n"
            + "| --- |\n"
            + "| 2 |\n"
            + "\n");
    }

    @Test
    public void flattenMultiRowHeaders() {
        SheetTable table = new SheetTable();
        table.sheetName = "Complex";
        table.headers.add(row("Basic", null, "Extra"));
        table.headers.add(row("Name", "Age", "Note"));
        table.rows.add(row("Tom", "18", "ok"));

        assertThat(table.toMarkdown()).isEqualTo(""
            + "## Complex\n"
            + "\n"
            + "| Basic Name | Age | Extra Note |\n"
            + "| --- | --- | --- |\n"
            + "| Tom | 18 | ok |\n");
    }

    @Test
    public void promoteFirstRowAsHeaderWhenHeadersMissing() {
        SheetTable table = new SheetTable();
        table.sheetName = "NoHead";
        table.rows.add(row("H1", "H2"));
        table.rows.add(row("v1", "v2"));

        assertThat(table.toMarkdown()).isEqualTo(""
            + "## NoHead\n"
            + "\n"
            + "| H1 | H2 |\n"
            + "| --- | --- |\n"
            + "| v1 | v2 |\n");
    }

    @Test
    public void padRaggedRowsToWidestRow() {
        SheetTable table = new SheetTable();
        table.sheetName = "Ragged";
        table.headers.add(row("A", "B", "C"));
        table.rows.add(row("only-a"));
        table.rows.add(row("a", "b", "c", "d"));

        assertThat(table.toMarkdown()).isEqualTo(""
            + "## Ragged\n"
            + "\n"
            + "| A | B | C |  |\n"
            + "| --- | --- | --- | --- |\n"
            + "| only-a |  |  |  |\n"
            + "| a | b | c | d |\n");
    }

    @Test
    public void escapePipesAndNewlines() {
        SheetTable table = new SheetTable();
        table.sheetName = "Escape";
        table.headers.add(row("a|b"));
        table.rows.add(row("line1\nline2"));

        assertThat(table.toMarkdown()).isEqualTo(""
            + "## Escape\n"
            + "\n"
            + "| a\\|b |\n"
            + "| --- |\n"
            + "| line1<br>line2 |\n");
    }

    @Test
    public void renderEmptySheetPlaceholder() {
        SheetTable table = new SheetTable();
        table.sheetName = "Empty";

        assertThat(table.toMarkdown()).isEqualTo(""
            + "## Empty\n"
            + "\n"
            + "(empty sheet)\n");
    }
}
