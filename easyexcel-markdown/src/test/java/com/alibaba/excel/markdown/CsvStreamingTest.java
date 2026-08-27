package com.alibaba.excel.markdown;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test the two-pass streaming CSV to Markdown path introduced for GB-scale CSV files.
 * Every test compares the {@code Writer} API output against the {@code String} API output
 * to guarantee byte-for-byte equivalence, or against known expected output for edge cases.
 *
 * @author wandl
 */
public class CsvStreamingTest {

    @TempDir
    File tempDir;

    /**
     * Generate a CSV file with {@code dataRows} data rows and 3 columns, then assert that
     * {@code toMarkdown(file, writer)} produces the same output as {@code toMarkdown(file)}.
     */
    @Test
    public void csvWriterOutputMatchesStringApiLargeRows() throws IOException {
        int dataRows = 50_000;
        StringBuilder sb = new StringBuilder();
        sb.append("Name,Age,City\n");
        for (int i = 0; i < dataRows; i++) {
            sb.append("user_").append(i).append(',').append(i % 100).append(",City\n");
        }
        File file = new File(tempDir, "large.csv");
        Files.write(file.toPath(), sb.toString().getBytes(StandardCharsets.UTF_8));

        String stringApi = XlsxToMarkdownConverter.toMarkdown(file);
        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(file, sw);

        assertThat(sw.toString()).isEqualTo(stringApi);
    }

    /**
     * Construct a CSV with ragged rows (different column counts) and assert that all rows in
     * the Markdown output are padded to the widest row's column count.
     */
    @Test
    public void csvStreamTwoPassColumnWidth() throws IOException {
        // Header: 2 cols, row1: 1 col, row2: 3 cols -> maxColumns = 3
        String csv = "h1,h2\nonly-a\na,b,c\n";
        File file = new File(tempDir, "ragged.csv");
        Files.write(file.toPath(), csv.getBytes(StandardCharsets.UTF_8));

        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(file, sw);
        String md = sw.toString();

        // Header padded to 3 columns: | h1 | h2 |  |
        assertThat(md).contains("| h1 | h2 |  |");
        // Data row padded to 3 columns: | only-a |  |  |
        assertThat(md).contains("| only-a |  |  |");
        // Data row with 3 columns: | a | b | c |
        assertThat(md).contains("| a | b | c |");
        // Separator must have 3 columns
        assertThat(md).contains("| --- | --- | --- |");
    }

    /**
     * Empty CSV produces the {@code (empty sheet)} placeholder.
     */
    @Test
    public void csvEmptyFileProducesEmptySheet() throws IOException {
        File file = new File(tempDir, "empty.csv");
        Files.write(file.toPath(), new byte[0]);

        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(file, sw);
        String md = sw.toString();

        assertThat(md).contains("(empty sheet)\n");
    }

    /**
     * CSV with only a header line and no data rows still produces a valid Markdown table
     * (header + separator, no data rows).
     */
    @Test
    public void csvHeaderOnlyProducesTableWithoutDataRows() throws IOException {
        File file = new File(tempDir, "headeronly.csv");
        Files.write(file.toPath(), "Name,Age\n".getBytes(StandardCharsets.UTF_8));

        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(file, sw);
        String md = sw.toString();

        assertThat(md).contains("| Name | Age |");
        assertThat(md).contains("| --- | --- |");
        // No data row after separator
        assertThat(md).doesNotContain("| Name | Age |\n| --- | --- |\n|");
        // Must not contain the empty sheet placeholder
        assertThat(md).doesNotContain("(empty sheet)");
    }

    /**
     * CSV with UTF-8 BOM and semicolon delimiter (same scenario as CsvDetectorTest) still works
     * through the streaming path. This is a regression guard for charset + delimiter detection.
     */
    @Test
    public void csvBomSemicolonStillWorks() throws IOException {
        String csv = "\uFEFFa;b;c\n1;2;3\n";
        File file = new File(tempDir, "bomsemi.csv");
        Files.write(file.toPath(), csv.getBytes(StandardCharsets.UTF_8));

        String stringApi = XlsxToMarkdownConverter.toMarkdown(file);
        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(file, sw);

        assertThat(sw.toString()).isEqualTo(stringApi);
        assertThat(sw.toString()).contains("| a | b | c |");
        assertThat(sw.toString()).contains("| 1 | 2 | 3 |");
    }

    /**
     * InputStream path spools CSV to a temp file and produces the same output as the String API.
     */
    @Test
    public void csvInputStreamSpoolsAndMatches() throws IOException {
        String csv = "Name,Score\nAlice,95\nBob,87\n";
        byte[] bytes = csv.getBytes(StandardCharsets.UTF_8);

        String stringApi = XlsxToMarkdownConverter.toMarkdown(new ByteArrayInputStream(bytes));
        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(new ByteArrayInputStream(bytes), sw);

        assertThat(sw.toString()).isEqualTo(stringApi);
    }
}
