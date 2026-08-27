package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Test multi-row header support in {@link XlsxToMarkdownConverter}.
 *
 * @author wandl
 */
public class MultiRowHeaderTest {

    private static final int TWO_HEADER_ROWS = 2;

    @TempDir
    Path tempDir;

    /**
     * Fixture: two-row header where row 0 has "Basic" merged across A:B and "Extra" in C,
     * row 1 has Name/Age/Note, plus one data row.
     */
    private File writeTwoRowHeaderXlsx() throws IOException {
        File file = tempDir.resolve("two-row-header.xlsx").toFile();
        Workbook workbook = new XSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("Grouped");
            Row row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue("Basic");
            row0.createCell(1);
            row0.createCell(2).setCellValue("Extra");
            Row row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue("Name");
            row1.createCell(1).setCellValue("Age");
            row1.createCell(2).setCellValue("Note");
            Row row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue("Tom");
            row2.createCell(1).setCellValue("18");
            row2.createCell(2).setCellValue("ok");
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
            try (OutputStream out = new FileOutputStream(file)) {
                workbook.write(out);
            }
        } finally {
            workbook.close();
        }
        return file;
    }

    /**
     * Fixture: two-row header with a vertical merge A1:A2 (same value).
     * Column B has no merge, just separate values in each row.
     */
    private File writeVerticalMergedHeaderXlsx() throws IOException {
        File file = tempDir.resolve("vertical-header.xlsx").toFile();
        Workbook workbook = new XSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("VMerged");
            Row row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue("Group");
            row0.createCell(1).setCellValue("Other");
            Row row1 = sheet.createRow(1);
            row1.createCell(0);
            row1.createCell(1).setCellValue("Sub");
            Row row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue("val");
            row2.createCell(1).setCellValue("data");
            sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
            try (OutputStream out = new FileOutputStream(file)) {
                workbook.write(out);
            }
        } finally {
            workbook.close();
        }
        return file;
    }

    @Test
    public void twoRowGroupHeaderFlattens() throws IOException {
        SheetDocument doc = XlsxToMarkdownConverter.toStructured(writeTwoRowHeaderXlsx(), TWO_HEADER_ROWS);

        assertThat(doc.sheets).hasSize(1);
        SheetTable table = doc.sheets.get(0);
        assertThat(table.sheetName).isEqualTo("Grouped");

        // Two header rows promoted
        assertThat(table.headers).hasSize(TWO_HEADER_ROWS);
        // Row 0: merge fill skipped for header cells, B1 stays empty
        assertThat(table.headers.get(0)).containsExactly("Basic", "", "Extra");
        // Row 1
        assertThat(table.headers.get(1)).containsExactly("Name", "Age", "Note");

        // Data rows
        assertThat(table.rows).hasSize(1);
        assertThat(table.rows.get(0)).containsExactly("Tom", "18", "ok");

        // Flattened: "Basic Name", "Age" (B1 empty), "Extra Note"
        String md = table.toMarkdown();
        assertThat(md).contains("| Basic Name | Age | Extra Note |");
        assertThat(md).contains("| --- | --- | --- |");
        assertThat(md).contains("| Tom | 18 | ok |");
    }

    @Test
    public void mergedVerticalHeaderCells() throws IOException {
        SheetDocument doc = XlsxToMarkdownConverter.toStructured(writeVerticalMergedHeaderXlsx(), TWO_HEADER_ROWS);

        SheetTable table = doc.sheets.get(0);
        assertThat(table.headers).hasSize(TWO_HEADER_ROWS);

        // Vertical merge A1:A2: A2 stays empty (not filled in header range)
        // Row 0: ["Group", "Other"], Row 1: ["", "Sub"]
        assertThat(table.headers.get(0)).containsExactly("Group", "Other");
        assertThat(table.headers.get(1)).containsExactly("", "Sub");

        // Flattened: col A = "Group" (single value, no duplicate), col B = "Other Sub"
        String md = table.toMarkdown();
        assertThat(md).contains("| Group | Other Sub |");
        assertThat(md).contains("| val | data |");
    }

    @Test
    public void defaultBehaviorUnchanged() throws IOException {
        File file = writeTwoRowHeaderXlsx();
        String defaultOutput = XlsxToMarkdownConverter.toMarkdown(file);
        String explicitOne = XlsxToMarkdownConverter.toMarkdown(file, 1);
        assertThat(explicitOne).isEqualTo(defaultOutput);
    }

    @Test
    public void invalidHeadRowNumberRejected() throws IOException {
        File file = writeTwoRowHeaderXlsx();
        assertThatThrownBy(() -> XlsxToMarkdownConverter.toStructured(file, 0))
            .isInstanceOf(IllegalArgumentException.class)
            .hasMessageContaining("headRowNumber");
        assertThatThrownBy(() -> XlsxToMarkdownConverter.toStructured(file, -1))
            .isInstanceOf(IllegalArgumentException.class)
            .hasMessageContaining("headRowNumber");
    }
}
