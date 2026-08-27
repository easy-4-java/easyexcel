package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test that a blank row in the middle of the data is kept while trailing blank rows are cut.
 *
 * @author wandl
 */
public class EmptyRowTest {

    @TempDir
    File tempDir;

    @Test
    public void keepMidTableBlankRowAndTrimTrailing() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Blank");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("A");
        header.createCell(1).setCellValue("B");
        Row first = sheet.createRow(1);
        first.createCell(0).setCellValue("1");
        first.createCell(1).setCellValue("2");
        // row index 2 exists in the XML but has no cell at all: the mid-table blank row
        sheet.createRow(2);
        Row third = sheet.createRow(3);
        third.createCell(0).setCellValue("3");
        third.createCell(1).setCellValue("4");
        // trailing styled-empty rows must be trimmed
        sheet.createRow(4);
        sheet.createRow(5);

        File file = new File(tempDir, "blank.xlsx");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);

        List<List<String>> rows = document.sheets.get(0).rows;
        assertThat(rows).containsExactly(
            Arrays.asList("1", "2"),
            Arrays.asList("", ""),
            Arrays.asList("3", "4"));

        assertThat(document.toMarkdown()).contains("|  |  |");
    }
}
