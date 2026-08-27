package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for cell style mapping: inline font markers (bold / italic / strikeout) and
 * column alignment derived from header row cell styles.
 *
 * @author wandl
 */
public class CellStyleMappingTest {

    private static final int INITIAL_CAPACITY = 4;

    @TempDir
    File tempDir;

    // ------------------------------------------------------------------ XLSX

    @Test
    public void xlsxInlineFontMarkers() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Fonts");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("A");
        header.createCell(1).setCellValue("B");
        header.createCell(2).setCellValue("C");
        header.createCell(3).setCellValue("D");

        Row data = sheet.createRow(1);
        // bold
        Cell boldCell = data.createCell(0);
        boldCell.setCellValue("bold");
        boldCell.setCellStyle(boldStyle(workbook));
        // italic
        Cell italicCell = data.createCell(1);
        italicCell.setCellValue("ital");
        italicCell.setCellStyle(italicStyle(workbook));
        // bold + italic + strikeout
        Cell mixCell = data.createCell(2);
        mixCell.setCellValue("mix");
        mixCell.setCellStyle(boldItalicStrikeStyle(workbook));
        // plain
        data.createCell(3).setCellValue("plain");

        String markdown = convertToMarkdown(workbook, "font.xlsx");

        assertThat(markdown).contains("**bold**");
        assertThat(markdown).contains("*ital*");
        assertThat(markdown).contains("~~***mix***~~");
        // plain cell should have no markers
        assertThat(markdown).contains("| plain |");
        assertThat(markdown).doesNotContain("**plain**");
        assertThat(markdown).doesNotContain("*plain*");
    }

    @Test
    public void xlsxColumnAlignmentFromHeader() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Align");
        Row header = sheet.createRow(0);

        CellStyle centerStyle = workbook.createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        Cell centerCell = header.createCell(0);
        centerCell.setCellValue("C");
        centerCell.setCellStyle(centerStyle);

        CellStyle rightStyle = workbook.createCellStyle();
        rightStyle.setAlignment(HorizontalAlignment.RIGHT);
        Cell rightCell = header.createCell(1);
        rightCell.setCellValue("R");
        rightCell.setCellStyle(rightStyle);

        // left / general (default)
        header.createCell(2).setCellValue("L");

        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("c1");
        data.createCell(1).setCellValue("c2");
        data.createCell(2).setCellValue("c3");

        String markdown = convertToMarkdown(workbook, "align.xlsx");

        assertThat(markdown).contains("| :---: |");
        assertThat(markdown).contains("| ---: |");
        assertThat(markdown).contains("| --- |");
        // Full separator line check
        assertThat(markdown).contains("| :---: | ---: | --- |");
    }

    @Test
    public void styleAndHyperlinkCombine() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Combo");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Link");

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue("click");

        // Bold style
        cell.setCellStyle(boldStyle(workbook));

        // Hyperlink
        CreationHelper helper = workbook.getCreationHelper();
        Hyperlink hyperlink = helper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress("https://example.com");
        cell.setHyperlink(hyperlink);

        String markdown = convertToMarkdown(workbook, "combo.xlsx");

        // bold wraps the hyperlink rendering: **[click](https://example.com)**
        assertThat(markdown).contains("**[click](https://example.com)**");
    }

    // ------------------------------------------------------------------ XLS

    @Test
    public void xlsStyleMapping() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("XlsFonts");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Bold");
        header.createCell(1).setCellValue("Right");

        CellStyle rightStyle = workbook.createCellStyle();
        rightStyle.setAlignment(HorizontalAlignment.RIGHT);
        header.getCell(1).setCellStyle(rightStyle);

        Row data = sheet.createRow(1);
        Cell boldCell = data.createCell(0);
        boldCell.setCellValue("strong");
        boldCell.setCellStyle(boldStyle(workbook));
        data.createCell(1).setCellValue("val");

        String markdown = convertToMarkdown(workbook, "style.xls");

        assertThat(markdown).contains("**strong**");
        assertThat(markdown).contains("| ---: |");
    }

    // ------------------------------------------------------------------ Plain

    @Test
    public void plainFileOutputUnchanged() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Plain");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Value");

        Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("Alice");
        r1.createCell(1).setCellValue("100");

        Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue("Bob");
        r2.createCell(1).setCellValue("200");

        String markdown = convertToMarkdown(workbook, "plain.xlsx");

        // Separator must be plain --- with no alignment colons
        assertThat(markdown).contains("| --- | --- |");
        assertThat(markdown).doesNotContain(":---:");
        assertThat(markdown).doesNotContain("---:");
        // No inline markers
        assertThat(markdown).doesNotContain("**");
        assertThat(markdown).doesNotContain("~~");
        // Values present
        assertThat(markdown).contains("| Alice |");
        assertThat(markdown).contains("| Bob |");
    }

    // ------------------------------------------------------------------ helpers

    private String convertToMarkdown(Workbook workbook, String fileName) throws IOException {
        File file = new File(tempDir, fileName);
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return XlsxToMarkdownConverter.toMarkdown(file);
    }

    private CellStyle boldStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private CellStyle italicStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setItalic(true);
        style.setFont(font);
        return style;
    }

    private CellStyle boldItalicStrikeStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setItalic(true);
        font.setStrikeout(true);
        style.setFont(font);
        return style;
    }
}
