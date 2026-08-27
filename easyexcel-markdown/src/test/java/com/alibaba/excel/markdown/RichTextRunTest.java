package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for rich text run-level inline styling: partial bold, uniform runs,
 * italic+strike combos, plain string isolation, and xls rich text fallback.
 *
 * @author wandl
 */
public class RichTextRunTest {

    private static final int INITIAL_CAPACITY = 4;

    @TempDir
    File tempDir;

    // --------------------------------------------------------------- XLSX

    @Test
    public void partialBoldSplitsRuns() throws IOException {
        // "a"(bold) + "b"(plain) -> markdown cell is **a**b
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Rich");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("H");

        XSSFFont boldFont = wb.createFont();
        boldFont.setBold(true);
        XSSFRichTextString rt = new XSSFRichTextString("ab");
        rt.applyFont(0, 1, boldFont);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue(rt);

        String markdown = convertToMarkdown(wb, "partial_bold.xlsx");
        assertThat(markdown).contains("**a**b");
    }

    @Test
    public void uniformRunsEqualWholeWrap() throws IOException {
        // Two runs both bold -> **ab** (not **a****b**)
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Rich");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("H");

        XSSFFont boldFont = wb.createFont();
        boldFont.setBold(true);
        XSSFRichTextString rt = new XSSFRichTextString("ab");
        rt.applyFont(0, 1, boldFont);
        rt.applyFont(1, 2, boldFont);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue(rt);

        String markdown = convertToMarkdown(wb, "uniform_bold.xlsx");
        assertThat(markdown).contains("**ab**");
        assertThat(markdown).doesNotContain("**a****b**");
    }

    @Test
    public void richTextWithItalicAndStrikeCombos() throws IOException {
        // "x"(italic+strike) + "y"(plain) -> ~*x*~y
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Rich");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("H");

        XSSFFont comboFont = wb.createFont();
        comboFont.setItalic(true);
        comboFont.setStrikeout(true);
        XSSFRichTextString rt = new XSSFRichTextString("xy");
        rt.applyFont(0, 1, comboFont);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue(rt);

        String markdown = convertToMarkdown(wb, "italic_strike.xlsx");
        // Combination order: strike wraps italic: ~~*x*~~
        assertThat(markdown).contains("~~*x*~~y");
    }

    @Test
    public void plainStringUnaffected() throws IOException {
        // Plain string cell should have no markers (existing behavior byte-for-byte)
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Plain");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Col");

        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("hello");

        String markdown = convertToMarkdown(wb, "plain.xlsx");
        assertThat(markdown).contains("| hello |");
        assertThat(markdown).doesNotContain("**hello**");
        assertThat(markdown).doesNotContain("*hello*");
        assertThat(markdown).doesNotContain("~~hello~~");
    }

    // --------------------------------------------------------------- XLS

    @Test
    public void xlsRichTextFallback() throws IOException {
        // XLS rich text: "a"(bold) + "b"(plain)
        // If HSSFRichTextString run extraction works -> **a**b
        // If not feasible -> falls back to whole-cell style
        HSSFWorkbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Rich");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("H");

        HSSFFont boldFont = wb.createFont();
        boldFont.setBold(true);
        HSSFRichTextString rt = new HSSFRichTextString("ab");
        rt.applyFont(0, 1, boldFont);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue(rt);

        String markdown = convertToMarkdown(wb, "partial_bold.xls");

        // Verify run-level extraction feasibility:
        // numFormattingRuns() should be > 0 for a rich text string with partial formatting
        HSSFWorkbook verify = new HSSFWorkbook();
        HSSFFont vBold = verify.createFont();
        vBold.setBold(true);
        HSSFRichTextString vRt = new HSSFRichTextString("ab");
        vRt.applyFont(0, 1, vBold);
        int formattingRunCount = vRt.numFormattingRuns();
        verify.close();

        if (formattingRunCount > 0) {
            // Run-level extraction is feasible: expect **a**b
            assertThat(markdown).contains("**a**b");
        } else {
            // Fallback: whole-cell treatment. The cell-level font is the default
            // (not bold), so plain rendering is expected.
            assertThat(markdown).contains("| ab |");
        }
    }

    // --------------------------------------------------------------- helpers

    private String convertToMarkdown(Workbook workbook, String fileName) throws IOException {
        File file = new File(tempDir, fileName);
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return XlsxToMarkdownConverter.toMarkdown(file);
    }
}
