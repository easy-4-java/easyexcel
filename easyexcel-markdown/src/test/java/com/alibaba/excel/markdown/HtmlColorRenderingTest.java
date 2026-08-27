package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * HTML 颜色渲染测试：开关控制、字体色、背景色、HTML 转义、XLS 路径。
 * 每个测试用 @AfterEach 恢复默认开关状态，避免污染其他测试。
 *
 * @author wandl
 */
public class HtmlColorRenderingTest {

    private static final byte MAX_COLOR_VALUE = (byte) 255;

    @TempDir
    File tempDir;

    @AfterEach
    void resetColorRendering() {
        XlsxToMarkdownConverter.setHtmlColorRendering(false);
    }

    // --------------------------------------------------------------- XLSX 字体色

    @Test
    public void xlsxFontColorRenderedWhenEnabled() throws IOException {
        XlsxToMarkdownConverter.setHtmlColorRendering(true);
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Color");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Col");

        XSSFFont redFont = wb.createFont();
        XSSFColor redColor = new XSSFColor();
        redColor.setRGB(new byte[]{MAX_COLOR_VALUE, 0, 0});
        redFont.setColor(redColor);

        XSSFCellStyle cs = wb.createCellStyle();
        cs.setFont(redFont);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue("red");
        cell.setCellStyle(cs);

        String markdown = convertToMarkdown(wb, "font_color.xlsx");
        assertThat(markdown).contains("<span style=\"color:#FF0000\">red</span>");
    }

    // --------------------------------------------------------------- XLSX 背景色

    @Test
    public void xlsxBackgroundColorRenderedWhenEnabled() throws IOException {
        XlsxToMarkdownConverter.setHtmlColorRendering(true);
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("BgColor");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Col");

        XSSFCellStyle cs = wb.createCellStyle();
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFColor bgColor = new XSSFColor();
        bgColor.setRGB(new byte[]{MAX_COLOR_VALUE, MAX_COLOR_VALUE, 0});
        cs.setFillForegroundColor(bgColor);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue("bg");
        cell.setCellStyle(cs);

        String markdown = convertToMarkdown(wb, "bg_color.xlsx");
        assertThat(markdown).contains("background-color:#FFFF00");
        assertThat(markdown).contains("<span style=\"background-color:#FFFF00\">bg</span>");
    }

    // --------------------------------------------------------------- 默认关闭

    @Test
    public void xlsxColorOffByDefault() throws IOException {
        // 不调用 setHtmlColorRendering，默认 false
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Default");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Col");

        XSSFFont redFont = wb.createFont();
        XSSFColor redColor = new XSSFColor();
        redColor.setRGB(new byte[]{MAX_COLOR_VALUE, 0, 0});
        redFont.setColor(redColor);

        XSSFCellStyle cs = wb.createCellStyle();
        cs.setFont(redFont);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue("no_color");
        cell.setCellStyle(cs);

        String markdown = convertToMarkdown(wb, "default_off.xlsx");

        // 默认关闭时不应有任何 span 标签
        assertThat(markdown).doesNotContain("<span");
        assertThat(markdown).doesNotContain("</span>");
    }

    // --------------------------------------------------------------- HTML 转义

    @Test
    public void htmlSpecialCharsEscaped() throws IOException {
        XlsxToMarkdownConverter.setHtmlColorRendering(true);
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Escape");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Col");

        XSSFFont redFont = wb.createFont();
        XSSFColor redColor = new XSSFColor();
        redColor.setRGB(new byte[]{MAX_COLOR_VALUE, 0, 0});
        redFont.setColor(redColor);

        XSSFCellStyle cs = wb.createCellStyle();
        cs.setFont(redFont);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        // 值含 < 和 &，应被转义且不破坏 span 结构
        cell.setCellValue("a<b&c");
        cell.setCellStyle(cs);

        String markdown = convertToMarkdown(wb, "escape.xlsx");

        // 转义顺序：先 marker -> 再转义 -> 再包 span
        assertThat(markdown).contains("&lt;");
        assertThat(markdown).contains("&amp;");
        // span 结构完整
        assertThat(markdown).contains("<span style=\"color:#FF0000\">");
        assertThat(markdown).contains("</span>");
    }

    // --------------------------------------------------------------- XLS 路径

    @Test
    public void xlsColorRenderedWhenEnabled() throws IOException {
        XlsxToMarkdownConverter.setHtmlColorRendering(true);
        HSSFWorkbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("XlsColor");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Font");
        header.createCell(1).setCellValue("Bg");

        // 字体色：通过自定义 palette 设置红色
        wb.getCustomPalette().setColorAtIndex(
            HSSFFont.COLOR_RED, MAX_COLOR_VALUE, (byte) 0, (byte) 0);
        HSSFFont redFont = wb.createFont();
        redFont.setColor(HSSFFont.COLOR_RED);

        org.apache.poi.ss.usermodel.CellStyle fontCs = wb.createCellStyle();
        fontCs.setFont(redFont);

        // 背景色：SOLID_FOREGROUND + 绿色
        org.apache.poi.ss.usermodel.CellStyle bgCs = wb.createCellStyle();
        bgCs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        bgCs.setFillForegroundColor(IndexedColors.GREEN.getIndex());

        Row data = sheet.createRow(1);
        Cell fontCell = data.createCell(0);
        fontCell.setCellValue("red_font");
        fontCell.setCellStyle(fontCs);

        Cell bgCell = data.createCell(1);
        bgCell.setCellValue("green_bg");
        bgCell.setCellStyle(bgCs);

        String markdown = convertToMarkdown(wb, "xls_color.xls");

        // 字体色应包含 color:
        assertThat(markdown).contains("color:");
        // 背景色应包含 background-color:
        assertThat(markdown).contains("background-color:");
        assertThat(markdown).contains("<span");
        assertThat(markdown).contains("</span>");
    }

    // --------------------------------------------------------------- 字体色+背景色组合

    @Test
    public void xlsxFontAndBgColorCombined() throws IOException {
        XlsxToMarkdownConverter.setHtmlColorRendering(true);
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Combo");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Col");

        XSSFFont blueFont = wb.createFont();
        XSSFColor blueColor = new XSSFColor();
        blueColor.setRGB(new byte[]{0, 0, MAX_COLOR_VALUE});
        blueFont.setColor(blueColor);

        XSSFCellStyle cs = wb.createCellStyle();
        cs.setFont(blueFont);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFColor bgColor = new XSSFColor();
        bgColor.setRGB(new byte[]{MAX_COLOR_VALUE, MAX_COLOR_VALUE, MAX_COLOR_VALUE});
        cs.setFillForegroundColor(bgColor);

        Row data = sheet.createRow(1);
        Cell cell = data.createCell(0);
        cell.setCellValue("combo");
        cell.setCellStyle(cs);

        String markdown = convertToMarkdown(wb, "combo_color.xlsx");

        // 两个颜色都应出现，color 在前 background-color 在后
        assertThat(markdown).contains("color:#0000FF");
        assertThat(markdown).contains("background-color:#FFFFFF");
        assertThat(markdown).contains(
            "<span style=\"color:#0000FF;background-color:#FFFFFF\">combo</span>");
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
