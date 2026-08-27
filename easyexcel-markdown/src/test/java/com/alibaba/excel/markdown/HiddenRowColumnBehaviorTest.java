package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Base64;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 验证 {@link XlsxToMarkdownConverter} 对隐藏行和隐藏列的过滤行为。
 * <p>
 * 隐藏的行和列不应出现在 Markdown 输出中。
 *
 * @author wandl
 */
public class HiddenRowColumnBehaviorTest {

    @TempDir
    File tempDir;

    // ==================== XLSX 测试 ====================

    /**
     * XLSX 隐藏行：row 1 标记为 hidden，Markdown 输出应跳过该行。
     */
    @Test
    public void xlsxHiddenRowIsExcludedFromOutput() throws IOException {
        File file = new File(tempDir, "hidden-row.xlsx");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("A");
        header.createCell(1).setCellValue("B");

        Row hiddenRow = sheet.createRow(1);
        hiddenRow.createCell(0).setCellValue("secret");
        hiddenRow.createCell(1).setCellValue("data");
        hiddenRow.setZeroHeight(true);

        Row visibleRow = sheet.createRow(2);
        visibleRow.createCell(0).setCellValue("open");
        visibleRow.createCell(1).setCellValue("info");

        writeAndClose(workbook, file);

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        SheetTable table = document.sheets.get(0);

        // 隐藏行应被过滤
        assertThat(table.rows).hasSize(1);
        assertThat(table.rows.get(0)).containsExactly("open", "info");

        String md = document.toMarkdown();
        assertThat(md).doesNotContain("secret");
        assertThat(md).doesNotContain("data");
        assertThat(md).contains("| open | info |");
    }

    /**
     * XLSX 隐藏列：column B 标记为 hidden，Markdown 输出应跳过该列。
     */
    @Test
    public void xlsxHiddenColumnIsExcludedFromOutput() throws IOException {
        File file = new File(tempDir, "hidden-col.xlsx");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Cols");

        sheet.setColumnHidden(1, true);

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Secret");
        header.createCell(2).setCellValue("Role");

        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("Alice");
        data.createCell(1).setCellValue("s3cret");
        data.createCell(2).setCellValue("Admin");

        writeAndClose(workbook, file);

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        SheetTable table = document.sheets.get(0);

        // 隐藏列应被过滤
        assertThat(table.headers.get(0)).containsExactly("Name", "Role");
        assertThat(table.rows.get(0)).containsExactly("Alice", "Admin");

        String md = document.toMarkdown();
        assertThat(md).doesNotContain("Secret");
        assertThat(md).doesNotContain("s3cret");
        assertThat(md).contains("| Name | Role |");
        assertThat(md).contains("| Alice | Admin |");
    }

    /**
     * XLSX 同时隐藏行和列。
     */
    @Test
    public void xlsxHiddenRowAndColumnCombined() throws IOException {
        File file = new File(tempDir, "hidden-both.xlsx");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Both");

        sheet.setColumnHidden(2, true); // 隐藏第3列

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("C1");
        header.createCell(1).setCellValue("C2");
        header.createCell(2).setCellValue("C3");
        header.createCell(3).setCellValue("C4");

        Row hiddenRow = sheet.createRow(1);
        hiddenRow.createCell(0).setCellValue("h1");
        hiddenRow.createCell(1).setCellValue("h2");
        hiddenRow.createCell(2).setCellValue("h3");
        hiddenRow.createCell(3).setCellValue("h4");
        hiddenRow.setZeroHeight(true);

        Row visible = sheet.createRow(2);
        visible.createCell(0).setCellValue("v1");
        visible.createCell(1).setCellValue("v2");
        visible.createCell(2).setCellValue("v3");
        visible.createCell(3).setCellValue("v4");

        writeAndClose(workbook, file);

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        SheetTable table = document.sheets.get(0);

        assertThat(table.headers.get(0)).containsExactly("C1", "C2", "C4");
        assertThat(table.rows).hasSize(1);
        assertThat(table.rows.get(0)).containsExactly("v1", "v2", "v4");

        String md = document.toMarkdown();
        assertThat(md).doesNotContain("h1");
        assertThat(md).doesNotContain("C3");
        assertThat(md).doesNotContain("v3");
    }

    // ==================== XLS 测试 ====================

    /**
     * XLS 隐藏行：row 1 标记为 hidden，Markdown 输出应跳过该行。
     */
    @Test
    public void xlsHiddenRowIsExcludedFromOutput() throws IOException {
        File file = new File(tempDir, "hidden-row.xls");
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("A");
        header.createCell(1).setCellValue("B");

        Row hiddenRow = sheet.createRow(1);
        hiddenRow.createCell(0).setCellValue("secret");
        hiddenRow.createCell(1).setCellValue("data");
        hiddenRow.setZeroHeight(true);

        Row visibleRow = sheet.createRow(2);
        visibleRow.createCell(0).setCellValue("open");
        visibleRow.createCell(1).setCellValue("info");

        writeAndClose(workbook, file);

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        SheetTable table = document.sheets.get(0);

        assertThat(table.rows).hasSize(1);
        assertThat(table.rows.get(0)).containsExactly("open", "info");

        String md = document.toMarkdown();
        assertThat(md).doesNotContain("secret");
        assertThat(md).contains("| open | info |");
    }

    /**
     * XLS 隐藏列：column B 标记为 hidden，Markdown 输出应跳过该列。
     */
    @Test
    public void xlsHiddenColumnIsExcludedFromOutput() throws IOException {
        File file = new File(tempDir, "hidden-col.xls");
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Cols");

        sheet.setColumnHidden(1, true);

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Secret");
        header.createCell(2).setCellValue("Role");

        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("Alice");
        data.createCell(1).setCellValue("s3cret");
        data.createCell(2).setCellValue("Admin");

        writeAndClose(workbook, file);

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        SheetTable table = document.sheets.get(0);

        assertThat(table.headers.get(0)).containsExactly("Name", "Role");
        assertThat(table.rows.get(0)).containsExactly("Alice", "Admin");

        String md = document.toMarkdown();
        assertThat(md).doesNotContain("Secret");
        assertThat(md).doesNotContain("s3cret");
    }

    /**
     * XLS 同时隐藏行和列。
     */
    @Test
    public void xlsHiddenRowAndColumnCombined() throws IOException {
        File file = new File(tempDir, "hidden-both.xls");
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Both");

        sheet.setColumnHidden(2, true);

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("C1");
        header.createCell(1).setCellValue("C2");
        header.createCell(2).setCellValue("C3");
        header.createCell(3).setCellValue("C4");

        Row hiddenRow = sheet.createRow(1);
        hiddenRow.createCell(0).setCellValue("h1");
        hiddenRow.createCell(1).setCellValue("h2");
        hiddenRow.createCell(2).setCellValue("h3");
        hiddenRow.createCell(3).setCellValue("h4");
        hiddenRow.setZeroHeight(true);

        Row visible = sheet.createRow(2);
        visible.createCell(0).setCellValue("v1");
        visible.createCell(1).setCellValue("v2");
        visible.createCell(2).setCellValue("v3");
        visible.createCell(3).setCellValue("v4");

        writeAndClose(workbook, file);

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        SheetTable table = document.sheets.get(0);

        assertThat(table.headers.get(0)).containsExactly("C1", "C2", "C4");
        assertThat(table.rows).hasSize(1);
        assertThat(table.rows.get(0)).containsExactly("v1", "v2", "v4");
    }

    /**
     * 没有隐藏行列时，输出与之前一致（回归测试）。
     */
    @Test
    public void noHiddenRowsOrColumnsUnchanged() throws IOException {
        File file = new File(tempDir, "no-hidden.xlsx");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Normal");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("X");
        header.createCell(1).setCellValue("Y");

        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("1");
        data.createCell(1).setCellValue("2");

        writeAndClose(workbook, file);

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        assertThat(document.sheets.get(0).headers.get(0)).containsExactly("X", "Y");
        assertThat(document.sheets.get(0).rows.get(0)).containsExactly("1", "2");
    }

    // ==================== 图片锚点 + 隐藏行列 测试 ====================

    private static final String PNG_1X1 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    /**
     * XLSX：图片锚点位于隐藏行上时，不应输出 [image] 占位符。
     */
    @Test
    public void xlsxImageOnHiddenRowIsSkipped() throws IOException {
        File file = new File(tempDir, "img-hidden-row.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("Pic");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Logo");
            header.createCell(1).setCellValue("Name");

            Row hiddenRow = sheet.createRow(1);
            hiddenRow.createCell(0).setCellValue("hidden-val");
            hiddenRow.setZeroHeight(true);

            // 图片锚定在隐藏行1的col 0
            byte[] png = Base64.getDecoder().decode(PNG_1X1);
            int pictureIndex = workbook.addPicture(png, Workbook.PICTURE_TYPE_PNG);
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 1, 1, 2);
            drawing.createPicture(anchor, pictureIndex);

            writeAndClose(workbook, file);
        } catch (Exception e) {
            workbook.close();
            throw e;
        }

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        String md = document.toMarkdown();
        // 隐藏行上的图片不应出现
        assertThat(md).doesNotContain("[image]");
        assertThat(md).doesNotContain("hidden-val");
    }

    /**
     * XLSX：图片锚点位于隐藏列上时，不应输出 [image] 占位符。
     */
    @Test
    public void xlsxImageOnHiddenColumnIsSkipped() throws IOException {
        File file = new File(tempDir, "img-hidden-col.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("Pic");
            sheet.setColumnHidden(0, true); // 隐藏第1列

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Logo");
            header.createCell(1).setCellValue("Name");

            sheet.createRow(1);

            // 图片锚定在col 0（隐藏列）
            byte[] png = Base64.getDecoder().decode(PNG_1X1);
            int pictureIndex = workbook.addPicture(png, Workbook.PICTURE_TYPE_PNG);
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 1, 1, 2);
            drawing.createPicture(anchor, pictureIndex);

            writeAndClose(workbook, file);
        } catch (Exception e) {
            workbook.close();
            throw e;
        }

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        String md = document.toMarkdown();
        assertThat(md).doesNotContain("[image]");
    }

    /**
     * XLS：图片锚点位于隐藏行上时，不应输出 [image] 占位符。
     */
    @Test
    public void xlsImageOnHiddenRowIsSkipped() throws IOException {
        File file = new File(tempDir, "img-hidden-row.xls");
        HSSFWorkbook workbook = new HSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("Pic");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Logo");
            header.createCell(1).setCellValue("Name");

            Row hiddenRow = sheet.createRow(1);
            hiddenRow.setZeroHeight(true);

            byte[] png = Base64.getDecoder().decode(PNG_1X1);
            int pictureIndex = workbook.addPicture(png, Workbook.PICTURE_TYPE_PNG);
            HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) 0, 1, (short) 1, 2);
            patriarch.createPicture(anchor, pictureIndex);

            writeAndClose(workbook, file);
        } catch (Exception e) {
            workbook.close();
            throw e;
        }

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        String md = document.toMarkdown();
        assertThat(md).doesNotContain("[image]");
    }

    private void writeAndClose(Workbook workbook, File file) throws IOException {
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
    }
}
