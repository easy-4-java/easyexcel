package com.alibaba.excel.markdown;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.Base64;
import java.util.List;
import java.util.UUID;

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

import com.alibaba.excel.util.FileUtils;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for legacy XLS picture scanning and InputStream spool-to-temp-file behaviour.
 *
 * @author wandl
 */
public class LegacyAndStreamImageTest {

    private static final String PNG_1X1 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    @TempDir
    File tempDir;

    /**
     * Build a .xls workbook with one picture anchored on an empty cell.
     */
    private File writeXlsWithPicture() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("Pic");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Logo");
            header.createCell(1).setCellValue("Name");
            // row 1 stays empty; picture anchored at col 0, row 1
            sheet.createRow(1);

            byte[] png = Base64.getDecoder().decode(PNG_1X1);
            int pictureIndex = workbook.addPicture(png, Workbook.PICTURE_TYPE_PNG);
            HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) 0, 1, (short) 1, 2);
            patriarch.createPicture(anchor, pictureIndex);

            File file = new File(tempDir, "pic.xls");
            try (OutputStream out = Files.newOutputStream(file.toPath())) {
                workbook.write(out);
            }
            return file;
        } finally {
            workbook.close();
        }
    }

    /**
     * Build a .xlsx workbook with one picture anchored on an empty cell.
     */
    private byte[] writeXlsxWithPictureBytes() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("Pic");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Logo");
            header.createCell(1).setCellValue("Name");
            sheet.createRow(1);

            byte[] png = Base64.getDecoder().decode(PNG_1X1);
            int pictureIndex = workbook.addPicture(png, Workbook.PICTURE_TYPE_PNG);
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 1, 1, 2);
            drawing.createPicture(anchor, pictureIndex);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        } finally {
            workbook.close();
        }
    }

    /**
     * Build a simple .xls workbook without pictures (for stream conversion test).
     */
    private byte[] writeSimpleXlsBytes() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("Old");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("A");
            header.createCell(1).setCellValue("B");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue("1");
            data.createCell(1).setCellValue("2");
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        } finally {
            workbook.close();
        }
    }

    // --- Test 1: .xls picture anchor renders placeholder ---

    @Test
    public void xlsPictureAnchorRendersPlaceholder() throws IOException {
        File xlsFile = writeXlsWithPicture();
        SheetDocument document = XlsxToMarkdownConverter.toStructured(xlsFile);

        List<List<String>> rows = document.sheets.get(0).rows;
        // row 0 is the data row (after header promotion) anchored at col 0
        assertThat(rows.get(0).get(0)).isEqualTo("[image]");
    }

    // --- Test 2: InputStream entry scans xlsx pictures ---

    @Test
    public void streamEntryScansXlsxPictures() throws IOException {
        byte[] bytes = writeXlsxWithPictureBytes();
        SheetDocument document = XlsxToMarkdownConverter.toStructured(new ByteArrayInputStream(bytes));

        List<List<String>> rows = document.sheets.get(0).rows;
        assertThat(rows.get(0).get(0)).isEqualTo("[image]");
    }

    // --- Test 3: InputStream entry with .xls converts normally ---

    @Test
    public void streamEntryXlsStillConverts() throws IOException {
        byte[] bytes = writeSimpleXlsBytes();
        SheetDocument document = XlsxToMarkdownConverter.toStructured(new ByteArrayInputStream(bytes));

        assertThat(document.sheets).hasSize(1);
        assertThat(document.sheets.get(0).sheetName).isEqualTo("Old");
        assertThat(document.sheets.get(0).headers.get(0)).containsExactly("A", "B");
        assertThat(document.sheets.get(0).rows.get(0)).containsExactly("1", "2");
    }

    // --- Test 4: temp files are cleaned up after stream entry ---

    @Test
    public void tempFileCleanedUp() throws IOException {
        byte[] bytes = writeXlsxWithPictureBytes();

        // Snapshot: list cache directories before the call
        File cacheRoot = new File(FileUtils.getCachePath());
        int beforeCount = countChildren(cacheRoot);

        XlsxToMarkdownConverter.toStructured(new ByteArrayInputStream(bytes));

        // After: should not have grown (the spool directory is cleaned up)
        int afterCount = countChildren(cacheRoot);
        assertThat(afterCount).isEqualTo(beforeCount);
    }

    private static int countChildren(File dir) {
        if (!dir.exists() || !dir.isDirectory()) {
            return 0;
        }
        String[] children = dir.list();
        return children == null ? 0 : children.length;
    }
}
