package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Base64;
import java.util.List;

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
 * Test that cells carrying an anchored picture render as a placeholder instead of vanishing.
 * The easyexcel SAX read path cannot see drawings, so the converter scans sheet drawing parts
 * itself and marks the anchor cell of every picture.
 *
 * @author wandl
 */
public class ImagePlaceholderTest {

    private static final String PNG_1X1 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    @TempDir
    File tempDir;

    private Workbook pictureWorkbook() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Pic");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Logo");
        header.createCell(1).setCellValue("Name");
        // row 1 stays without any cell value, the picture is anchored on that empty position
        sheet.createRow(1);

        byte[] png = Base64.getDecoder().decode(PNG_1X1);
        int pictureIndex = workbook.addPicture(png, Workbook.PICTURE_TYPE_PNG);
        XSSFDrawing drawing = (XSSFDrawing)sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 1, 1, 2);
        drawing.createPicture(anchor, pictureIndex);
        return workbook;
    }

    private File convertToFile(Workbook workbook) throws IOException {
        File file = new File(tempDir, "pic.xlsx");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return file;
    }

    @Test
    public void pictureOverEmptyCellRendersPlaceholder() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(convertToFile(pictureWorkbook()));

        List<List<String>> rows = document.sheets.get(0).rows;
        assertThat(rows.get(0)).containsExactly("[image]");
    }

    @Test
    public void pictureAnchoredOverValuedCellKeepsTheValue() throws IOException {
        Workbook workbook = pictureWorkbook();
        Sheet sheet = workbook.getSheet("Pic");
        Row valued = sheet.getRow(1);
        if (valued == null) {
            valued = sheet.createRow(1);
        }
        valued.createCell(1).setCellValue("kept");

        SheetDocument document = XlsxToMarkdownConverter.toStructured(convertToFile(workbook));

        assertThat(document.sheets.get(0).rows.get(0)).containsExactly("[image]", "kept");
    }

    @Test
    public void markdownContainsPlaceholder() throws IOException {
        assertThat(XlsxToMarkdownConverter.toMarkdown(convertToFile(pictureWorkbook())))
            .contains("| [image] |  |");
    }
}
