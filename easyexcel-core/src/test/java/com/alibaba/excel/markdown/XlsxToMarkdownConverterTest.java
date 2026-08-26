package com.alibaba.excel.markdown;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.write.metadata.WriteSheet;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Test {@link XlsxToMarkdownConverter} against fixtures generated on the fly: XLSX through the
 * EasyExcel write path, XLS and merged regions through POI, CSV through raw bytes.
 *
 * @author wandl
 */
public class XlsxToMarkdownConverterTest {

    @TempDir
    Path tempDir;

    private static List<List<String>> head(String... names) {
        List<List<String>> head = new ArrayList<>();
        for (String name : names) {
            head.add(Collections.singletonList(name));
        }
        return head;
    }

    private File writeMultiSheetXlsx() throws IOException {
        File file = tempDir.resolve("multi.xlsx").toFile();
        ExcelWriter writer = EasyExcel.write(file).build();
        try {
            WriteSheet alpha = EasyExcel.writerSheet(0, "Alpha").head(head("Name", "Age")).build();
            writer.write(Arrays.<List<Object>>asList(
                Arrays.asList("Zhang San", 20),
                Arrays.asList("only-a", null, "c")), alpha);
            WriteSheet beta = EasyExcel.writerSheet(1, "Beta").head(head("Key")).build();
            writer.write(Arrays.<List<Object>>asList(
                Arrays.asList("v1")), beta);
        } finally {
            writer.finish();
        }
        return file;
    }

    private File writeMergedXlsx() throws IOException {
        File file = tempDir.resolve("merged.xlsx").toFile();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Merged");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Span");
        header.createCell(1);
        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("x");
        data.createCell(1).setCellValue("y");
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return file;
    }

    private File writeXls() throws IOException {
        File file = tempDir.resolve("legacy.xls").toFile();
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Old");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("A");
        header.createCell(1).setCellValue("B");
        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("1");
        data.createCell(1).setCellValue("2");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return file;
    }

    private File writeCsv(String content) throws IOException {
        File file = tempDir.resolve("sample.csv").toFile();
        Files.write(file.toPath(), content.getBytes(StandardCharsets.UTF_8));
        return file;
    }

    @Test
    public void convertMultiSheetXlsx() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(writeMultiSheetXlsx());

        assertThat(document.title).isEqualTo("multi.xlsx");
        assertThat(document.sheets).hasSize(2);

        SheetTable alpha = document.sheets.get(0);
        assertThat(alpha.sheetName).isEqualTo("Alpha");
        assertThat(alpha.headers).containsExactly(Arrays.asList("Name", "Age"));
        assertThat(alpha.rows).containsExactly(
            Arrays.asList("Zhang San", "20"),
            Arrays.asList("only-a", "", "c"));

        SheetTable beta = document.sheets.get(1);
        assertThat(beta.sheetName).isEqualTo("Beta");
        assertThat(beta.headers).containsExactly(Arrays.asList("Key"));
        assertThat(beta.rows).containsExactly(Arrays.asList("v1"));

        assertThat(document.toMarkdown())
            .contains("## Alpha")
            .contains("| Name | Age |  |")
            .contains("| only-a |  | c |")
            .contains("## Beta");
    }

    @Test
    public void fillMergedRegionValues() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(writeMergedXlsx());

        assertThat(document.sheets).hasSize(1);
        SheetTable merged = document.sheets.get(0);
        assertThat(merged.sheetName).isEqualTo("Merged");
        assertThat(merged.headers).containsExactly(Arrays.asList("Span", "Span"));
        assertThat(merged.rows).containsExactly(Arrays.asList("x", "y"));
    }

    @Test
    public void convertXls() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(writeXls());

        assertThat(document.sheets).hasSize(1);
        assertThat(document.sheets.get(0).sheetName).isEqualTo("Old");
        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("A", "B"));
        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("1", "2"));
    }

    @Test
    public void convertCsvWithBomAndQuotedComma() throws IOException {
        File file = writeCsv("\uFEFFh1,h2\nv1,\"b,c\"\n");

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);

        assertThat(document.title).isEqualTo("sample.csv");
        assertThat(document.sheets).hasSize(1);
        SheetTable table = document.sheets.get(0);
        assertThat(table.sheetName).isEqualTo("sample");
        assertThat(table.headers).containsExactly(Arrays.asList("h1", "h2"));
        assertThat(table.rows).containsExactly(Arrays.asList("v1", "b,c"));
    }

    @Test
    public void convertFromInputStream() throws IOException {
        byte[] xlsx = Files.readAllBytes(writeMultiSheetXlsx().toPath());
        SheetDocument fromXlsx = XlsxToMarkdownConverter.toStructured(new ByteArrayInputStream(xlsx));
        assertThat(fromXlsx.sheets).hasSize(2);
        assertThat(fromXlsx.sheets.get(0).sheetName).isEqualTo("Alpha");

        byte[] csv = "h1,h2\nv1,v2\n".getBytes(StandardCharsets.UTF_8);
        SheetDocument fromCsv = XlsxToMarkdownConverter.toStructured(new ByteArrayInputStream(csv));
        assertThat(fromCsv.sheets.get(0).headers).containsExactly(Arrays.asList("h1", "h2"));
        assertThat(fromCsv.sheets.get(0).rows).containsExactly(Arrays.asList("v1", "v2"));
    }

    @Test
    public void escapeMarkdownEndToEnd() throws IOException {
        File file = tempDir.resolve("escape.xlsx").toFile();
        ExcelWriter writer = EasyExcel.write(file).head(head("a|b")).build();
        try {
            writer.write(Arrays.<List<Object>>asList(
                Arrays.asList("line1\nline2")), EasyExcel.writerSheet(0, "Esc").build());
        } finally {
            writer.finish();
        }

        assertThat(XlsxToMarkdownConverter.toMarkdown(file))
            .contains("| a\\|b |")
            .contains("| line1<br>line2 |");
    }

    @Test
    public void rejectNullAndMissingFile() {
        assertThatThrownBy(() -> XlsxToMarkdownConverter.toStructured((File)null))
            .isInstanceOf(NullPointerException.class)
            .hasMessageContaining("file must not be null");
        assertThatThrownBy(() -> XlsxToMarkdownConverter.toStructured((java.io.InputStream)null))
            .isInstanceOf(NullPointerException.class)
            .hasMessageContaining("inputStream must not be null");
        assertThatThrownBy(() -> XlsxToMarkdownConverter.toStructured(new File("no-such-file.xlsx")))
            .isInstanceOf(ExcelAnalysisException.class)
            .hasMessageContaining("File not found");
    }
}
