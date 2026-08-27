package com.alibaba.excel.markdown;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Hyperlink;
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
 * Test the streaming {@link Writer} overloads of {@link XlsxToMarkdownConverter}.
 * Every test compares the writer-based output against the existing String-based API
 * to guarantee byte-for-byte equivalence.
 *
 * @author wandl
 */
public class StreamingMarkdownTest {

    private static final char UTF8_BOM = '\uFEFF';
    private static final String CSV_SEMICOLON_CONTENT = "h1;h2\nv1;v2\n";

    @TempDir
    File tempDir;

    private Workbook twoSheetWorkbook() {
        Workbook workbook = new XSSFWorkbook();

        // Sheet 0: "Alpha" with header, data, merged region, hyperlink, comment
        Sheet alpha = workbook.createSheet("Alpha");
        Row header0 = alpha.createRow(0);
        header0.createCell(0).setCellValue("Name");
        header0.createCell(1).setCellValue("Age");
        header0.createCell(2).setCellValue("Link");

        Row data0 = alpha.createRow(1);
        data0.createCell(0).setCellValue("Zhang San");
        data0.createCell(1).setCellValue(20);

        // Hyperlink on cell C2
        CreationHelper helper = workbook.getCreationHelper();
        Hyperlink hyperlink = helper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress("https://example.com");
        Cell linkCell = data0.createCell(2);
        linkCell.setCellValue("click");
        linkCell.setHyperlink(hyperlink);

        // Comment on cell B2
        Cell commentCell = data0.getCell(1);
        Drawing<?> drawing = alpha.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(1);
        anchor.setCol2(2);
        anchor.setRow1(1);
        anchor.setRow2(2);
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(helper.createRichTextString("a note"));
        commentCell.setCellComment(comment);

        Row data1 = alpha.createRow(2);
        data1.createCell(0).setCellValue("Li Si");
        data1.createCell(1).setCellValue(30);
        data1.createCell(2).setCellValue("plain");

        // Merged region: A1:B1 (header row)
        alpha.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));

        // Sheet 1: "Beta" with simple data
        Sheet beta = workbook.createSheet("Beta");
        Row header1 = beta.createRow(0);
        header1.createCell(0).setCellValue("Key");
        Row data2 = beta.createRow(1);
        data2.createCell(0).setCellValue("v1");

        return workbook;
    }

    private File writeWorkbook(Workbook workbook, String name) throws IOException {
        File file = new File(tempDir, name);
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return file;
    }

    private File writeCsv(String content) throws IOException {
        File file = new File(tempDir, "sample.csv");
        Files.write(file.toPath(), content.getBytes(StandardCharsets.UTF_8));
        return file;
    }

    @Test
    public void writerOutputMatchesStringApi() throws IOException {
        File file = writeWorkbook(twoSheetWorkbook(), "multi.xlsx");

        String expected = XlsxToMarkdownConverter.toMarkdown(file);
        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(file, sw);
        String actual = sw.toString();

        assertThat(actual).isEqualTo(expected);
    }

    @Test
    public void csvStreamMatches() throws IOException {
        // CSV with BOM and semicolon delimiter
        File file = writeCsv(UTF8_BOM + CSV_SEMICOLON_CONTENT);

        String expected = XlsxToMarkdownConverter.toMarkdown(file);
        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(file, sw);
        String actual = sw.toString();

        assertThat(actual).isEqualTo(expected);
    }

    @Test
    public void inputStreamWriterVariant() throws IOException {
        File file = writeWorkbook(twoSheetWorkbook(), "stream.xlsx");
        byte[] bytes = Files.readAllBytes(file.toPath());

        String expected = XlsxToMarkdownConverter.toMarkdown(new ByteArrayInputStream(bytes));
        StringWriter sw = new StringWriter();
        XlsxToMarkdownConverter.toMarkdown(new ByteArrayInputStream(bytes), sw);
        String actual = sw.toString();

        assertThat(actual).isEqualTo(expected);
    }

    @Test
    public void writerIsNotClosedAndIsFlushed() throws IOException {
        File file = writeWorkbook(twoSheetWorkbook(), "flush.xlsx");
        TrackingWriter tw = new TrackingWriter(new StringWriter());

        XlsxToMarkdownConverter.toMarkdown(file, tw);

        assertThat(tw.closed).isFalse();
        assertThat(tw.flushed).isTrue();
    }

    @Test
    public void nullArgumentsRejected() throws IOException {
        File file = writeWorkbook(twoSheetWorkbook(), "null.xlsx");

        assertThatThrownBy(() -> XlsxToMarkdownConverter.toMarkdown((File) null, new StringWriter()))
            .isInstanceOf(NullPointerException.class)
            .hasMessageContaining("file must not be null");

        assertThatThrownBy(() -> XlsxToMarkdownConverter.toMarkdown(file, null))
            .isInstanceOf(NullPointerException.class)
            .hasMessageContaining("writer must not be null");

        assertThatThrownBy(() -> XlsxToMarkdownConverter.toMarkdown((java.io.InputStream) null, new StringWriter()))
            .isInstanceOf(NullPointerException.class)
            .hasMessageContaining("inputStream must not be null");
    }

    /**
     * Writer wrapper that tracks whether close() and flush() were called.
     */
    private static final class TrackingWriter extends Writer {
        private final Writer delegate;
        boolean closed;
        boolean flushed;

        TrackingWriter(Writer delegate) {
            this.delegate = delegate;
        }

        @Override
        public void write(char[] cbuf, int off, int len) throws IOException {
            delegate.write(cbuf, off, len);
        }

        @Override
        public void flush() throws IOException {
            flushed = true;
            delegate.flush();
        }

        @Override
        public void close() throws IOException {
            closed = true;
            delegate.close();
        }
    }
}
