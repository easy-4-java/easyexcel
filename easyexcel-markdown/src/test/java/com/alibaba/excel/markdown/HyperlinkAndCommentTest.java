package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test hyperlink and comment rendering end to end.
 *
 * @author wandl
 */
public class HyperlinkAndCommentTest {

    @TempDir
    File tempDir;

    private Workbook hyperlinkAndCommentWorkbook() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Link");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Site");
        header.createCell(1).setCellValue("Note");
        Row data = sheet.createRow(1);

        CreationHelper helper = workbook.getCreationHelper();
        Hyperlink hyperlink = helper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress("https://github.com/alibaba/easyexcel");
        Cell linkCell = data.createCell(0);
        linkCell.setCellValue("easyexcel");
        linkCell.setHyperlink(hyperlink);

        Cell commentCell = data.createCell(1);
        commentCell.setCellValue("v4.0.3");
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(1);
        anchor.setCol2(2);
        anchor.setRow1(1);
        anchor.setRow2(2);
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(helper.createRichTextString("line1\nline2"));
        commentCell.setCellComment(comment);
        return workbook;
    }

    private File convertToFile(Workbook workbook) throws IOException {
        File file = new File(tempDir, "link.xlsx");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return file;
    }

    @Test
    public void renderHyperlinkAsMarkdownLink() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(convertToFile(hyperlinkAndCommentWorkbook()));

        List<String> row = document.sheets.get(0).rows.get(0);
        assertThat(row.get(0)).isEqualTo("[easyexcel](https://github.com/alibaba/easyexcel)");
    }

    @Test
    public void renderCommentWithFoldedNewlines() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(convertToFile(hyperlinkAndCommentWorkbook()));

        List<String> row = document.sheets.get(0).rows.get(0);
        assertThat(row.get(1)).isEqualTo("v4.0.3 <!-- line1 line2 -->");
    }

    @Test
    public void markdownContainsRenderedExtras() throws IOException {
        assertThat(XlsxToMarkdownConverter.toMarkdown(convertToFile(hyperlinkAndCommentWorkbook())))
            .contains("[easyexcel](https://github.com/alibaba/easyexcel)")
            .contains("<!-- line1 line2 -->");
    }

    @Test
    public void urlWithSpacesIsAngleWrapped() {
        assertThat(XlsxToMarkdownConverter.markdownLink("doc", "https://example.com/a b"))
            .isEqualTo("[doc](<https://example.com/a b>)");
        assertThat(XlsxToMarkdownConverter.markdownLink("", "https://example.com"))
            .isEqualTo("[https://example.com](https://example.com)");
    }
}
