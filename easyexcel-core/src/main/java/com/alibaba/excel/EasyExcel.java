package com.alibaba.excel;

import java.io.File;
import java.io.InputStream;

import com.alibaba.excel.markdown.SheetDocument;
import com.alibaba.excel.markdown.XlsxToMarkdownConverter;

/**
 * This is actually {@link EasyExcelFactory}, and short names look better.
 *
 * @author jipengfei
 */
public class EasyExcel extends EasyExcelFactory {

    /**
     * Read an XLSX / XLS / CSV workbook and keep every sheet as a structured
     * {@link SheetDocument}, see {@link XlsxToMarkdownConverter#toStructured(File)}.
     *
     * @param file
     *            the workbook file
     * @return the structured document with all sheets
     */
    public static SheetDocument xlsxToStructured(File file) {
        return XlsxToMarkdownConverter.toStructured(file);
    }

    /**
     * Read an XLSX / XLS / CSV workbook from a stream and keep every sheet as a structured
     * {@link SheetDocument}, see {@link XlsxToMarkdownConverter#toStructured(InputStream)}.
     *
     * @param inputStream
     *            the workbook content
     * @return the structured document with all sheets
     */
    public static SheetDocument xlsxToStructured(InputStream inputStream) {
        return XlsxToMarkdownConverter.toStructured(inputStream);
    }

    /**
     * Convert an XLSX / XLS / CSV workbook to multi-sheet Markdown,
     * see {@link XlsxToMarkdownConverter#toMarkdown(File)}.
     *
     * @param file
     *            the workbook file
     * @return the Markdown text
     */
    public static String xlsxToMarkdown(File file) {
        return XlsxToMarkdownConverter.toMarkdown(file);
    }

    /**
     * Convert an XLSX / XLS / CSV workbook read from a stream to multi-sheet Markdown,
     * see {@link XlsxToMarkdownConverter#toMarkdown(InputStream)}.
     *
     * @param inputStream
     *            the workbook content
     * @return the Markdown text
     */
    public static String xlsxToMarkdown(InputStream inputStream) {
        return XlsxToMarkdownConverter.toMarkdown(inputStream);
    }
}
