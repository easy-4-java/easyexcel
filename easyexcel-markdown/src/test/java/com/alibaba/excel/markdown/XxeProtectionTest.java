package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * XXE 防护测试：恶意 xlsx 的 sheet XML / styles.xml 中内嵌 DOCTYPE 外部实体引用时，
 * 转换必须被拒绝（SAX 解析器禁用 DOCTYPE 声明），而不是解析外部实体（读本地文件 / SSRF）。
 *
 * @author wandl
 */
public class XxeProtectionTest {

    @TempDir
    File tempDir;

    /**
     * 构造一个最小的 xlsx 包，其中 sheet XML 或 styles.xml 可携带 DOCTYPE 外部实体引用。
     * 手工写 zip（POI 不会生成 DOCTYPE），模拟恶意工作簿。
     * <p>
     * 注意：workbook.xml.rels 必须包含 styles relationship，否则
     * {@code XSSFReader.getStylesData()} 返回 null、styles.xml 不会被扫描器解析，
     * styles 路径的 XXE 测试就无法命中防护（这是本测试最容易写漏的点）。
     */
    private File maliciousWorkbook(String sheetXml, String stylesXml) throws IOException {
        File file = new File(tempDir, "xxe-" + System.nanoTime() + ".xlsx");
        try (ZipOutputStream zip = new ZipOutputStream(new FileOutputStream(file))) {
            put(zip, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                    + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
                    + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                    + "</Types>");
            put(zip, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                    + "</Relationships>");
            put(zip, "xl/workbook.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                    + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                    + "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            put(zip, "xl/_rels/workbook.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                    + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
                    + "</Relationships>");
            put(zip, "xl/styles.xml", stylesXml);
            put(zip, "xl/worksheets/sheet1.xml", sheetXml);
        }
        return file;
    }

    private void put(ZipOutputStream zip, String name, String content) throws IOException {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(content.getBytes("UTF-8"));
        zip.closeEntry();
    }

    private static final String NORMAL_SHEET =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>hello</t></is></c></row></sheetData>"
            + "</worksheet>";

    private static final String NORMAL_STYLES =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
            + "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>"
            + "<borders count=\"1\"><border/></borders>"
            + "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>"
            + "<cellXfs count=\"1\"><xf xfId=\"0\"/></cellXfs>"
            + "</styleSheet>";

    private static final String XXE_DOCTYPE =
        "<!DOCTYPE foo [<!ENTITY xxe SYSTEM \"file:///etc/passwd\">]>";

    @Test
    public void doctypeInSheetXmlIsRejected() throws IOException {
        File file = maliciousWorkbook(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + XXE_DOCTYPE
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>&xxe;</t></is></c></row></sheetData>"
                + "</worksheet>",
            NORMAL_STYLES);

        assertThatThrownBy(() -> XlsxToMarkdownConverter.toMarkdown(file))
            .isInstanceOf(RuntimeException.class)
            .hasMessageContaining("metadata");
    }

    @Test
    public void doctypeInStylesXmlIsRejected() throws IOException {
        File file = maliciousWorkbook(NORMAL_SHEET,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + XXE_DOCTYPE
                + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
                + "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>"
                + "<borders count=\"1\"><border/></borders>"
                + "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>"
                + "<cellXfs count=\"1\"><xf xfId=\"0\"/></cellXfs>"
                + "</styleSheet>");

        assertThatThrownBy(() -> XlsxToMarkdownConverter.toMarkdown(file))
            .isInstanceOf(RuntimeException.class)
            .hasMessageContaining("metadata");
    }
}
