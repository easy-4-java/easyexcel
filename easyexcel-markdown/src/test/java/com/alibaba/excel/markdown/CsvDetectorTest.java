package com.alibaba.excel.markdown;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.StringReader;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test CSV charset detection and delimiter sniffing end to end.
 *
 * @author wandl
 */
public class CsvDetectorTest {

    @TempDir
    Path tempDir;

    private File csv(byte[] bytes) throws IOException {
        File file = tempDir.resolve("sniff.csv").toFile();
        Files.write(file.toPath(), bytes);
        return file;
    }

    @Test
    public void decodeGbkCsvWithChineseHeader() throws IOException {
        byte[] gbk = "姓名,分数\n张三,90\n".getBytes(Charset.forName("GBK"));
        SheetDocument document = XlsxToMarkdownConverter.toStructured(csv(gbk));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("姓名", "分数"));
        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("张三", "90"));
    }

    @Test
    public void decodeUtf16LeWithBomAndSemicolon() throws IOException {
        byte[] utf16 = "\uFEFF姓名;分数\n张三;90\n".getBytes(StandardCharsets.UTF_16LE);
        SheetDocument document = XlsxToMarkdownConverter.toStructured(csv(utf16));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("姓名", "分数"));
        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("张三", "90"));
    }

    @Test
    public void sniffSemicolonDelimiter() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(
            csv("a;b;c\n1;2;3\n".getBytes(StandardCharsets.UTF_8)));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("a", "b", "c"));
        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("1", "2", "3"));
    }

    @Test
    public void sniffTabDelimiter() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(
            csv("a\tb\n1\t2\n".getBytes(StandardCharsets.UTF_8)));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("a", "b"));
    }

    @Test
    public void quotedDelimiterDoesNotConfuseSniffing() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(
            csv("name;note\na;\"x;y\"\n".getBytes(StandardCharsets.UTF_8)));

        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("a", "x;y"));
    }

    @Test
    public void utf8StillWorksAndBomStripped() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(
            csv("\uFEFFh1,h2\nv1,v2\n".getBytes(StandardCharsets.UTF_8)));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("h1", "h2"));
    }

    @Test
    public void detectDelimiterResetsReader() throws IOException {
        BufferedReader reader = new BufferedReader(new StringReader("a;b\n1;2\n"));
        char delimiter = CsvDetector.detectDelimiter(reader);

        assertThat(delimiter).isEqualTo(';');
        assertThat(reader.readLine()).isEqualTo("a;b");
    }
}
