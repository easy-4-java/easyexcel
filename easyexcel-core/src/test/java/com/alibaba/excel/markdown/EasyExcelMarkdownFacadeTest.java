package com.alibaba.excel.markdown;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test the {@link EasyExcel#xlsxToMarkdown} / {@link EasyExcel#xlsxToStructured} facade methods.
 *
 * @author wandl
 */
public class EasyExcelMarkdownFacadeTest {

    @TempDir
    Path tempDir;

    private File writeSimpleXlsx() throws IOException {
        File file = tempDir.resolve("facade.xlsx").toFile();
        ExcelWriter writer = EasyExcel.write(file)
            .head(Collections.singletonList(Collections.singletonList("Name")))
            .build();
        try {
            writer.write(Collections.<List<Object>>singletonList(
                Collections.singletonList("Zhang San")), EasyExcel.writerSheet(0, "People").build());
        } finally {
            writer.finish();
        }
        return file;
    }

    @Test
    public void facadeReadsStructuredAndMarkdownFromFile() throws IOException {
        SheetDocument document = EasyExcel.xlsxToStructured(writeSimpleXlsx());

        assertThat(document.title).isEqualTo("facade.xlsx");
        assertThat(document.sheets.get(0).sheetName).isEqualTo("People");

        assertThat(EasyExcel.xlsxToMarkdown(writeSimpleXlsx()))
            .startsWith("# facade.xlsx")
            .contains("## People")
            .contains("| Name |")
            .contains("| Zhang San |");
    }

    @Test
    public void facadeReadsFromInputStream() throws IOException {
        byte[] csv = "h1\nv1\n".getBytes(StandardCharsets.UTF_8);

        SheetDocument document = EasyExcel.xlsxToStructured(new ByteArrayInputStream(csv));
        assertThat(document.sheets.get(0).rows).containsExactly(Collections.singletonList("v1"));

        assertThat(EasyExcel.xlsxToMarkdown(new ByteArrayInputStream(
            Files.readAllBytes(writeSimpleXlsx().toPath())))).contains("## People");
    }
}
