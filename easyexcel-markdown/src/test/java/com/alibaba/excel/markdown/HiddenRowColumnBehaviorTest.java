package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Document the current behavior of {@link XlsxToMarkdownConverter} with respect to hidden rows
 * and hidden columns.
 * <p>
 * <b>Fact-checked behavior (2026-08-27):</b> easyexcel-core's SAX read path does <em>not</em>
 * inspect the {@code hidden} attribute on {@code <row>} or {@code <col>} XML elements. Evidence:
 * <ul>
 *   <li>{@code RowTagHandler.startElement()} (easyexcel-core, line 26-39) reads only the {@code r}
 *       attribute to determine the row index; it never calls
 *       {@code attributes.getValue("hidden")}.</li>
 *   <li>{@code XlsxRowHandler} registers no handler for the {@code <col>} element at all, so
 *       column-level attributes (including {@code hidden}) are never consumed.</li>
 *   <li>{@code ExcelXmlConstants} contains no {@code "hidden"} constant.</li>
 * </ul>
 * As a result, hidden rows and columns are included in the output as if they were visible.
 * This test locks in that behavior so a future change would be a conscious, documented decision.
 *
 * @author wandl
 */
public class HiddenRowColumnBehaviorTest {

    @TempDir
    File tempDir;

    /**
     * Create an XLSX with:
     * <ul>
     *   <li>Row 0 (header): visible, cells "A", "B"</li>
     *   <li>Row 1: <b>hidden</b>, cells "secret", "data"</li>
     *   <li>Row 2: visible, cells "open", "info"</li>
     * </ul>
     * Assert that the hidden row appears in the Markdown output (current behavior: no filtering).
     */
    @Test
    public void hiddenRowIsIncludedInOutput() throws IOException {
        File file = new File(tempDir, "hidden-row.xlsx");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("A");
        header.createCell(1).setCellValue("B");

        Row hiddenRow = sheet.createRow(1);
        hiddenRow.createCell(0).setCellValue("secret");
        hiddenRow.createCell(1).setCellValue("data");
        hiddenRow.setZeroHeight(true);  // marks the row as hidden in OOXML

        Row visibleRow = sheet.createRow(2);
        visibleRow.createCell(0).setCellValue("open");
        visibleRow.createCell(1).setCellValue("info");

        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        SheetTable table = document.sheets.get(0);

        // The hidden row is present in the data rows (current behavior: no filtering)
        assertThat(table.rows).hasSize(2);
        assertThat(table.rows.get(0)).containsExactly("secret", "data");
        assertThat(table.rows.get(1)).containsExactly("open", "info");

        // Verify the Markdown output includes the hidden row content
        String md = document.toMarkdown();
        assertThat(md).contains("| secret | data |");
        assertThat(md).contains("| open | info |");
    }

    /**
     * Create an XLSX where column B is declared hidden via {@code sheet.setColumnHidden(1, true)}.
     * <p>
     * Note: easyexcel-core does not parse the {@code <col>} element at all, so hidden-column
     * information is not available during SAX reading. POI's {@code setColumnHidden} only affects
     * the column metadata (CTCol), not the cell data itself. The cell values in column B are
     * still written to the sheet XML and are read normally by the SAX parser.
     * <p>
     * This test asserts the current behavior: hidden columns and their data are included.
     */
    @Test
    public void hiddenColumnDataIsIncludedInOutput() throws IOException {
        File file = new File(tempDir, "hidden-col.xlsx");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Cols");

        sheet.setColumnHidden(1, true);  // hide column B (index 1)

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Secret");
        header.createCell(2).setCellValue("Role");

        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("Alice");
        data.createCell(1).setCellValue("s3cret");
        data.createCell(2).setCellValue("Admin");

        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
        SheetTable table = document.sheets.get(0);

        // The hidden column's data is present (current behavior: no column filtering)
        assertThat(table.headers.get(0)).containsExactly("Name", "Secret", "Role");
        assertThat(table.rows.get(0)).containsExactly("Alice", "s3cret", "Admin");

        String md = document.toMarkdown();
        assertThat(md).contains("| Name | Secret | Role |");
        assertThat(md).contains("| Alice | s3cret | Admin |");
    }
}
