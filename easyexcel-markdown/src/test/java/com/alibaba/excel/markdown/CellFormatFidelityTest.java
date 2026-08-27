package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Cell format fidelity tests: values written by POI with real Excel number formats must come
 * out of the converter exactly as Excel would display them.
 *
 * @author wandl
 */
public class CellFormatFidelityTest {

    @TempDir
    File tempDir;

    private CellStyle style(Workbook workbook, String format) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat(format));
        return style;
    }

    private SheetDocument convert(Workbook workbook) throws IOException {
        File file = new File(tempDir, "fidelity.xlsx");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return XlsxToMarkdownConverter.toStructured(file);
    }

    @Test
    public void renderDatePercentThousandsAndBooleanAsDisplayed() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Formats");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Date");
        header.createCell(1).setCellValue("Percent");
        header.createCell(2).setCellValue("Amount");
        header.createCell(3).setCellValue("Flag");
        Row data = sheet.createRow(1);

        Cell date = data.createCell(0);
        date.setCellValue(new GregorianCalendar(2026, Calendar.AUGUST, 27).getTime());
        date.setCellStyle(style(workbook, "yyyy-mm-dd"));

        Cell percent = data.createCell(1);
        percent.setCellValue(0.5);
        percent.setCellStyle(style(workbook, "0%"));

        Cell amount = data.createCell(2);
        amount.setCellValue(1234.5);
        amount.setCellStyle(style(workbook, "#,##0.00"));

        data.createCell(3).setCellValue(true);

        SheetDocument document = convert(workbook);

        List<String> row = document.sheets.get(0).rows.get(0);
        assertThat(row.get(0)).isEqualTo("2026-08-27");
        assertThat(row.get(1)).isEqualTo("50%");
        assertThat(row.get(2)).isEqualTo("1,234.50");
        assertThat(row.get(3)).isEqualTo("true");
    }

    @Test
    public void keepEmptySheetsAndWorkbookOrder() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        workbook.createSheet("First");
        Sheet second = workbook.createSheet("Second");
        Row row = second.createRow(0);
        row.createCell(0).setCellValue("only");
        workbook.createSheet("Third");

        SheetDocument document = convert(workbook);

        assertThat(document.sheets).hasSize(3);
        assertThat(document.sheets.get(0).sheetName).isEqualTo("First");
        assertThat(document.sheets.get(0).headers).isEmpty();
        assertThat(document.sheets.get(0).rows).isEmpty();
        assertThat(document.sheets.get(1).sheetName).isEqualTo("Second");
        assertThat(document.sheets.get(1).headers.get(0)).containsExactly("only");
        assertThat(document.sheets.get(2).sheetName).isEqualTo("Third");

        assertThat(document.toMarkdown())
            .contains("## First")
            .contains("(empty sheet)")
            .contains("## Second")
            .contains("| only |")
            .contains("## Third");
    }

    @Test
    public void renderFormulaCachedResult() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Formula");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Sum");
        Row data = sheet.createRow(1);
        data.createCell(0).setCellFormula("1+1");
        // POI does not compute formulas on write, evaluate so the file carries the cached result
        // the same way a real Excel save does.
        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

        SheetDocument document = convert(workbook);

        assertThat(document.sheets.get(0).rows.get(0).get(0)).isEqualTo("2");
    }
}
