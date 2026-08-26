package com.alibaba.excel.markdown;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FilterReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PushbackInputStream;
import java.io.Reader;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.TreeMap;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.read.metadata.holder.ReadSheetHolder;

/**
 * Convert XLSX / XLS / CSV workbooks to structured Markdown using the Alibaba EasyExcel streaming
 * read path, so 100MB+ files can be processed without loading the whole workbook into memory.
 * <p>
 * Every sheet is preserved: the first row of each sheet becomes the table header, merged regions
 * reported through {@link CellExtraTypeEnum#MERGE} are filled with the top-left value before
 * rendering, and CSV input is parsed with Apache Commons CSV.
 *
 * @author wandl
 */
public final class XlsxToMarkdownConverter {

    private static final byte[] ZIP_MAGIC = {0x50, 0x4B, 0x03, 0x04};
    private static final byte[] OLE2_MAGIC = {(byte)0xD0, (byte)0xCF, 0x11, (byte)0xE0};
    private static final String CSV_SUFFIX = ".csv";
    private static final char UTF8_BOM = '\uFEFF';
    private static final String IMAGE_PLACEHOLDER = "[image]";
    private static final String CELL_KEY_SEPARATOR = ":";

    private XlsxToMarkdownConverter() {}

    /**
     * Read the whole workbook and keep every sheet as a {@link SheetTable}.
     *
     * @param file
     *            an XLSX, XLS or CSV file, dispatch is done by extension
     * @return the structured document, never {@code null}
     */
    public static SheetDocument toStructured(File file) {
        Objects.requireNonNull(file, "file must not be null");
        if (!file.isFile()) {
            throw new ExcelAnalysisException("File not found: " + file.getAbsolutePath());
        }
        SheetDocument document = new SheetDocument();
        document.title = file.getName();
        if (file.getName().toLowerCase(Locale.ROOT).endsWith(CSV_SUFFIX)) {
            document.sheets.add(loadCsv(file));
        } else {
            document.sheets.addAll(loadExcel(file, null));
        }
        return document;
    }

    /**
     * Read the whole workbook from a stream, dispatch is done by content sniffing of the first
     * bytes (ZIP magic means XLSX, OLE2 magic means XLS, anything else is treated as CSV).
     *
     * @param inputStream
     *            the workbook content, UTF-8 is assumed for CSV
     * @return the structured document, never {@code null}
     */
    public static SheetDocument toStructured(InputStream inputStream) {
        Objects.requireNonNull(inputStream, "inputStream must not be null");
        PushbackInputStream pushback = new PushbackInputStream(inputStream, ZIP_MAGIC.length);
        byte[] head = sniff(pushback);
        SheetDocument document = new SheetDocument();
        document.title = "document";
        if (startsWith(head, ZIP_MAGIC) || startsWith(head, OLE2_MAGIC)) {
            document.sheets.addAll(loadExcel(null, pushback));
        } else {
            document.sheets.add(loadCsv(pushback, "CSV"));
        }
        return document;
    }

    /**
     * Convert the workbook to a multi-sheet Markdown document.
     */
    public static String toMarkdown(File file) {
        return toStructured(file).toMarkdown();
    }

    /**
     * Convert the workbook read from a stream to a multi-sheet Markdown document.
     */
    public static String toMarkdown(InputStream inputStream) {
        return toStructured(inputStream).toMarkdown();
    }

    private static List<SheetTable> loadExcel(File file, InputStream inputStream) {
        // TreeMap keeps the deterministic workbook order by sheet number.
        final Map<Integer, SheetTable> sheetMap = new TreeMap<>();
        final Map<Integer, List<int[]>> mergeMap = new HashMap<>(16);
        Map<Integer, Set<String>> pictureAnchors = Collections.emptyMap();
        if (file != null && isZipFile(file)) {
            try {
                pictureAnchors = DrawingAnchorScanner.scanPictureAnchors(file);
            } catch (IOException e) {
                throw new ExcelAnalysisException("Read picture anchors failure", e);
            }
        }
        ExcelReaderBuilder builder = file != null ? EasyExcel.read(file) : EasyExcel.read(inputStream);
        try (ExcelReader reader = builder.headRowNumber(0)
            .extraRead(CellExtraTypeEnum.MERGE)
            .registerReadListener(new MarkdownReadListener(sheetMap, mergeMap))
            .build()) {
            // Pre register every sheet so that sheets without any row are still present.
            for (ReadSheet readSheet : reader.excelExecutor().sheetList()) {
                SheetTable table = new SheetTable();
                table.sheetName = readSheet.getSheetName();
                sheetMap.put(readSheet.getSheetNo(), table);
            }
            reader.readAll();
        }

        List<SheetTable> tables = new ArrayList<>();
        for (Map.Entry<Integer, SheetTable> entry : sheetMap.entrySet()) {
            SheetTable table = entry.getValue();
            applyMergedRegions(table, mergeMap.get(entry.getKey()));
            applyImagePlaceholders(table, pictureAnchors.get(entry.getKey()));
            promoteFirstRowToHeader(table);
            tables.add(table);
        }
        return tables;
    }

    /**
     * Mark the anchor cell of every picture with a placeholder, keeping any value the cell
     * already carries. Rows and columns that only exist because of a picture are created.
     */
    private static void applyImagePlaceholders(SheetTable table, Set<String> anchors) {
        if (anchors == null || anchors.isEmpty()) {
            return;
        }
        int maxRow = -1;
        for (String anchor : anchors) {
            maxRow = Math.max(maxRow, Integer.parseInt(anchor.substring(0, anchor.indexOf(CELL_KEY_SEPARATOR))));
        }
        while (table.rows.size() <= maxRow) {
            table.rows.add(new ArrayList<>());
        }
        for (String anchor : anchors) {
            int separator = anchor.indexOf(CELL_KEY_SEPARATOR);
            int rowIndex = Integer.parseInt(anchor.substring(0, separator));
            int columnIndex = Integer.parseInt(anchor.substring(separator + CELL_KEY_SEPARATOR.length()));
            List<String> row = table.rows.get(rowIndex);
            for (int columnIndexToPad = row.size(); columnIndexToPad <= columnIndex; columnIndexToPad++) {
                row.add("");
            }
            if (row.get(columnIndex).isEmpty()) {
                row.set(columnIndex, IMAGE_PLACEHOLDER);
            }
        }
    }

    private static SheetTable loadCsv(File file) {
        SheetTable table = new SheetTable();
        table.sheetName = baseName(file.getName());
        try (InputStream in = Files.newInputStream(file.toPath())) {
            fillFromCsv(table, in);
        } catch (IOException e) {
            throw new ExcelAnalysisException("Read CSV failure: " + file.getAbsolutePath(), e);
        }
        return table;
    }

    private static SheetTable loadCsv(InputStream inputStream, String sheetName) {
        SheetTable table = new SheetTable();
        table.sheetName = sheetName;
        try {
            fillFromCsv(table, inputStream);
        } catch (IOException e) {
            throw new ExcelAnalysisException("Read CSV failure", e);
        }
        return table;
    }

    private static void fillFromCsv(SheetTable table, InputStream in) throws IOException {
        BufferedInputStream buffered = in instanceof BufferedInputStream ? (BufferedInputStream)in
            : new BufferedInputStream(in, CsvDetector.SNIFF_LIMIT);
        buffered.mark(CsvDetector.SNIFF_LIMIT);
        byte[] head = new byte[CsvDetector.SNIFF_LIMIT];
        int headLength = 0;
        int count;
        while (headLength < head.length
            && (count = buffered.read(head, headLength, head.length - headLength)) != -1) {
            headLength += count;
        }
        buffered.reset();
        Reader decoded = new BomStrippingReader(
            new InputStreamReader(buffered, CsvDetector.detectCharset(head)));
        BufferedReader lineReader = new BufferedReader(decoded, CsvDetector.SNIFF_LIMIT);
        char delimiter = CsvDetector.detectDelimiter(lineReader);
        try (CSVParser parser = new CSVParser(lineReader,
            CSVFormat.DEFAULT.builder().setDelimiter(delimiter).build())) {
            for (CSVRecord record : parser) {
                List<String> row = new ArrayList<>();
                for (String value : record) {
                    row.add(value);
                }
                table.rows.add(row);
            }
        }
        promoteFirstRowToHeader(table);
    }

    /**
     * Fill every cell of a merged region with the value of its top-left cell, so the flat
     * Markdown table still shows the merged content on each covered position.
     */
    private static void applyMergedRegions(SheetTable table, List<int[]> regions) {
        if (regions == null || regions.isEmpty() || table.rows.isEmpty()) {
            return;
        }
        int width = 0;
        for (List<String> row : table.rows) {
            width = Math.max(width, row.size());
        }
        for (List<String> row : table.rows) {
            for (int columnIndex = row.size(); columnIndex < width; columnIndex++) {
                row.add("");
            }
        }
        for (int[] region : regions) {
            int firstRow = region[0];
            int lastRow = region[1];
            int firstColumn = region[2];
            int lastColumn = region[3];
            if (firstRow >= table.rows.size() || firstColumn >= width) {
                continue;
            }
            String value = cell(table.rows, firstRow, firstColumn);
            if (value == null || value.isEmpty()) {
                continue;
            }
            for (int rowIndex = firstRow; rowIndex <= lastRow && rowIndex < table.rows.size(); rowIndex++) {
                for (int columnIndex = firstColumn; columnIndex <= lastColumn && columnIndex < width;
                    columnIndex++) {
                    table.rows.get(rowIndex).set(columnIndex, value);
                }
            }
        }
    }

    private static void promoteFirstRowToHeader(SheetTable table) {
        if (!table.rows.isEmpty()) {
            table.headers.add(table.rows.remove(0));
        }
    }

    private static String cell(List<List<String>> rows, int rowIndex, int columnIndex) {
        List<String> row = rows.get(rowIndex);
        return columnIndex < row.size() ? row.get(columnIndex) : null;
    }

    private static String baseName(String fileName) {
        int dot = fileName.lastIndexOf('.');
        return dot > 0 ? fileName.substring(0, dot) : fileName;
    }

    private static byte[] sniff(PushbackInputStream pushback) {
        byte[] head = new byte[ZIP_MAGIC.length];
        int read = 0;
        try {
            while (read < head.length) {
                int count = pushback.read(head, read, head.length - read);
                if (count < 0) {
                    break;
                }
                read += count;
            }
            pushback.unread(head, 0, read);
        } catch (IOException e) {
            throw new ExcelAnalysisException("Sniff stream failure", e);
        }
        return head;
    }

    private static boolean startsWith(byte[] head, byte[] magic) {
        if (head.length < magic.length) {
            return false;
        }
        for (int i = 0; i < magic.length; i++) {
            if (head[i] != magic[i]) {
                return false;
            }
        }
        return true;
    }

    /**
     * Only ZIP based workbooks (XLSX) can carry drawings the scanner understands, legacy XLS is
     * skipped on purpose.
     */
    private static boolean isZipFile(File file) {
        byte[] head = new byte[ZIP_MAGIC.length];
        int read = 0;
        try (InputStream in = Files.newInputStream(file.toPath())) {
            while (read < head.length && (read += in.read(head, read, head.length - read)) != -1) {
            }
        } catch (IOException e) {
            return false;
        }
        return startsWith(head, ZIP_MAGIC);
    }

    /**
     * Reader wrapper that drops a single leading UTF-8 byte order mark.
     */
    private static final class BomStrippingReader extends FilterReader {
        private boolean first = true;

        BomStrippingReader(Reader delegate) {
            super(delegate);
        }

        @Override
        public int read(char[] cbuf, int off, int len) throws IOException {
            int count = super.read(cbuf, off, len);
            if (first && count > 0 && cbuf[off] == UTF8_BOM) {
                System.arraycopy(cbuf, off + 1, cbuf, off, count - 1);
                count--;
                if (count == 0) {
                    count = super.read(cbuf, off, len);
                }
            }
            first = false;
            return count;
        }
    }

    /**
     * Collect one {@link SheetTable} per sheet number and every merged region reported by the
     * {@code extra} callback.
     */
    private static final class MarkdownReadListener implements ReadListener<Map<Integer, String>> {
        private final Map<Integer, SheetTable> sheetMap;
        private final Map<Integer, List<int[]>> mergeMap;

        MarkdownReadListener(Map<Integer, SheetTable> sheetMap, Map<Integer, List<int[]>> mergeMap) {
            this.sheetMap = sheetMap;
            this.mergeMap = mergeMap;
        }

        @Override
        public void invoke(Map<Integer, String> data, AnalysisContext context) {
            ReadSheetHolder sheetHolder = context.readSheetHolder();
            SheetTable table = sheetMap.get(sheetHolder.getSheetNo());
            if (table == null) {
                table = new SheetTable();
                table.sheetName = sheetHolder.getSheetName();
                sheetMap.put(sheetHolder.getSheetNo(), table);
            }
            List<String> row = new ArrayList<>();
            int width = 0;
            for (Integer columnIndex : data.keySet()) {
                width = Math.max(width, columnIndex + 1);
            }
            for (int columnIndex = 0; columnIndex < width; columnIndex++) {
                String value = data.get(columnIndex);
                row.add(value == null ? "" : value);
            }
            table.rows.add(row);
        }

        @Override
        public void extra(CellExtra extra, AnalysisContext context) {
            if (extra.getType() != CellExtraTypeEnum.MERGE) {
                return;
            }
            List<int[]> regions = mergeMap.get(context.readSheetHolder().getSheetNo());
            if (regions == null) {
                regions = new ArrayList<>();
                mergeMap.put(context.readSheetHolder().getSheetNo(), regions);
            }
            regions.add(new int[] {extra.getFirstRowIndex(), extra.getLastRowIndex(),
                extra.getFirstColumnIndex(), extra.getLastColumnIndex()});
        }

        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {}
    }
}
