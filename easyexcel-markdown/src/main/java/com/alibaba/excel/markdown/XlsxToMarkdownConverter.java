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
import java.io.Writer;
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
import java.util.UUID;

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
import com.alibaba.excel.util.FileUtils;

/**
 * Convert XLSX / XLS / CSV workbooks to structured Markdown using the Alibaba EasyExcel streaming
 * read path, so 100MB+ files can be processed without loading the whole workbook into memory.
 * <p>
 * Every sheet is preserved: the first row of each sheet becomes the table header, merged regions
 * reported through {@link CellExtraTypeEnum#MERGE} are filled with the top-left value before
 * rendering, and CSV input is parsed with Apache Commons CSV.
 * <p>
 * Picture placeholders: XLSX and legacy XLS drawings are scanned so that cells carrying an
 * anchored picture render as {@code [image]}. The {@link InputStream} overloads spool the stream
 * to a temporary file for XLSX/XLS content to enable picture scanning; if you want to avoid disk
 * I/O, use the {@link File} overloads instead (CSV streams are never spooled).
 * <p>
 * Cell style mapping: bold, italic and strikeout fonts are rendered as inline Markdown markers
 * ({@code **bold**}, {@code *italic*}, {@code ~~strike~~}, combined as {@code ~~***mixed***~~}).
 * Header-row horizontal alignment maps to GFM column alignment in the delimiter row
 * ({@code :---:} for center, {@code ---:} for right, {@code ---} for left/general).
 * <p>
 * Known format boundary: color, border and background styling cannot be represented in GitHub
 * Flavored Markdown tables and are intentionally dropped.
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
     * <p>
     * For XLSX and XLS content the stream is spooled to a temporary file so that picture anchor
     * scanning is possible. CSV content is read directly from the stream without spooling.
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
            File cacheDir = FileUtils.createCacheTmpFile();
            try {
                File tmp = new File(cacheDir, UUID.randomUUID() + ".tmp");
                FileUtils.writeToFile(tmp, pushback, false);
                document.sheets.addAll(loadExcel(tmp, null));
            } finally {
                FileUtils.delete(cacheDir);
            }
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

    /**
     * Convert the workbook to Markdown and write the result into the given {@link Writer}.
     * The content is identical to {@link #toMarkdown(File)} but processes one sheet at a time
     * so that memory usage is bounded by a single sheet rather than the whole document.
     * <p>
     * The writer is flushed but <em>not</em> closed.
     *
     * @param file
     *            an XLSX, XLS or CSV file
     * @param writer
     *            the destination, flushed on completion
     * @throws NullPointerException
     *             if {@code file} or {@code writer} is {@code null}
     */
    public static void toMarkdown(File file, Writer writer) throws IOException {
        Objects.requireNonNull(file, "file must not be null");
        Objects.requireNonNull(writer, "writer must not be null");
        if (!file.isFile()) {
            throw new ExcelAnalysisException("File not found: " + file.getAbsolutePath());
        }
        writeTitle(writer, file.getName());
        if (file.getName().toLowerCase(Locale.ROOT).endsWith(CSV_SUFFIX)) {
            SheetTable table = loadCsv(file);
            writer.write(table.toMarkdown());
            writer.write('\n');
        } else {
            writeExcelStreaming(file, null, writer);
        }
        writer.flush();
    }

    /**
     * Convert the workbook read from a stream to Markdown and write the result into the given
     * {@link Writer}. The content is identical to {@link #toMarkdown(InputStream)} but processes
     * one sheet at a time so that memory usage is bounded by a single sheet rather than the whole
     * document.
     * <p>
     * For XLSX and XLS content the stream is spooled to a temporary file so that picture anchor
     * scanning is possible. CSV content is read directly from the stream without spooling.
     * <p>
     * The writer is flushed but <em>not</em> closed.
     *
     * @param inputStream
     *            the workbook content, UTF-8 is assumed for CSV
     * @param writer
     *            the destination, flushed on completion
     * @throws NullPointerException
     *             if {@code inputStream} or {@code writer} is {@code null}
     */
    public static void toMarkdown(InputStream inputStream, Writer writer) throws IOException {
        Objects.requireNonNull(inputStream, "inputStream must not be null");
        Objects.requireNonNull(writer, "writer must not be null");
        PushbackInputStream pushback = new PushbackInputStream(inputStream, ZIP_MAGIC.length);
        byte[] head = sniff(pushback);
        writeTitle(writer, "document");
        if (startsWith(head, ZIP_MAGIC) || startsWith(head, OLE2_MAGIC)) {
            File cacheDir = FileUtils.createCacheTmpFile();
            try {
                File tmp = new File(cacheDir, UUID.randomUUID() + ".tmp");
                FileUtils.writeToFile(tmp, pushback, false);
                writeExcelStreaming(tmp, null, writer);
            } finally {
                FileUtils.delete(cacheDir);
            }
        } else {
            SheetTable table = loadCsv(pushback, "CSV");
            writer.write(table.toMarkdown());
            writer.write('\n');
        }
        writer.flush();
    }

    private static void writeTitle(Writer writer, String title) throws IOException {
        writer.write("# ");
        writer.write(title == null ? "" : title);
        writer.write("\n\n");
    }

    /**
     * Streaming Excel read: process one sheet at a time, write its Markdown to the writer, then
     * discard it from memory before reading the next sheet.
     */
    private static void writeExcelStreaming(File file, InputStream inputStream, Writer writer) throws IOException {
        Map<Integer, Set<String>> pictureAnchors = scanAnchors(file);
        Map<Integer, Map<String, CellStyle>> cellStyles = scanCellStyles(file);
        final Map<Integer, SheetTable> sheetMap = new TreeMap<>();
        final Map<Integer, List<int[]>> mergeMap = new HashMap<>(16);
        final Map<Integer, Map<String, String>> hyperlinkMap = new HashMap<>(16);
        final Map<Integer, Map<String, String>> commentMap = new HashMap<>(16);
        ExcelReaderBuilder builder = file != null ? EasyExcel.read(file) : EasyExcel.read(inputStream);
        try (ExcelReader reader = builder.headRowNumber(0)
            .ignoreEmptyRow(false)
            .extraRead(CellExtraTypeEnum.MERGE)
            .extraRead(CellExtraTypeEnum.HYPERLINK)
            .extraRead(CellExtraTypeEnum.COMMENT)
            .registerReadListener(
                new MarkdownReadListener(sheetMap, mergeMap, hyperlinkMap, commentMap))
            .build()) {
            for (ReadSheet readSheet : reader.excelExecutor().sheetList()) {
                SheetTable table = new SheetTable();
                table.sheetName = readSheet.getSheetName();
                sheetMap.put(readSheet.getSheetNo(), table);
            }
            for (ReadSheet readSheet : reader.excelExecutor().sheetList()) {
                int sheetNo = readSheet.getSheetNo();
                reader.read(readSheet);
                SheetTable table = sheetMap.get(sheetNo);
                if (table != null) {
                    normalizeWidth(table);
                    applyMergedRegions(table, mergeMap.get(sheetNo));
                    applyCellExtras(table, hyperlinkMap.get(sheetNo), commentMap.get(sheetNo));
                    applyImagePlaceholders(table, pictureAnchors.get(sheetNo));
                    trimTrailingEmptyRows(table);
                    promoteFirstRowToHeader(table);
                    applyCellStyles(table, cellStyles.get(sheetNo));
                    applyColumnAlignments(table, cellStyles.get(sheetNo));
                    writer.write(table.toMarkdown());
                    writer.write('\n');
                }
                sheetMap.remove(sheetNo);
                mergeMap.remove(sheetNo);
                hyperlinkMap.remove(sheetNo);
                commentMap.remove(sheetNo);
            }
        }
    }

    private static List<SheetTable> loadExcel(File file, InputStream inputStream) {
        // TreeMap keeps the deterministic workbook order by sheet number.
        final Map<Integer, SheetTable> sheetMap = new TreeMap<>();
        final Map<Integer, List<int[]>> mergeMap = new HashMap<>(16);
        final Map<Integer, Map<String, String>> hyperlinkMap = new HashMap<>(16);
        final Map<Integer, Map<String, String>> commentMap = new HashMap<>(16);
        Map<Integer, Set<String>> pictureAnchors = scanAnchors(file);
        Map<Integer, Map<String, CellStyle>> cellStyles = scanCellStyles(file);
        ExcelReaderBuilder builder = file != null ? EasyExcel.read(file) : EasyExcel.read(inputStream);
        try (ExcelReader reader = builder.headRowNumber(0)
            .ignoreEmptyRow(false)
            .extraRead(CellExtraTypeEnum.MERGE)
            .extraRead(CellExtraTypeEnum.HYPERLINK)
            .extraRead(CellExtraTypeEnum.COMMENT)
            .registerReadListener(
                new MarkdownReadListener(sheetMap, mergeMap, hyperlinkMap, commentMap))
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
            normalizeWidth(table);
            applyMergedRegions(table, mergeMap.get(entry.getKey()));
            applyCellExtras(table, hyperlinkMap.get(entry.getKey()), commentMap.get(entry.getKey()));
            applyImagePlaceholders(table, pictureAnchors.get(entry.getKey()));
            trimTrailingEmptyRows(table);
            promoteFirstRowToHeader(table);
            applyCellStyles(table, cellStyles.get(entry.getKey()));
            applyColumnAlignments(table, cellStyles.get(entry.getKey()));
            tables.add(table);
        }
        return tables;
    }

    /**
     * Scan cell styles from the given file. XLSX (ZIP) files use the streaming SAX scanner over
     * styles.xml and sheet XML; legacy XLS (OLE2) files use the full HSSFWorkbook enumeration.
     * Returns an empty map when {@code file} is {@code null} (InputStream-only path without
     * spooling).
     */
    private static Map<Integer, Map<String, CellStyle>> scanCellStyles(File file) {
        if (file == null) {
            return Collections.emptyMap();
        }
        try {
            if (isZipFile(file)) {
                return CellStyleScanner.scanStyles(file);
            }
            return CellStyleScanner.scanLegacyStyles(file);
        } catch (IOException e) {
            throw new ExcelAnalysisException("Read cell styles failure", e);
        }
    }

    /**
     * Scan picture anchors from the given file. XLSX (ZIP) files use the streaming SAX scanner,
     * legacy XLS (OLE2) files use the full HSSFWorkbook enumeration. Returns an empty map when
     * {@code file} is {@code null} (InputStream-only path without spooling).
     */
    private static Map<Integer, Set<String>> scanAnchors(File file) {
        if (file == null) {
            return Collections.emptyMap();
        }
        try {
            if (isZipFile(file)) {
                return DrawingAnchorScanner.scanPictureAnchors(file);
            }
            return DrawingAnchorScanner.scanLegacyPictureAnchors(file);
        } catch (IOException e) {
            throw new ExcelAnalysisException("Read picture anchors failure", e);
        }
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
     * Pad every row to the widest row so that all later passes and the Markdown rendering see
     * a rectangular grid.
     */
    private static void normalizeWidth(SheetTable table) {
        if (table.rows.isEmpty()) {
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

    /**
     * Wrap hyperlinked cells as Markdown links and append cell comments after the value.
     */
    private static void applyCellExtras(SheetTable table, Map<String, String> hyperlinks,
        Map<String, String> comments) {
        boolean noHyperlinks = hyperlinks == null || hyperlinks.isEmpty();
        boolean noComments = comments == null || comments.isEmpty();
        if (noHyperlinks && noComments) {
            return;
        }
        for (int rowIndex = 0; rowIndex < table.rows.size(); rowIndex++) {
            List<String> row = table.rows.get(rowIndex);
            for (int columnIndex = 0; columnIndex < row.size(); columnIndex++) {
                if (!noHyperlinks) {
                    String url = hyperlinks.get(rowIndex + CELL_KEY_SEPARATOR + columnIndex);
                    if (url != null) {
                        row.set(columnIndex, markdownLink(row.get(columnIndex), url));
                    }
                }
                if (!noComments) {
                    String comment = comments.get(rowIndex + CELL_KEY_SEPARATOR + columnIndex);
                    if (comment != null) {
                        row.set(columnIndex, appendComment(row.get(columnIndex), comment));
                    }
                }
            }
        }
    }

    static String markdownLink(String text, String url) {
        String safeText = text == null || text.isEmpty() ? url : text;
        boolean needsAngleBrackets = url.contains(" ") || url.contains("(") || url.contains(")");
        return "[" + safeText + "](" + (needsAngleBrackets ? "<" + url + ">" : url) + ")";
    }

    static String appendComment(String value, String comment) {
        String folded = comment.replace("\r\n", " ").replace('\r', ' ').replace('\n', ' ').trim();
        String cell = value == null ? "" : value;
        return cell.isEmpty() ? "<!-- " + folded + " -->" : cell + " <!-- " + folded + " -->";
    }

    /**
     * Wrap non-empty cells with inline Markdown font markers according to their cell style.
     * The wrapping order is strike(bold(italic(value))), so bold+italic+strike becomes
     * {@code ~~***value***~~}. Styles are applied on top of any existing rendering (hyperlinks,
     * comments).
     */
    private static void applyCellStyles(SheetTable table, Map<String, CellStyle> styles) {
        if (styles == null || styles.isEmpty()) {
            return;
        }
        // After promoteFirstRowToHeader, table.rows indices are shifted by the number of
        // header rows removed.  Style map keys use the original row indices, so we must
        // add the header row count to translate from table.rows position to original index.
        int headerOffset = table.headers.size();
        for (int rowIndex = 0; rowIndex < table.rows.size(); rowIndex++) {
            List<String> row = table.rows.get(rowIndex);
            int originalRowIndex = rowIndex + headerOffset;
            for (int columnIndex = 0; columnIndex < row.size(); columnIndex++) {
                CellStyle style = styles.get(originalRowIndex + CELL_KEY_SEPARATOR + columnIndex);
                if (style == null) {
                    continue;
                }
                String value = row.get(columnIndex);
                if (value == null || value.isEmpty()) {
                    continue;
                }
                row.set(columnIndex, wrapWithFontMarkers(value, style));
            }
        }
    }

    /**
     * Wrap a value with inline Markdown font markers. Combination order:
     * strike wraps bold wraps italic wraps value.
     */
    static String wrapWithFontMarkers(String value, CellStyle style) {
        if (style.italic) {
            value = "*" + value + "*";
        }
        if (style.bold) {
            value = "**" + value + "**";
        }
        if (style.strikeout) {
            value = "~~" + value + "~~";
        }
        return value;
    }

    /**
     * Derive per-column alignment from the header row (row 0 before promotion) cell styles
     * and store the result in {@link SheetTable#columnAlignments}.
     */
    private static void applyColumnAlignments(SheetTable table, Map<String, CellStyle> styles) {
        if (styles == null || styles.isEmpty() || table.headers.isEmpty()) {
            return;
        }
        List<String> headerRow = table.headers.get(0);
        int columnCount = headerRow == null ? 0 : headerRow.size();
        if (columnCount == 0) {
            return;
        }
        List<String> alignments = new ArrayList<>(columnCount);
        for (int col = 0; col < columnCount; col++) {
            CellStyle style = styles.get("0" + CELL_KEY_SEPARATOR + col);
            String horizontal = (style != null) ? style.horizontal : null;
            if ("center".equals(horizontal)) {
                alignments.add("center");
            } else if ("right".equals(horizontal)) {
                alignments.add("right");
            } else {
                alignments.add(null);
            }
        }
        table.columnAlignments = alignments;
    }

    /**
     * Blank rows in the middle of the data are kept, but the styled-empty rows that trail real
     * content in most workbooks would only add noise, so they are cut.
     */
    private static void trimTrailingEmptyRows(SheetTable table) {
        while (!table.rows.isEmpty() && isBlankRow(table.rows.get(table.rows.size() - 1))) {
            table.rows.remove(table.rows.size() - 1);
        }
    }

    private static boolean isBlankRow(List<String> row) {
        for (String value : row) {
            if (value != null && !value.isEmpty()) {
                return false;
            }
        }
        return true;
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
     * Check whether a file starts with the ZIP magic bytes ({@code PK\x03\x04}).
     * Both XLSX (ZIP) and legacy XLS (OLE2) workbooks are dispatched here: ZIP means XLSX,
     * non-ZIP means legacy XLS.
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
     * Collect one {@link SheetTable} per sheet number plus every merged region, hyperlink and
     * comment reported by the {@code extra} callback.
     */
    private static final class MarkdownReadListener implements ReadListener<Map<Integer, String>> {
        private final Map<Integer, SheetTable> sheetMap;
        private final Map<Integer, List<int[]>> mergeMap;
        private final Map<Integer, Map<String, String>> hyperlinkMap;
        private final Map<Integer, Map<String, String>> commentMap;

        MarkdownReadListener(Map<Integer, SheetTable> sheetMap, Map<Integer, List<int[]>> mergeMap,
            Map<Integer, Map<String, String>> hyperlinkMap, Map<Integer, Map<String, String>> commentMap) {
            this.sheetMap = sheetMap;
            this.mergeMap = mergeMap;
            this.hyperlinkMap = hyperlinkMap;
            this.commentMap = commentMap;
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
                if (extra.getType() == CellExtraTypeEnum.HYPERLINK) {
                    putByCell(hyperlinkMap, context.readSheetHolder().getSheetNo(), extra.getRowIndex(),
                        extra.getColumnIndex(), extra.getText());
                } else if (extra.getType() == CellExtraTypeEnum.COMMENT) {
                    putByCell(commentMap, context.readSheetHolder().getSheetNo(), extra.getRowIndex(),
                        extra.getColumnIndex(), extra.getText());
                }
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

        private static void putByCell(Map<Integer, Map<String, String>> map, Integer sheetNo,
            Integer rowIndex, Integer columnIndex, String value) {
            Map<String, String> cells = map.get(sheetNo);
            if (cells == null) {
                cells = new HashMap<>(16);
                map.put(sheetNo, cells);
            }
            cells.put(rowIndex + CELL_KEY_SEPARATOR + columnIndex, value);
        }

        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {}
    }
}
