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
import java.nio.charset.Charset;
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

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;

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
 * Hidden rows and columns: during the same SAX pass that collects cell styles,
 * {@link CellStyleScanner} also reads the {@code hidden} attribute on {@code <row>} and
 * {@code <col>} XML elements. Hidden rows and columns are filtered out of the Markdown output.
 * For legacy XLS files, {@link HSSFRow#getZeroHeight()} and
 * {@link HSSFSheet#isColumnHidden(int)} are used instead.
 * <p>
 * Picture placeholders: XLSX and legacy XLS drawings are scanned so that cells carrying an
 * anchored picture render as {@code [image]}. The {@link InputStream} overloads spool the stream
     * to a temporary file for XLSX/XLS content to enable picture scanning and for CSV content to
     * enable two-pass column-width detection; if you want to avoid disk I/O, use the {@link File}
     * overloads instead.
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
    private static final int MIN_HEAD_ROWS = 1;

    /** HTML span 开始标签常量。 */
    private static final String SPAN_STYLE_OPEN = "<span style=\"";
    /** HTML span 关闭标签常量。 */
    private static final String SPAN_CLOSE = "</span>";
    /** CSS color 属性前缀常量。 */
    private static final String CSS_COLOR = "color:";
    /** CSS background-color 属性前缀常量。 */
    private static final String CSS_BG_COLOR = "background-color:";

    /**
     * HTML 颜色渲染开关。默认 {@code false}，开启后用 {@code <span style="...">} 输出
     * 字体色和背景色。关闭时输出与旧版本完全一致。
     */
    private static boolean htmlColorRendering = false;

    /**
     * 设置 HTML 颜色渲染开关。
     *
     * @param enabled {@code true} 开启颜色渲染，{@code false} 关闭（默认）
     */
    public static void setHtmlColorRendering(boolean enabled) {
        htmlColorRendering = enabled;
    }

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
     * Read the whole workbook and keep every sheet as a {@link SheetTable}, using the specified
     * number of header rows per sheet. When {@code headRowNumber} is 1 the result is identical to
     * {@link #toStructured(File)}. For {@code headRowNumber >= 2}, the first N rows of each sheet
     * are promoted to multi-row headers (flattened with spaces when rendering to Markdown), and
     * the remaining rows become data rows.
     *
     * @param file
     *            an XLSX, XLS or CSV file
     * @param headRowNumber
     *            number of rows to use as header rows (&ge; 1)
     * @return the structured document, never {@code null}
     * @throws IllegalArgumentException
     *             if {@code headRowNumber} is less than 1, or if {@code headRowNumber > 1} and
     *             the file is a CSV file (multi-row headers are only supported for Excel files)
     */
    public static SheetDocument toStructured(File file, int headRowNumber) {
        Objects.requireNonNull(file, "file must not be null");
        if (headRowNumber < MIN_HEAD_ROWS) {
            throw new IllegalArgumentException(
                "headRowNumber must be >= " + MIN_HEAD_ROWS + ", got " + headRowNumber);
        }
        if (!file.isFile()) {
            throw new ExcelAnalysisException("File not found: " + file.getAbsolutePath());
        }
        boolean isCsv = file.getName().toLowerCase(Locale.ROOT).endsWith(CSV_SUFFIX);
        if (isCsv && headRowNumber > MIN_HEAD_ROWS) {
            throw new IllegalArgumentException(
                "headRowNumber > 1 is only supported for Excel files, got CSV: " + file.getName());
        }
        SheetDocument document = new SheetDocument();
        document.title = file.getName();
        if (isCsv) {
            document.sheets.add(loadCsv(file));
        } else {
            document.sheets.addAll(loadExcel(file, null, headRowNumber));
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
     * Convert the workbook to a multi-sheet Markdown document, using the specified number of
     * header rows per sheet.
     *
     * @param file
     *            an XLSX, XLS or CSV file
     * @param headRowNumber
     *            number of rows to use as header rows (&ge; 1)
     * @return the Markdown output, never {@code null}
     */
    public static String toMarkdown(File file, int headRowNumber) {
        return toStructured(file, headRowNumber).toMarkdown();
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
            writeCsvStreaming(file, null, baseName(file.getName()), writer);
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
     * scanning is possible. CSV content is also spooled to a temporary file so that the two-pass
     * streaming approach (pass 1: determine column count; pass 2: write rows) can re-read the
     * input without buffering the entire content in memory.
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
            writeCsvStreaming(null, pushback, "CSV", writer);
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
        Map<Integer, Set<String>> pictureAnchors = new HashMap<>(16);
        Map<Integer, Map<String, CellStyle>> cellStyles = new HashMap<>(16);
        Map<Integer, CellStyleScanner.HiddenInfo> hiddenInfoMap = new HashMap<>(16);
        scanFileMetadata(file, pictureAnchors, cellStyles, hiddenInfoMap);
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
                    CellStyleScanner.HiddenInfo hiddenInfo = hiddenInfoMap.get(sheetNo);
                    applyImagePlaceholders(table, pictureAnchors.get(sheetNo), hiddenInfo);
                    trimTrailingEmptyRows(table);
                    promoteFirstRowToHeader(table);
                    HiddenFilterResult filterResult = filterHiddenRowsAndColumns(table, hiddenInfo);
                    applyCellStyles(table, cellStyles.get(sheetNo),
                        filterResult.newToOrigRow, filterResult.newToOrigCol);
                    applyColumnAlignments(table, cellStyles.get(sheetNo), filterResult.newToOrigCol);
                    writer.write(table.toMarkdown());
                    writer.write('\n');
                }
                sheetMap.remove(sheetNo);
                mergeMap.remove(sheetNo);
                hyperlinkMap.remove(sheetNo);
                commentMap.remove(sheetNo);
                pictureAnchors.remove(sheetNo);
                cellStyles.remove(sheetNo);
                hiddenInfoMap.remove(sheetNo);
            }
        }
    }

    private static List<SheetTable> loadExcel(File file, InputStream inputStream) {
        return loadExcel(file, inputStream, MIN_HEAD_ROWS);
    }

    private static List<SheetTable> loadExcel(File file, InputStream inputStream, int headRowNumber) {
        // TreeMap keeps the deterministic workbook order by sheet number.
        final Map<Integer, SheetTable> sheetMap = new TreeMap<>();
        final Map<Integer, List<int[]>> mergeMap = new HashMap<>(16);
        final Map<Integer, Map<String, String>> hyperlinkMap = new HashMap<>(16);
        final Map<Integer, Map<String, String>> commentMap = new HashMap<>(16);
        Map<Integer, Set<String>> pictureAnchors = new HashMap<>(16);
        Map<Integer, Map<String, CellStyle>> cellStyles = new HashMap<>(16);
        Map<Integer, CellStyleScanner.HiddenInfo> hiddenInfoMap = new HashMap<>(16);
        scanFileMetadata(file, pictureAnchors, cellStyles, hiddenInfoMap);
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
            int sheetNo = entry.getKey();
            SheetTable table = entry.getValue();
            normalizeWidth(table);
            applyMergedRegions(table, mergeMap.get(sheetNo), headRowNumber);
            applyCellExtras(table, hyperlinkMap.get(sheetNo), commentMap.get(sheetNo));
            CellStyleScanner.HiddenInfo hiddenInfo = hiddenInfoMap.get(sheetNo);
            applyImagePlaceholders(table, pictureAnchors.get(sheetNo), hiddenInfo);
            trimTrailingEmptyRows(table);
            promoteRowsToHeader(table, headRowNumber);
            HiddenFilterResult filterResult = filterHiddenRowsAndColumns(table, hiddenInfo);
            applyCellStyles(table, cellStyles.get(sheetNo),
                filterResult.newToOrigRow, filterResult.newToOrigCol);
            applyColumnAlignments(table, cellStyles.get(sheetNo), filterResult.newToOrigCol);
            tables.add(table);
        }
        return tables;
    }

    /**
     * Scan picture anchors, cell styles and hidden row/col info from the given file in a single pass.
     * For XLSX (ZIP) files, opens {@link OPCPackage} once and shares the {@link XSSFReader}
     * between the two scanners, avoiding redundant package openings. The package is closed
     * before this method returns so that downstream readers (e.g. EasyExcel) can open their
     * own copy without conflict.
     * <p>
     * For legacy XLS (OLE2) files, the two scanners are called independently (each loads the
     * workbook into memory, acceptable given XLS row/column limits).
     * <p>
     * When {@code file} is {@code null}, both output maps are left empty.
     *
     * @param file
     *            the workbook file, or {@code null}
     * @param anchorsOut
     *            output map for picture anchors (populated in-place)
     * @param stylesOut
     *            output map for cell styles (populated in-place)
     * @param hiddenOut
     *            output map for hidden row/col info (populated in-place)
     */
    private static void scanFileMetadata(File file,
        Map<Integer, Set<String>> anchorsOut,
        Map<Integer, Map<String, CellStyle>> stylesOut,
        Map<Integer, CellStyleScanner.HiddenInfo> hiddenOut) {
        if (file == null) {
            return;
        }
        try {
            if (isZipFile(file)) {
                // XLSX: 打开一次 OPCPackage，把同一个 XSSFReader 传给两个扫描器
                OPCPackage pkg = null;
                try {
                    pkg = OPCPackage.open(file, PackageAccess.READ);
                    XSSFReader reader = new XSSFReader(pkg);
                    anchorsOut.putAll(DrawingAnchorScanner.scanPictureAnchors(reader));
                    CellStyleScanner.ScanResult scanResult = CellStyleScanner.scanStyles(reader);
                    stylesOut.putAll(scanResult.styles);
                    hiddenOut.putAll(scanResult.hiddenInfo);
                } catch (IOException e) {
                    throw e;
                } finally {
                    if (pkg != null) {
                        try {
                            pkg.close();
                        } catch (IOException e) {
                            // closing a read only package best effort
                        }
                    }
                }
            } else {
                // Legacy XLS
                anchorsOut.putAll(DrawingAnchorScanner.scanLegacyPictureAnchors(file));
                CellStyleScanner.ScanResult scanResult = CellStyleScanner.scanLegacyStyles(file);
                stylesOut.putAll(scanResult.styles);
                hiddenOut.putAll(scanResult.hiddenInfo);
            }
        } catch (IOException e) {
            throw new ExcelAnalysisException("Read metadata failure", e);
        } catch (Exception e) {
            throw new ExcelAnalysisException("Read metadata failure", e);
        }
    }

    /**
     * Mark the anchor cell of every picture with a placeholder, keeping any value the cell
     * already carries. Rows and columns that only exist because of a picture are created.
     * 跳过位于隐藏行或隐藏列上的图片锚点。
     *
     * @param table
     *            the sheet table
     * @param anchors
     *            picture anchor positions in {@code "row:col"} format
     * @param hiddenInfo
     *            hidden row/col info, may be {@code null}
     */
    private static void applyImagePlaceholders(SheetTable table, Set<String> anchors,
        CellStyleScanner.HiddenInfo hiddenInfo) {
        if (anchors == null || anchors.isEmpty()) {
            return;
        }
        Set<Integer> hiddenRows = (hiddenInfo != null)
            ? hiddenInfo.hiddenRows : Collections.<Integer>emptySet();
        Set<Integer> hiddenCols = (hiddenInfo != null)
            ? hiddenInfo.hiddenCols : Collections.<Integer>emptySet();
        int maxRow = -1;
        for (String anchor : anchors) {
            int separator = anchor.indexOf(CELL_KEY_SEPARATOR);
            int rowIndex = Integer.parseInt(anchor.substring(0, separator));
            int columnIndex = Integer.parseInt(anchor.substring(separator + CELL_KEY_SEPARATOR.length()));
            if (hiddenRows.contains(rowIndex) || hiddenCols.contains(columnIndex)) {
                continue;
            }
            maxRow = Math.max(maxRow, rowIndex);
        }
        while (table.rows.size() <= maxRow) {
            table.rows.add(new ArrayList<>());
        }
        for (String anchor : anchors) {
            int separator = anchor.indexOf(CELL_KEY_SEPARATOR);
            int rowIndex = Integer.parseInt(anchor.substring(0, separator));
            int columnIndex = Integer.parseInt(anchor.substring(separator + CELL_KEY_SEPARATOR.length()));
            if (hiddenRows.contains(rowIndex) || hiddenCols.contains(columnIndex)) {
                continue;
            }
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
     * Two-pass streaming CSV to Markdown writer.
     * <p>
     * Pass 1: stream every record to determine {@code maxColumns} (no row storage).
     * Pass 2: re-open the CSV and write each Markdown row on the fly.
     * <p>
     * When called with an {@link InputStream} (file is {@code null}), the stream is spooled to a
     * temporary file because two passes require seekable re-reading. The temporary file is deleted
     * on completion. For the {@link File} path the file is simply re-opened on the second pass.
     *
     * @param file           source file, or {@code null} when called from the InputStream path
     * @param inputStream    the CSV bytes (only used when file is {@code null})
     * @param sheetName      the sheet/section heading name
     * @param writer         destination writer, <em>not</em> closed by this method
     */
    private static void writeCsvStreaming(File file, InputStream inputStream, String sheetName,
        Writer writer) throws IOException {
        // InputStream path: spool to a temporary file so both passes can re-read it.
        File cacheDir = null;
        File csvFile = file;
        if (csvFile == null) {
            cacheDir = FileUtils.createCacheTmpFile();
            csvFile = new File(cacheDir, UUID.randomUUID() + ".csv");
            FileUtils.writeToFile(csvFile, inputStream, false);
        }
        try {
            // --- Pass 1: count max columns (no row storage) ---
            int maxColumns = countCsvColumns(csvFile);

            // --- Pass 2: write Markdown ---
            writer.write("## ");
            writer.write(sheetName == null ? "" : sheetName);
            writer.write("\n\n");

            if (maxColumns == 0) {
                writer.write("(empty sheet)\n");
                return;
            }

            try (InputStream in = Files.newInputStream(csvFile.toPath())) {
                BufferedInputStream buffered = new BufferedInputStream(in, CsvDetector.SNIFF_LIMIT);
                buffered.mark(CsvDetector.SNIFF_LIMIT);
                byte[] head = new byte[CsvDetector.SNIFF_LIMIT];
                int headLength = 0;
                int count;
                while (headLength < head.length
                    && (count = buffered.read(head, headLength, head.length - headLength)) != -1) {
                    headLength += count;
                }
                buffered.reset();
                Charset charset = CsvDetector.detectCharset(head);
                Reader decoded = new BomStrippingReader(new InputStreamReader(buffered, charset));
                BufferedReader lineReader = new BufferedReader(decoded, CsvDetector.SNIFF_LIMIT);
                char delimiter = CsvDetector.detectDelimiter(lineReader);
                try (CSVParser parser = new CSVParser(lineReader,
                    CSVFormat.DEFAULT.builder().setDelimiter(delimiter).build())) {
                    List<String> headerRow = null;
                    for (CSVRecord record : parser) {
                        List<String> row = new ArrayList<>();
                        for (String value : record) {
                            row.add(value);
                        }
                        if (headerRow == null) {
                            headerRow = row;
                            StringBuilder hb = new StringBuilder();
                            SheetTable.appendTableRowStatic(hb, headerRow, maxColumns, null);
                            SheetTable.appendTableSeparatorStatic(hb, maxColumns, null);
                            writer.write(hb.toString());
                        } else {
                            StringBuilder rb = new StringBuilder();
                            SheetTable.appendTableRowStatic(rb, row, maxColumns, null);
                            writer.write(rb.toString());
                        }
                    }
                    if (headerRow == null) {
                        // File had bytes but CSVParser produced no records
                        writer.write("(empty sheet)\n");
                    }
                }
            }
        } finally {
            if (cacheDir != null) {
                FileUtils.delete(cacheDir);
            }
        }
    }

    /**
     * Pass-1 helper: stream every CSV record and return the maximum column count without storing
     * any rows in memory. Returns 0 when the file is empty or contains no parseable records.
     */
    private static int countCsvColumns(File file) throws IOException {
        int maxColumns = 0;
        try (InputStream in = Files.newInputStream(file.toPath())) {
            BufferedInputStream buffered = new BufferedInputStream(in, CsvDetector.SNIFF_LIMIT);
            buffered.mark(CsvDetector.SNIFF_LIMIT);
            byte[] head = new byte[CsvDetector.SNIFF_LIMIT];
            int headLength = 0;
            int count;
            while (headLength < head.length
                && (count = buffered.read(head, headLength, head.length - headLength)) != -1) {
                headLength += count;
            }
            buffered.reset();
            Charset charset = CsvDetector.detectCharset(head);
            Reader decoded = new BomStrippingReader(new InputStreamReader(buffered, charset));
            BufferedReader lineReader = new BufferedReader(decoded, CsvDetector.SNIFF_LIMIT);
            char delimiter = CsvDetector.detectDelimiter(lineReader);
            try (CSVParser parser = new CSVParser(lineReader,
                CSVFormat.DEFAULT.builder().setDelimiter(delimiter).build())) {
                for (CSVRecord record : parser) {
                    int size = (int)record.size();
                    if (size > maxColumns) {
                        maxColumns = size;
                    }
                }
            }
        }
        return maxColumns;
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
        applyMergedRegions(table, regions, MIN_HEAD_ROWS);
    }

    /**
     * Fill every cell of a merged region with the value of its top-left cell. When
     * {@code headerRowCount} is greater than 1, cells that fall entirely within the header row
     * range (rows 0..headerRowCount-1) are <em>not</em> filled, so that multi-row header
     * flattening produces clean results without duplicate values from vertical merges.
     */
    private static void applyMergedRegions(SheetTable table, List<int[]> regions, int headerRowCount) {
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
            // For multi-row headers, skip merged cells that fall within header rows so that
            // flattenedHeaders() does not produce "X X" from a vertical merge of "X".
            int fillStartRow = firstRow;
            if (headerRowCount > MIN_HEAD_ROWS && firstRow < headerRowCount) {
                fillStartRow = headerRowCount;
            }
            for (int rowIndex = fillStartRow; rowIndex <= lastRow && rowIndex < table.rows.size(); rowIndex++) {
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
     * <p>
     * When HTML color rendering is enabled, the final value is wrapped with
     * {@code <span style="...">} containing font color and/or background color attributes.
     * HTML special characters in the cell value are escaped before wrapping.
     *
     * @param table
     *            the sheet table
     * @param styles
     *            cell style map keyed by {@code "origRow:origCol"}
     * @param newToOrigRow
     *            new data row index to original row index mapping, may be {@code null}
     * @param newToOrigCol
     *            new column index to original column index mapping, may be {@code null}
     */
    private static void applyCellStyles(SheetTable table, Map<String, CellStyle> styles,
        List<Integer> newToOrigRow, List<Integer> newToOrigCol) {
        if (styles == null || styles.isEmpty()) {
            return;
        }
        int headerOffset = table.headers.size();
        for (int rowIndex = 0; rowIndex < table.rows.size(); rowIndex++) {
            List<String> row = table.rows.get(rowIndex);
            int originalRowIndex = (newToOrigRow != null && rowIndex < newToOrigRow.size())
                ? newToOrigRow.get(rowIndex) : rowIndex + headerOffset;
            for (int colIndex = 0; colIndex < row.size(); colIndex++) {
                int originalColIndex = (newToOrigCol != null && colIndex < newToOrigCol.size())
                    ? newToOrigCol.get(colIndex) : colIndex;
                CellStyle style = styles.get(originalRowIndex + CELL_KEY_SEPARATOR + originalColIndex);
                if (style == null) {
                    continue;
                }
                String value = row.get(colIndex);
                if (value == null || value.isEmpty()) {
                    continue;
                }
                // 1. 先做 Markdown marker 包装（纯 GFM 字符）
                value = wrapWithFontMarkers(value, style);
                // 2. HTML 颜色渲染：转义 + 包 span
                if (htmlColorRendering) {
                    value = wrapWithColorSpan(value, style);
                }
                row.set(colIndex, value);
            }
        }
    }

    /**
     * Wrap a value with inline Markdown font markers. Combination order:
     * strike wraps bold wraps italic wraps value.
     * <p>
     * When the style carries rich text {@link CellStyle#runs}, per-segment wrapping is used
     * instead of whole-cell wrapping. Adjacent segments with identical flags are merged so
     * that uniform formatting across the entire cell produces the same output as whole-cell
     * wrapping (e.g. {@code **ab**} rather than {@code **a****b**}).
     */
    static String wrapWithFontMarkers(String value, CellStyle style) {
        if (style.runs != null && !style.runs.isEmpty()) {
            return wrapWithRichTextRuns(value, style.runs);
        }
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
     * Wrap a rich text value by its per-run style segments. Adjacent runs with identical
     * flags are merged before wrapping so that uniform formatting produces a single wrapped
     * span rather than adjacent identical spans. Runs are clamped to the value length to
     * guard against offset drift.
     */
    private static String wrapWithRichTextRuns(String value,
        List<CellStyle.RunStyle> runs) {
        int textLen = value.length();
        if (textLen == 0 || runs.isEmpty()) {
            return value;
        }

        // Build merged segments: adjacent runs with identical flags are coalesced.
        List<int[]> offsets = new ArrayList<>();
        List<boolean[]> flags = new ArrayList<>();
        for (int i = 0; i < runs.size(); i++) {
            CellStyle.RunStyle run = runs.get(i);
            int start = Math.min(run.start, textLen);
            int end = Math.min(run.end, textLen);
            if (start >= end) {
                continue;
            }
            if (!flags.isEmpty()) {
                boolean[] prev = flags.get(flags.size() - 1);
                if (prev[0] == run.bold && prev[1] == run.italic && prev[2] == run.strikeout) {
                    // Merge with previous segment: extend its end offset.
                    offsets.get(offsets.size() - 1)[1] = end;
                    continue;
                }
            }
            offsets.add(new int[] {start, end});
            flags.add(new boolean[] {run.bold, run.italic, run.strikeout});
        }

        // Wrap each merged segment and concatenate.
        StringBuilder sb = new StringBuilder(textLen);
        for (int i = 0; i < offsets.size(); i++) {
            int[] off = offsets.get(i);
            boolean[] f = flags.get(i);
            String segment = value.substring(off[0], off[1]);
            if (f[1]) {
                segment = "*" + segment + "*";
            }
            if (f[0]) {
                segment = "**" + segment + "**";
            }
            if (f[2]) {
                segment = "~~" + segment + "~~";
            }
            sb.append(segment);
        }
        return sb.toString();
    }

    // -------------------------------------------------------------------------
    // HTML 颜色渲染
    // -------------------------------------------------------------------------

    /**
     * 将已包装好 Markdown marker 的值用 HTML span 包裹颜色信息。
     * 顺序：先转义 HTML 敏感字符（保护值内容），再包 span（span 标签本身不受影响）。
     * <p>
     * 若 style 无颜色信息，原样返回。
     */
    static String wrapWithColorSpan(String value, CellStyle style) {
        String styleAttr = buildColorStyleAttr(style);
        if (styleAttr == null) {
            return value;
        }
        // 转义 HTML 敏感字符（发生在 marker 包装之后、span 包装之前）
        value = escapeHtml(value);
        return SPAN_STYLE_OPEN + styleAttr + "\">" + value + SPAN_CLOSE;
    }

    /**
     * 构建 CSS 颜色样式属性值，如 {@code color:#FF0000;background-color:#FFFF00}。
     * 仅输出非 null 的颜色，color 在前 background-color 在后。
     * 若无任何颜色返回 {@code null}。
     */
    private static String buildColorStyleAttr(CellStyle style) {
        boolean hasFont = style.fontColorHex != null;
        boolean hasBg = style.backgroundColorHex != null;
        if (!hasFont && !hasBg) {
            return null;
        }
        StringBuilder sb = new StringBuilder(48);
        if (hasFont) {
            sb.append(CSS_COLOR).append(style.fontColorHex);
        }
        if (hasBg) {
            if (hasFont) {
                sb.append(';');
            }
            sb.append(CSS_BG_COLOR).append(style.backgroundColorHex);
        }
        return sb.toString();
    }

    /**
     * 转义 HTML 敏感字符：{@code &} → {@code &amp;}，{@code <} → {@code &lt;}，
     * {@code >} → {@code &gt;}。
     * 仅在 HTML 颜色渲染开启时调用，防止值内容破坏 span 结构。
     */
    private static String escapeHtml(String value) {
        if (value == null) {
            return null;
        }
        // 先替换 &（避免对已转义的 &amp; 再次转义时产生 &amp;amp;）
        // 注意：此处只做一次替换，& → &amp; 后不会再被替换
        StringBuilder sb = null;
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            if (c == '&' || c == '<' || c == '>') {
                if (sb == null) {
                    sb = new StringBuilder(value.length() + 16);
                    sb.append(value, 0, i);
                }
                if (c == '&') {
                    sb.append("&amp;");
                } else if (c == '<') {
                    sb.append("&lt;");
                } else {
                    sb.append("&gt;");
                }
            } else if (sb != null) {
                sb.append(c);
            }
        }
        return sb == null ? value : sb.toString();
    }

    /**
     * Derive per-column alignment from the header row (row 0 before promotion) cell styles
     * and store the result in {@link SheetTable#columnAlignments}.
     *
     * @param table
     *            the sheet table
     * @param styles
     *            cell style map keyed by {@code "origRow:origCol"}
     * @param newToOrigCol
     *            new column index to original column index mapping, may be {@code null}
     */
    private static void applyColumnAlignments(SheetTable table, Map<String, CellStyle> styles,
        List<Integer> newToOrigCol) {
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
            int originalColIndex = (newToOrigCol != null && col < newToOrigCol.size())
                ? newToOrigCol.get(col) : col;
            CellStyle style = styles.get("0" + CELL_KEY_SEPARATOR + originalColIndex);
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

    /**
     * Move the first {@code headRowNumber} rows from {@code rows} to {@code headers}, preserving
     * their order. If fewer rows are available, all rows become headers.
     */
    private static void promoteRowsToHeader(SheetTable table, int headRowNumber) {
        int count = Math.min(headRowNumber, table.rows.size());
        table.headers.addAll(table.rows.subList(0, count));
        table.rows.subList(0, count).clear();
    }

    /**
     * 隐藏行列过滤的结果：新索引到原始索引的映射。
     * 供 {@link #applyCellStyles} 和 {@link #applyColumnAlignments} 使用。
     */
    private static final class HiddenFilterResult {
        /** 新数据行索引 → 原始行索引（过滤前）。{@code null} 表示未做过滤。 */
        final List<Integer> newToOrigRow;
        /** 新列索引 → 原始列索引（过滤前）。{@code null} 表示未做过滤。 */
        final List<Integer> newToOrigCol;

        HiddenFilterResult(List<Integer> newToOrigRow, List<Integer> newToOrigCol) {
            this.newToOrigRow = newToOrigRow;
            this.newToOrigCol = newToOrigCol;
        }
    }

    /**
     * 过滤隐藏的行和列。在 promoteFirstRowToHeader 之后调用。
     * <p>
     * 过滤后，返回新索引到原始索引的映射，供后续样式应用步骤使用
     * （因为样式 map 的 key 是原始行列索引）。
     *
     * @param table
     *            the sheet table (modified in-place)
     * @param hiddenInfo
     *            hidden row/col info, may be {@code null}
     * @return 新索引到原始索引的映射
     */
    private static HiddenFilterResult filterHiddenRowsAndColumns(SheetTable table,
        CellStyleScanner.HiddenInfo hiddenInfo) {
        if (hiddenInfo == null) {
            return new HiddenFilterResult(null, null);
        }
        Set<Integer> hiddenRows = hiddenInfo.hiddenRows;
        Set<Integer> hiddenCols = hiddenInfo.hiddenCols;
        boolean hasHiddenRows = hiddenRows != null && !hiddenRows.isEmpty();
        boolean hasHiddenCols = hiddenCols != null && !hiddenCols.isEmpty();
        if (!hasHiddenRows && !hasHiddenCols) {
            return new HiddenFilterResult(null, null);
        }

        int originalHeaderCount = table.headers.size();

        // 构建数据行的新→原始索引映射
        List<Integer> newToOrigRow = new ArrayList<>(table.rows.size());
        for (int i = 0; i < table.rows.size(); i++) {
            int origIdx = i + originalHeaderCount;
            if (!hasHiddenRows || !hiddenRows.contains(origIdx)) {
                newToOrigRow.add(origIdx);
            }
        }

        // 构建列的新→原始索引映射
        int maxCol = 0;
        for (List<String> row : table.rows) {
            maxCol = Math.max(maxCol, row.size());
        }
        for (List<String> header : table.headers) {
            maxCol = Math.max(maxCol, header.size());
        }
        List<Integer> newToOrigCol = new ArrayList<>(maxCol);
        for (int i = 0; i < maxCol; i++) {
            if (!hasHiddenCols || !hiddenCols.contains(i)) {
                newToOrigCol.add(i);
            }
        }

        // 过滤隐藏的表头行
        if (hasHiddenRows) {
            List<List<String>> filteredHeaders = new ArrayList<>(table.headers.size());
            for (int i = 0; i < table.headers.size(); i++) {
                if (!hiddenRows.contains(i)) {
                    filteredHeaders.add(table.headers.get(i));
                }
            }
            table.headers = filteredHeaders;
        }

        // 过滤隐藏的数据行
        if (hasHiddenRows) {
            List<List<String>> filteredRows = new ArrayList<>(table.rows.size());
            for (int i = 0; i < table.rows.size(); i++) {
                int origIdx = i + originalHeaderCount;
                if (!hiddenRows.contains(origIdx)) {
                    filteredRows.add(table.rows.get(i));
                }
            }
            table.rows = filteredRows;
        }

        // 过滤隐藏的列
        if (hasHiddenCols) {
            for (List<String> header : table.headers) {
                filterColumnsFromRow(header, hiddenCols);
            }
            for (List<String> row : table.rows) {
                filterColumnsFromRow(row, hiddenCols);
            }
        }

        return new HiddenFilterResult(newToOrigRow, newToOrigCol);
    }

    /**
     * 从一行中移除隐藏列对应的单元格（就地修改）。
     */
    private static void filterColumnsFromRow(List<String> row, Set<Integer> hiddenCols) {
        if (row == null || hiddenCols == null || hiddenCols.isEmpty()) {
            return;
        }
        List<String> filtered = new ArrayList<>(row.size());
        for (int i = 0; i < row.size(); i++) {
            if (!hiddenCols.contains(i)) {
                filtered.add(row.get(i));
            }
        }
        row.clear();
        row.addAll(filtered);
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
     * Both single-char {@link #read()} and bulk {@link #read(char[], int, int)} paths
     * are handled so that the BOM is never leaked to the caller regardless of read pattern.
     */
    static final class BomStrippingReader extends FilterReader {
        private boolean first = true;

        BomStrippingReader(Reader delegate) {
            super(delegate);
        }

        @Override
        public int read() throws IOException {
            int c = super.read();
            if (first) {
                first = false;
                if (c == UTF8_BOM) {
                    c = super.read();
                }
            }
            return c;
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
            if (count > 0) {
                first = false;
            }
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
