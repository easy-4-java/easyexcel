package com.alibaba.excel.markdown;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRPrElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Scan cell styles (bold / italic / strikeout / horizontal alignment) from XLSX or legacy XLS
 * files. Two static entry points return a sparse map keyed by {@code "row:col"} per sheet index.
 * <p>
 * Only cells whose styles are non-plain (i.e.&nbsp;carry at least one Markdown-representable
 * formatting attribute) are collected, keeping the map small for typical workbooks.
 * <p>
 * Rich text support: XLSX shared strings with run-level formatting (partial bold, etc.) are
 * parsed from the {@code sharedStrings.xml} via {@link XSSFReader#getSharedStringsTable()}.
 * Each shared string's {@link XSSFRichTextString} runs are inspected for bold/italic/strikeout
 * flags via the {@link CTRPrElt} XML bean API. Cells referencing such shared strings receive a
 * {@link CellStyle} whose {@link CellStyle#runs} list carries per-segment styling. Legacy XLS
 * rich text is extracted via {@link HSSFRichTextString#numFormattingRuns()}.
 *
 * @author wandl
 */
final class CellStyleScanner {

    private static final String CELL_KEY_SEPARATOR = ":";
    private static final String HORIZONTAL_GENERAL = "general";

    /** XML element name for the fonts container in styles.xml. */
    private static final String EL_FONTS = "fonts";
    /** XML element name for a single font definition in styles.xml. */
    private static final String EL_FONT = "font";
    /** XML element name for the bold flag inside a font. */
    private static final String EL_BOLD = "b";
    /** XML element name for the italic flag inside a font. */
    private static final String EL_ITALIC = "i";
    /** XML element name for the strikeout flag inside a font. */
    private static final String EL_STRIKE = "strike";
    /** XML element name for the cellXfs container in styles.xml. */
    private static final String EL_CELL_XFS = "cellXfs";
    /** XML element name for a single cell format entry in styles.xml. */
    private static final String EL_XF = "xf";
    /** XML element name for the alignment child of a cell format entry. */
    private static final String EL_ALIGNMENT = "alignment";
    /** XML element name for a cell element in a sheet XML. */
    private static final String EL_C = "c";

    /**
     * Shared SAX parser factory, thread-safe and reused across all style scans.
     */
    private static final SAXParserFactory SAX_FACTORY;

    static {
        SAX_FACTORY = SAXParserFactory.newInstance();
        SAX_FACTORY.setNamespaceAware(true);
    }

    private CellStyleScanner() {}

    /**
     * 每个 sheet 的隐藏行列信息，行/列索引均为 0-based。
     */
    static final class HiddenInfo {
        final Set<Integer> hiddenRows;
        final Set<Integer> hiddenCols;

        HiddenInfo(Set<Integer> hiddenRows, Set<Integer> hiddenCols) {
            this.hiddenRows = hiddenRows != null ? hiddenRows : Collections.<Integer>emptySet();
            this.hiddenCols = hiddenCols != null ? hiddenCols : Collections.<Integer>emptySet();
        }
    }

    /**
     * 样式扫描的完整结果：单元格样式 + 隐藏行列信息。
     */
    static final class ScanResult {
        final Map<Integer, Map<String, CellStyle>> styles;
        final Map<Integer, HiddenInfo> hiddenInfo;

        ScanResult(Map<Integer, Map<String, CellStyle>> styles, Map<Integer, HiddenInfo> hiddenInfo) {
            this.styles = styles;
            this.hiddenInfo = hiddenInfo;
        }
    }

    /**
     * 单个 sheet 的 SAX 解析结果：单元格样式 + 隐藏行列。
     */
    private static final class SheetParseResult {
        final Map<String, CellStyle> styles;
        final Set<Integer> hiddenRows;
        final Set<Integer> hiddenCols;

        SheetParseResult(Map<String, CellStyle> styles, Set<Integer> hiddenRows, Set<Integer> hiddenCols) {
            this.styles = styles;
            this.hiddenRows = hiddenRows;
            this.hiddenCols = hiddenCols;
        }
    }

    /**
     * Scan XLSX cell styles in a streaming fashion (SAX over styles.xml and each sheet).
     * Also parses sharedStrings.xml to extract rich text run-level formatting for cells
     * that reference shared strings with mixed formatting (e.g. partial bold).
     * 同时收集隐藏行/列信息。
     *
     * @param file
     *            an XLSX file (ZIP-based)
     * @return 包含单元格样式和隐藏行列信息的扫描结果
     * @throws IOException
     *             if the file cannot be read
     */
    static ScanResult scanStyles(File file) throws IOException {
        OPCPackage pkg = null;
        try {
            pkg = OPCPackage.open(file, PackageAccess.READ);
            XSSFReader reader = new XSSFReader(pkg);
            return scanStyles(reader);
        } catch (IOException e) {
            throw e;
        } catch (Exception e) {
            throw new IOException("Scan cell styles failure", e);
        } finally {
            if (pkg != null) {
                try {
                    pkg.close();
                } catch (IOException e) {
                    // closing a read only package best effort
                }
            }
        }
    }

    /**
     * Scan XLSX cell styles using an already-opened {@link XSSFReader}.
     * This overload allows the caller to share a single {@link OPCPackage} opening
     * across multiple scanners. 同时收集隐藏行/列信息。
     *
     * @param reader
     *            an already-opened XSSFReader (caller retains ownership)
     * @return 包含单元格样式和隐藏行列信息的扫描结果
     * @throws IOException
     *             if the styles or sheet data cannot be read
     */
    static ScanResult scanStyles(XSSFReader reader) throws IOException {
        Map<Integer, Map<String, CellStyle>> styles = new HashMap<>(16);
        Map<Integer, HiddenInfo> hiddenInfo = new HashMap<>(16);
        try {
            // 1. Parse styles.xml to build font and cellXfs lookup lists.
            List<FontFlags> fonts;
            List<Integer> fontIds;
            List<String> horizontals;
            InputStream stylesStream = reader.getStylesData();
            if (stylesStream != null) {
                try {
                    StylesInfo info = parseStyles(stylesStream);
                    fonts = info.fonts;
                    fontIds = info.fontIds;
                    horizontals = info.horizontals;
                } finally {
                    stylesStream.close();
                }
            } else {
                fonts = Collections.emptyList();
                fontIds = Collections.emptyList();
                horizontals = Collections.emptyList();
            }

            // 2. Parse sharedStrings.xml for rich text run-level formatting.
            List<List<CellStyle.RunStyle>> sharedStringRuns = parseSharedStringRuns(reader);

            // 3. For each sheet, SAX-parse the sheet XML to collect non-plain cells and hidden rows/cols.
            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            int sheetIndex = 0;
            while (sheets.hasNext()) {
                InputStream sheetStream = sheets.next();
                try {
                    SheetParseResult sheetResult = parseSheetStyles(
                        sheetStream, fonts, fontIds, horizontals, sharedStringRuns);
                    if (!sheetResult.styles.isEmpty()) {
                        styles.put(sheetIndex, sheetResult.styles);
                    }
                    if (!sheetResult.hiddenRows.isEmpty() || !sheetResult.hiddenCols.isEmpty()) {
                        hiddenInfo.put(sheetIndex, new HiddenInfo(sheetResult.hiddenRows, sheetResult.hiddenCols));
                    }
                } finally {
                    sheetStream.close();
                }
                sheetIndex++;
            }
        } catch (IOException e) {
            throw e;
        } catch (Exception e) {
            throw new IOException("Scan cell styles failure", e);
        }
        return new ScanResult(styles, hiddenInfo);
    }

    /**
     * Scan legacy XLS cell styles by loading the workbook into memory and enumerating every cell.
     * Acceptable because XLS is capped at 65536 rows x 256 columns.
     * 同时收集隐藏行/列信息（通过 {@link HSSFRow#getZeroHeight()} 和
     * {@link HSSFSheet#isColumnHidden(int)}）。
     * <p>
     * Rich text run extraction: cells with {@link HSSFRichTextString} that carry multiple
     * formatting runs are detected and their per-run bold/italic/strikeout flags are extracted
     * via {@link HSSFRichTextString#numFormattingRuns()},
     * {@link HSSFRichTextString#getIndexOfFormattingRun(int)} and
     * {@link HSSFFont#getBold()} / {@link HSSFFont#getItalic()} / {@link HSSFFont#getStrikeout()}.
     *
     * @param file
     *            an XLS file (OLE2-based)
     * @return 包含单元格样式和隐藏行列信息的扫描结果
     * @throws IOException
     *             if the file cannot be read
     */
    static ScanResult scanLegacyStyles(File file) throws IOException {
        Map<Integer, Map<String, CellStyle>> styles = new HashMap<>(16);
        Map<Integer, HiddenInfo> hiddenInfo = new HashMap<>(16);
        try (InputStream in = Files.newInputStream(file.toPath());
             HSSFWorkbook workbook = new HSSFWorkbook(in)) {
            int sheetCount = workbook.getNumberOfSheets();
            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
                Map<String, CellStyle> sheetStyles = new HashMap<>(16);
                Set<Integer> sheetHiddenRows = new HashSet<>(16);
                // 收集隐藏行和单元格样式
                int maxCol = 0;
                for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    HSSFRow row = sheet.getRow(rowIndex);
                    if (row == null) {
                        continue;
                    }
                    if (row.getZeroHeight()) {
                        sheetHiddenRows.add(rowIndex);
                    }
                    int lastCell = row.getLastCellNum();
                    if (lastCell > maxCol) {
                        maxCol = lastCell;
                    }
                    for (int colIndex = row.getFirstCellNum(); colIndex <= lastCell; colIndex++) {
                        HSSFCell cell = row.getCell(colIndex);
                        if (cell == null) {
                            continue;
                        }
                        CellStyle style = readHssfCellStyle(cell, workbook);
                        if (style != null && !style.isPlain()) {
                            sheetStyles.put(rowIndex + CELL_KEY_SEPARATOR + colIndex, style);
                        }
                    }
                }
                // 收集隐藏列
                Set<Integer> sheetHiddenCols = new HashSet<>(16);
                for (int colIndex = 0; colIndex < maxCol; colIndex++) {
                    if (sheet.isColumnHidden(colIndex)) {
                        sheetHiddenCols.add(colIndex);
                    }
                }
                if (!sheetStyles.isEmpty()) {
                    styles.put(sheetIndex, sheetStyles);
                }
                if (!sheetHiddenRows.isEmpty() || !sheetHiddenCols.isEmpty()) {
                    hiddenInfo.put(sheetIndex, new HiddenInfo(sheetHiddenRows, sheetHiddenCols));
                }
            }
        }
        return new ScanResult(styles, hiddenInfo);
    }

    /**
     * Read the style of a single HSSF cell and build a {@link CellStyle} if the cell has any
     * non-default formatting. Rich text cells with mixed run formatting are detected and their
     * per-run styles are populated in {@link CellStyle#runs}.
     */
    private static CellStyle readHssfCellStyle(HSSFCell cell, HSSFWorkbook workbook) {
        // Check for rich text run-level formatting first.
        List<CellStyle.RunStyle> runs = extractHssfRichTextRuns(cell, workbook);

        org.apache.poi.ss.usermodel.CellStyle poiStyle = cell.getCellStyle();
        if (poiStyle == null) {
            return runs != null ? buildRichTextOnlyStyle(runs) : null;
        }
        HSSFFont font = workbook.getFontAt(poiStyle.getFontIndex());
        boolean hasFont = font.getBold() || font.getItalic() || font.getStrikeout();
        HorizontalAlignment align = poiStyle.getAlignment();
        boolean hasAlign = align != null && align != HorizontalAlignment.GENERAL;
        if (!hasFont && !hasAlign && runs == null) {
            return null;
        }
        CellStyle style = new CellStyle();
        if (runs != null) {
            // Rich text: run-level flags take precedence over cell-level font flags.
            style.runs = runs;
        } else {
            style.bold = font.getBold();
            style.italic = font.getItalic();
            style.strikeout = font.getStrikeout();
        }
        if (hasAlign) {
            style.horizontal = align.name().toLowerCase(Locale.ROOT);
        }
        return style;
    }

    /**
     * Build a CellStyle that carries only horizontal alignment (if any) and no font flags,
     * used when a cell has rich text runs but no cell-level font formatting.
     */
    private static CellStyle buildRichTextOnlyStyle(List<CellStyle.RunStyle> runs) {
        CellStyle style = new CellStyle();
        style.runs = runs;
        return style;
    }

    /**
     * Extract rich text run-level formatting from an HSSF cell. Returns a non-empty list of
     * {@link CellStyle.RunStyle} if the cell contains a rich text string with at least one
     * formatting run that carries bold/italic/strikeout. Returns {@code null} if the cell does
     * not contain rich text or if no run has Markdown-representable formatting.
     */
    private static List<CellStyle.RunStyle> extractHssfRichTextRuns(HSSFCell cell,
        HSSFWorkbook workbook) {
        if (cell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING) {
            return null;
        }
        org.apache.poi.ss.usermodel.RichTextString rts = cell.getRichStringCellValue();
        if (!(rts instanceof HSSFRichTextString)) {
            return null;
        }
        HSSFRichTextString richText = (HSSFRichTextString) rts;
        int runCount = richText.numFormattingRuns();
        if (runCount <= 0) {
            return null;
        }
        String text = richText.getString();
        int textLen = text.length();
        List<CellStyle.RunStyle> runs = new ArrayList<>(runCount);
        boolean anyFormatted = false;
        for (int i = 0; i < runCount; i++) {
            int start = richText.getIndexOfFormattingRun(i);
            int end = (i + 1 < runCount) ? richText.getIndexOfFormattingRun(i + 1) : textLen;
            short fontIdx = richText.getFontOfFormattingRun(i);
            boolean bold = false;
            boolean italic = false;
            boolean strikeout = false;
            if (fontIdx != HSSFRichTextString.NO_FONT) {
                HSSFFont runFont = workbook.getFontAt(fontIdx);
                bold = runFont.getBold();
                italic = runFont.getItalic();
                strikeout = runFont.getStrikeout();
            }
            if (bold || italic || strikeout) {
                anyFormatted = true;
            }
            runs.add(new CellStyle.RunStyle(start, end, bold, italic, strikeout));
        }
        return anyFormatted ? runs : null;
    }

    // -------------------------------------------------------------------------
    // XLSX SAX parsing
    // -------------------------------------------------------------------------

    /**
     * Intermediate container for the two parallel lists extracted from styles.xml.
     */
    private static final class StylesInfo {
        final List<FontFlags> fonts;
        final List<Integer> fontIds;
        final List<String> horizontals;

        StylesInfo(List<FontFlags> fonts, List<Integer> fontIds, List<String> horizontals) {
            this.fonts = fonts;
            this.fontIds = fontIds;
            this.horizontals = horizontals;
        }
    }

    /**
     * Per-font flags extracted from the {@code <fonts>} section of styles.xml.
     */
    private static final class FontFlags {
        boolean bold;
        boolean italic;
        boolean strikeout;
    }

    /**
     * Parse styles.xml to extract font flags and cellXfs alignment info.
     */
    private static StylesInfo parseStyles(InputStream stylesStream) throws Exception {
        SAXParser parser = SAX_FACTORY.newSAXParser();
        XMLReader xmlReader = parser.getXMLReader();
        StylesHandler handler = new StylesHandler();
        xmlReader.setContentHandler(handler);
        xmlReader.parse(new InputSource(stylesStream));
        return new StylesInfo(handler.fonts, handler.fontIds, handler.horizontals);
    }

    /**
     * SAX handler for {@code /xl/styles.xml}: collects {@code <fonts>} and {@code <cellXfs>}.
     */
    private static final class StylesHandler extends DefaultHandler {
        private static final int INITIAL_CAPACITY = 32;

        private final List<FontFlags> fonts = new ArrayList<>(INITIAL_CAPACITY);
        private final List<Integer> fontIds = new ArrayList<>(INITIAL_CAPACITY);
        private final List<String> horizontals = new ArrayList<>(INITIAL_CAPACITY);

        private boolean inFonts;
        private boolean inCellXfs;
        private boolean inFont;
        private boolean inXf;
        private boolean hasAlignmentChild;
        private FontFlags currentFont;
        private int currentFontId;
        private boolean currentApplyAlignment;

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if (EL_FONTS.equals(localName)) {
                inFonts = true;
            } else if (EL_CELL_XFS.equals(localName)) {
                inCellXfs = true;
            } else if (inFonts && EL_FONT.equals(localName)) {
                inFont = true;
                currentFont = new FontFlags();
            } else if (inFont) {
                if (EL_BOLD.equals(localName)) {
                    currentFont.bold = true;
                } else if (EL_ITALIC.equals(localName)) {
                    currentFont.italic = true;
                } else if (EL_STRIKE.equals(localName)) {
                    currentFont.strikeout = true;
                }
            } else if (inCellXfs && EL_XF.equals(localName)) {
                inXf = true;
                hasAlignmentChild = false;
                String fontIdStr = attributes.getValue("fontId");
                currentFontId = parseInt(fontIdStr);
                String applyAlign = attributes.getValue("applyAlignment");
                currentApplyAlignment = "1".equals(applyAlign) || "true".equalsIgnoreCase(applyAlign);
            } else if (inXf && EL_ALIGNMENT.equals(localName)) {
                hasAlignmentChild = true;
                String horizontal = attributes.getValue("horizontal");
                if (currentApplyAlignment && horizontal != null) {
                    horizontals.add(horizontal.toLowerCase(Locale.ROOT));
                } else {
                    horizontals.add(null);
                }
                fontIds.add(currentFontId);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if (EL_FONTS.equals(localName)) {
                inFonts = false;
            } else if (EL_CELL_XFS.equals(localName)) {
                inCellXfs = false;
            } else if (inFonts && EL_FONT.equals(localName)) {
                fonts.add(currentFont);
                currentFont = null;
                inFont = false;
            } else if (inCellXfs && EL_XF.equals(localName)) {
                if (!hasAlignmentChild) {
                    fontIds.add(currentFontId);
                    horizontals.add(null);
                }
                inXf = false;
            }
        }

        private static int parseInt(String value) {
            if (value == null || value.isEmpty()) {
                return 0;
            }
            try {
                return Integer.parseInt(value.trim());
            } catch (NumberFormatException e) {
                return 0;
            }
        }
    }

    // -------------------------------------------------------------------------
    // XLSX sharedStrings rich text run parsing
    // -------------------------------------------------------------------------

    /**
     * Parse the shared strings table from the XLSX package and extract run-level formatting
     * for each shared string that contains rich text. Uses {@link XSSFReader#getSharedStringsTable()}
     * and the {@link XSSFRichTextString} API to iterate over formatting runs.
     * <p>
     * Returns a list indexed by shared string index. Entries are {@code null} for plain strings
     * (no run-level formatting) or non-{@code null} lists of {@link CellStyle.RunStyle} for
     * strings with Markdown-representable rich text formatting.
     *
     * @param reader the XSSFReader for the workbook
     * @return per-shared-string run style lists, never {@code null}
     */
    private static List<List<CellStyle.RunStyle>> parseSharedStringRuns(XSSFReader reader) {
        List<List<CellStyle.RunStyle>> result = new ArrayList<>();
        try {
            SharedStrings sst = reader.getSharedStringsTable();
            if (sst == null) {
                return result;
            }
            int count = sst.getCount();
            for (int i = 0; i < count; i++) {
                result.add(extractXssfRichTextRuns(sst.getItemAt(i)));
            }
        } catch (Exception e) {
            // Best effort: if shared strings parsing fails, rich text runs are simply absent
            // and the converter falls back to cell-level or plain rendering.
        }
        return result;
    }

    /**
     * Extract run-level formatting from a single shared string's {@link XSSFRichTextString}.
     * Returns a non-empty list of {@link CellStyle.RunStyle} when the string has at least one
     * run with bold/italic/strikeout. Returns {@code null} for plain strings.
     * <p>
     * Uses the {@link CTRPrElt} XML bean API via {@link XSSFRichTextString#getCTRst()} to
     * inspect run properties. The {@code <b/>}, {@code <i/>} and {@code <strike/>} elements
     * are detected by checking whether their respective lists on {@link CTRPrElt} are non-empty.
     */
    private static List<CellStyle.RunStyle> extractXssfRichTextRuns(
        org.apache.poi.ss.usermodel.RichTextString rts) {
        if (!(rts instanceof XSSFRichTextString)) {
            return null;
        }
        XSSFRichTextString xssfRts = (XSSFRichTextString) rts;
        if (!xssfRts.hasFormatting()) {
            return null;
        }
        CTRst ctRst = xssfRts.getCTRst();
        List<CTRElt> rList = ctRst.getRList();
        if (rList.isEmpty()) {
            return null;
        }
        List<CellStyle.RunStyle> runs = new ArrayList<>(rList.size());
        boolean anyFormatted = false;
        int offset = 0;
        for (int i = 0; i < rList.size(); i++) {
            CTRElt run = rList.get(i);
            String text = run.getT();
            int len = (text != null) ? text.length() : 0;
            boolean bold = false;
            boolean italic = false;
            boolean strikeout = false;
            if (run.isSetRPr()) {
                CTRPrElt rPr = run.getRPr();
                bold = !rPr.getBList().isEmpty();
                italic = !rPr.getIList().isEmpty();
                strikeout = !rPr.getStrikeList().isEmpty();
            }
            if (bold || italic || strikeout) {
                anyFormatted = true;
            }
            runs.add(new CellStyle.RunStyle(offset, offset + len, bold, italic, strikeout));
            offset += len;
        }
        return anyFormatted ? runs : null;
    }

    // -------------------------------------------------------------------------
    // Sheet XML parsing
    // -------------------------------------------------------------------------

    /**
     * SAX-parse a sheet XML to collect cells with non-plain styles.
     * For cells referencing shared strings ({@code t="s"}), rich text run styles are looked
     * up from the pre-parsed shared string runs list. For other cells, the cellXfs-based
     * font lookup is used.
     */
    private static SheetParseResult parseSheetStyles(InputStream sheetStream,
        List<FontFlags> fonts, List<Integer> fontIds, List<String> horizontals,
        List<List<CellStyle.RunStyle>> sharedStringRuns) throws Exception {
        SAXParser parser = SAX_FACTORY.newSAXParser();
        XMLReader xmlReader = parser.getXMLReader();
        SheetStyleHandler handler = new SheetStyleHandler(fonts, fontIds, horizontals,
            sharedStringRuns);
        xmlReader.setContentHandler(handler);
        xmlReader.parse(new InputSource(sheetStream));
        // 将 colHiddenState 转换为 hiddenCols 集合
        Set<Integer> hiddenCols = new HashSet<>(handler.colHiddenState.size());
        for (Map.Entry<Integer, Boolean> entry : handler.colHiddenState.entrySet()) {
            if (entry.getValue()) {
                hiddenCols.add(entry.getKey());
            }
        }
        return new SheetParseResult(handler.result, handler.hiddenRows, hiddenCols);
    }

    /**
     * SAX handler for a sheet XML: for each {@code <c>} element, reads the {@code s} attribute
     * to look up the cellXfs entry, then the font, and builds a {@link CellStyle} if non-plain.
     * <p>
     * For cells with {@code t="s"} (shared string reference), the {@code <v>} child element
     * text is captured via {@link #characters(char[], int, int)} to obtain the shared string
     * index, then rich text run styles are looked up from the pre-parsed list and take
     * precedence over cell-level font flags.
     */
    private static final class SheetStyleHandler extends DefaultHandler {
        private static final int INITIAL_CAPACITY = 64;
        private static final String ATTR_TYPE = "t";
        private static final String TYPE_SHARED_STRING = "s";
        private static final String EL_V = "v";
        private static final String EL_ROW = "row";
        private static final String EL_COL = "col";
        private static final String ATTR_HIDDEN = "hidden";
        private static final String ATTR_R = "r";
        private static final String ATTR_MIN = "min";
        private static final String ATTR_MAX = "max";
        private static final String HIDDEN_VALUE = "1";
        private static final String HIDDEN_VALUE_TRUE = "true";

        private final List<FontFlags> fonts;
        private final List<Integer> fontIds;
        private final List<String> horizontals;
        private final List<List<CellStyle.RunStyle>> sharedStringRuns;
        private final Map<String, CellStyle> result = new HashMap<>(INITIAL_CAPACITY);
        /** 隐藏行集合，0-based 行索引。 */
        private final Set<Integer> hiddenRows = new HashSet<>(INITIAL_CAPACITY);
        /** 每列的隐藏状态，用于正确处理 {@code <col>} 元素的覆盖语义。 */
        private final Map<Integer, Boolean> colHiddenState = new HashMap<>(INITIAL_CAPACITY);

        /** Set when a {@code <c t="s">} is encountered; cleared on {@code </c>}. */
        private String pendingSsKey;
        /** Saved {@code s} attribute from the current {@code <c>} for fallback lookup. */
        private String pendingStyleIdx;
        /** True while inside a {@code <v>} child of a shared-string {@code <c>}. */
        private boolean inV;
        /** Accumulates characters inside {@code <v>}. */
        private final StringBuilder vBuffer = new StringBuilder();

        SheetStyleHandler(List<FontFlags> fonts, List<Integer> fontIds,
            List<String> horizontals, List<List<CellStyle.RunStyle>> sharedStringRuns) {
            this.fonts = fonts;
            this.fontIds = fontIds;
            this.horizontals = horizontals;
            this.sharedStringRuns = sharedStringRuns;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if (EL_ROW.equals(localName)) {
                handleRowStart(attributes);
            } else if (EL_COL.equals(localName)) {
                handleColStart(attributes);
            } else if (EL_C.equals(localName)) {
                handleCellStart(attributes);
            } else if (EL_V.equals(localName) && pendingSsKey != null) {
                inV = true;
                vBuffer.setLength(0);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (inV) {
                vBuffer.append(ch, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if (EL_V.equals(localName) && inV) {
                inV = false;
                resolveSharedString();
            } else if (EL_C.equals(localName)) {
                pendingSsKey = null;
                pendingStyleIdx = null;
            }
        }

        /**
         * 处理 {@code <row>} 元素。若 hidden 属性为 "1" 或 "true"，将该行（0-based）加入隐藏集合。
         */
        private void handleRowStart(Attributes attributes) {
            String hiddenAttr = attributes.getValue(ATTR_HIDDEN);
            if (isHiddenValue(hiddenAttr)) {
                String rAttr = attributes.getValue(ATTR_R);
                if (rAttr != null) {
                    try {
                        int rowNum = Integer.parseInt(rAttr.trim());
                        if (rowNum > 0) {
                            hiddenRows.add(rowNum - 1);
                        }
                    } catch (NumberFormatException e) {
                        // 忽略无效的行号
                    }
                }
            }
        }

        /**
         * 处理 {@code <col>} 元素。若 hidden 属性为 "1" 或 "true"，将 min~max 范围内的列标记为隐藏。
         * 后出现的 {@code <col>} 元素会覆盖同列的先前设置（XLSX 覆盖语义）。
         */
        private void handleColStart(Attributes attributes) {
            String hiddenAttr = attributes.getValue(ATTR_HIDDEN);
            String minStr = attributes.getValue(ATTR_MIN);
            String maxStr = attributes.getValue(ATTR_MAX);
            if (minStr == null || maxStr == null) {
                return;
            }
            int min;
            int max;
            try {
                min = Integer.parseInt(minStr.trim());
                max = Integer.parseInt(maxStr.trim());
            } catch (NumberFormatException e) {
                return;
            }
            if (min <= 0 || max <= 0 || min > max) {
                return;
            }
            boolean isHidden = isHiddenValue(hiddenAttr);
            for (int col = min - 1; col <= max - 1; col++) {
                colHiddenState.put(col, isHidden);
            }
        }

        /**
         * 判断属性值是否表示隐藏状态（"1" 或 "true"）。
         */
        private static boolean isHiddenValue(String value) {
            return HIDDEN_VALUE.equals(value) || HIDDEN_VALUE_TRUE.equals(value);
        }

        /**
         * Process the start of a {@code <c>} element. If it references a shared string
         * ({@code t="s"}), the key is stored for deferred resolution when the {@code <v>}
         * text content arrives. Otherwise, cellXfs-based font lookup is done immediately.
         */
        private void handleCellStart(Attributes attributes) {
            String ref = attributes.getValue("r");
            String styleIdx = attributes.getValue("s");
            String typeAttr = attributes.getValue(ATTR_TYPE);
            if (ref == null) {
                return;
            }

            CellReference cellRef = new CellReference(ref);
            String key = cellRef.getRow() + CELL_KEY_SEPARATOR + cellRef.getCol();

            // If this is a shared string cell, defer resolution until <v> text arrives.
            if (TYPE_SHARED_STRING.equals(typeAttr) && sharedStringRuns != null
                && !sharedStringRuns.isEmpty()) {
                pendingSsKey = key;
                pendingStyleIdx = styleIdx;
                return;
            }

            // Otherwise, cellXfs-based font lookup (whole-cell style).
            applyCellStyle(attributes, key, styleIdx);
        }

        /**
         * Resolve the shared string index from the accumulated {@code <v>} text and look up
         * rich text run styles. If the shared string has run-level formatting, a
         * {@link CellStyle} with runs is emitted; otherwise, falls back to cellXfs-based
         * font lookup using the {@code s} attribute saved from {@code <c>}.
         */
        private void resolveSharedString() {
            String key = pendingSsKey;
            if (key == null) {
                return;
            }
            String vText = vBuffer.toString().trim();
            if (!vText.isEmpty()) {
                try {
                    int ssIndex = Integer.parseInt(vText);
                    if (ssIndex >= 0 && ssIndex < sharedStringRuns.size()) {
                        List<CellStyle.RunStyle> runs = sharedStringRuns.get(ssIndex);
                        if (runs != null && !runs.isEmpty()) {
                            CellStyle style = new CellStyle();
                            style.runs = runs;
                            result.put(key, style);
                            return;
                        }
                    }
                } catch (NumberFormatException e) {
                    // fall through to cellXfs-based lookup
                }
            }
            // No rich text runs found; fall back to cellXfs-based font lookup.
            if (pendingStyleIdx != null) {
                int xfIndex;
                try {
                    xfIndex = Integer.parseInt(pendingStyleIdx.trim());
                } catch (NumberFormatException e) {
                    return;
                }
                if (xfIndex >= 0 && xfIndex < fontIds.size()) {
                    int fontId = fontIds.get(xfIndex);
                    FontFlags flags = (fontId >= 0 && fontId < fonts.size())
                        ? fonts.get(fontId) : null;
                    String horizontal = (xfIndex < horizontals.size())
                        ? horizontals.get(xfIndex) : null;
                    boolean bold = flags != null && flags.bold;
                    boolean italic = flags != null && flags.italic;
                    boolean strikeout = flags != null && flags.strikeout;
                    boolean hasAlign = horizontal != null
                        && !HORIZONTAL_GENERAL.equals(horizontal);
                    if (bold || italic || strikeout || hasAlign) {
                        CellStyle style = new CellStyle();
                        style.bold = bold;
                        style.italic = italic;
                        style.strikeout = strikeout;
                        style.horizontal = hasAlign ? horizontal : null;
                        result.put(key, style);
                    }
                }
            }
        }

        /**
         * Apply cellXfs-based font lookup for a non-shared-string cell. Reads the {@code s}
         * attribute to find the font and alignment, and emits a {@link CellStyle} if non-plain.
         */
        private void applyCellStyle(Attributes attributes, String key, String styleIdx) {
            if (styleIdx == null) {
                return;
            }
            int xfIndex;
            try {
                xfIndex = Integer.parseInt(styleIdx.trim());
            } catch (NumberFormatException e) {
                return;
            }
            if (xfIndex < 0 || xfIndex >= fontIds.size()) {
                return;
            }

            int fontId = fontIds.get(xfIndex);
            FontFlags flags = (fontId >= 0 && fontId < fonts.size()) ? fonts.get(fontId) : null;
            String horizontal = (xfIndex < horizontals.size()) ? horizontals.get(xfIndex) : null;

            boolean bold = flags != null && flags.bold;
            boolean italic = flags != null && flags.italic;
            boolean strikeout = flags != null && flags.strikeout;
            boolean hasAlign = horizontal != null && !HORIZONTAL_GENERAL.equals(horizontal);

            if (!bold && !italic && !strikeout && !hasAlign) {
                return;
            }

            CellStyle style = new CellStyle();
            style.bold = bold;
            style.italic = italic;
            style.strikeout = strikeout;
            style.horizontal = hasAlign ? horizontal : null;
            result.put(key, style);
        }
    }
}
