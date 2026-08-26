package com.alibaba.excel.markdown;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
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
     * Scan XLSX cell styles in a streaming fashion (SAX over styles.xml and each sheet).
     *
     * @param file
     *            an XLSX file (ZIP-based)
     * @return per-sheet map of {@code "row:col"} to non-plain {@link CellStyle}, never {@code null}
     * @throws IOException
     *             if the file cannot be read
     */
    static Map<Integer, Map<String, CellStyle>> scanStyles(File file) throws IOException {
        Map<Integer, Map<String, CellStyle>> result = new HashMap<>(16);
        OPCPackage pkg = null;
        try {
            pkg = OPCPackage.open(file, PackageAccess.READ);
            XSSFReader reader = new XSSFReader(pkg);

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

            // 2. For each sheet, SAX-parse the sheet XML to collect non-plain cells.
            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            int sheetIndex = 0;
            while (sheets.hasNext()) {
                InputStream sheetStream = sheets.next();
                try {
                    Map<String, CellStyle> sheetStyles = parseSheetStyles(sheetStream, fonts, fontIds, horizontals);
                    if (!sheetStyles.isEmpty()) {
                        result.put(sheetIndex, sheetStyles);
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
        } finally {
            if (pkg != null) {
                try {
                    pkg.close();
                } catch (IOException e) {
                    // closing a read only package best effort
                }
            }
        }
        return result;
    }

    /**
     * Scan legacy XLS cell styles by loading the workbook into memory and enumerating every cell.
     * Acceptable because XLS is capped at 65536 rows x 256 columns.
     *
     * @param file
     *            an XLS file (OLE2-based)
     * @return per-sheet map of {@code "row:col"} to non-plain {@link CellStyle}, never {@code null}
     * @throws IOException
     *             if the file cannot be read
     */
    static Map<Integer, Map<String, CellStyle>> scanLegacyStyles(File file) throws IOException {
        Map<Integer, Map<String, CellStyle>> result = new HashMap<>(16);
        try (InputStream in = Files.newInputStream(file.toPath());
             HSSFWorkbook workbook = new HSSFWorkbook(in)) {
            int sheetCount = workbook.getNumberOfSheets();
            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
                Map<String, CellStyle> sheetStyles = new HashMap<>(16);
                for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    HSSFRow row = sheet.getRow(rowIndex);
                    if (row == null) {
                        continue;
                    }
                    for (int colIndex = row.getFirstCellNum(); colIndex <= row.getLastCellNum(); colIndex++) {
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
                if (!sheetStyles.isEmpty()) {
                    result.put(sheetIndex, sheetStyles);
                }
            }
        }
        return result;
    }

    /**
     * Read the style of a single HSSF cell and build a {@link CellStyle} if the cell has any
     * non-default formatting.
     */
    private static CellStyle readHssfCellStyle(HSSFCell cell, HSSFWorkbook workbook) {
        org.apache.poi.ss.usermodel.CellStyle poiStyle = cell.getCellStyle();
        if (poiStyle == null) {
            return null;
        }
        HSSFFont font = workbook.getFontAt(poiStyle.getFontIndex());
        boolean hasFont = font.getBold() || font.getItalic() || font.getStrikeout();
        HorizontalAlignment align = poiStyle.getAlignment();
        boolean hasAlign = align != null && align != HorizontalAlignment.GENERAL;
        if (!hasFont && !hasAlign) {
            return null;
        }
        CellStyle style = new CellStyle();
        style.bold = font.getBold();
        style.italic = font.getItalic();
        style.strikeout = font.getStrikeout();
        if (hasAlign) {
            style.horizontal = align.name().toLowerCase(Locale.ROOT);
        }
        return style;
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

    /**
     * SAX-parse a sheet XML to collect cells with non-plain styles.
     */
    private static Map<String, CellStyle> parseSheetStyles(InputStream sheetStream,
        List<FontFlags> fonts, List<Integer> fontIds, List<String> horizontals) throws Exception {
        SAXParser parser = SAX_FACTORY.newSAXParser();
        XMLReader xmlReader = parser.getXMLReader();
        SheetStyleHandler handler = new SheetStyleHandler(fonts, fontIds, horizontals);
        xmlReader.setContentHandler(handler);
        xmlReader.parse(new InputSource(sheetStream));
        return handler.result;
    }

    /**
     * SAX handler for a sheet XML: for each {@code <c>} element, reads the {@code s} attribute
     * to look up the cellXfs entry, then the font, and builds a {@link CellStyle} if non-plain.
     */
    private static final class SheetStyleHandler extends DefaultHandler {
        private static final int INITIAL_CAPACITY = 64;

        private final List<FontFlags> fonts;
        private final List<Integer> fontIds;
        private final List<String> horizontals;
        private final Map<String, CellStyle> result = new HashMap<>(INITIAL_CAPACITY);

        SheetStyleHandler(List<FontFlags> fonts, List<Integer> fontIds, List<String> horizontals) {
            this.fonts = fonts;
            this.fontIds = fontIds;
            this.horizontals = horizontals;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if (!EL_C.equals(localName)) {
                return;
            }
            String ref = attributes.getValue("r");
            String styleIdx = attributes.getValue("s");
            if (ref == null || styleIdx == null) {
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

            CellReference cellRef = new CellReference(ref);
            String key = cellRef.getRow() + CELL_KEY_SEPARATOR + cellRef.getCol();

            CellStyle style = new CellStyle();
            style.bold = bold;
            style.italic = italic;
            style.strikeout = strikeout;
            style.horizontal = hasAlign ? horizontal : null;
            result.put(key, style);
        }
    }
}
