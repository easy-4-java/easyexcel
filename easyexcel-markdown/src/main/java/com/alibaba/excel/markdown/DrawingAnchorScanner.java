package com.alibaba.excel.markdown;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.xml.sax.SAXException;

import org.apache.poi.hssf.usermodel.HSSFAnchor;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFShapeContainer;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Scan the drawing parts of an XLSX or legacy XLS file and report, per sheet index, the cells
 * where a picture is anchored. The easyexcel SAX read path cannot see drawings, so this scanner
 * is what makes image placeholders possible.
 *
 * @author wandl
 */
final class DrawingAnchorScanner {

    private static final String ELEMENT_TWO_CELL_ANCHOR = "twoCellAnchor";
    private static final String ELEMENT_ONE_CELL_ANCHOR = "oneCellAnchor";
    private static final String ELEMENT_FROM = "from";
    private static final String ELEMENT_COL = "col";
    private static final String ELEMENT_ROW = "row";
    private static final String ELEMENT_PIC = "pic";
    private static final String CELL_KEY_SEPARATOR = ":";

    /**
     * Shared SAX parser factory, created once and reused across all drawing scans.
     */
    private static final SAXParserFactory SAX_FACTORY;

    static {
        SAX_FACTORY = SAXParserFactory.newInstance();
        SAX_FACTORY.setNamespaceAware(true);
        // XXE 防护：禁用 DOCTYPE 声明与外部实体解析。恶意 xlsx 的 drawing XML
        // 可内嵌外部实体引用（读本地文件 / SSRF）；失败即抛，不做静默降级。
        try {
            SAX_FACTORY.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            SAX_FACTORY.setFeature("http://xml.org/sax/features/external-general-entities", false);
            SAX_FACTORY.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        } catch (ParserConfigurationException | SAXException e) {
            throw new ExceptionInInitializerError(
                "Unable to configure XXE protection for drawing scanner: " + e.getMessage());
        }
    }

    private DrawingAnchorScanner() {}

    /**
     * Scan XLSX (ZIP-based) drawing anchors in a streaming fashion.
     *
     * @param file
     *            an XLSX file (ZIP-based)
     * @return sheet index to the set of {@code rowIndex:columnIndex} anchor cells that carry a
     *         picture, never {@code null}
     * @throws IOException
     *             if the file cannot be read
     */
    static Map<Integer, Set<String>> scanPictureAnchors(File file) throws IOException {
        OPCPackage pkg = null;
        try {
            pkg = OPCPackage.open(file, PackageAccess.READ);
            XSSFReader reader = new XSSFReader(pkg);
            return scanPictureAnchors(reader);
        } catch (IOException e) {
            throw e;
        } catch (Exception e) {
            throw new IOException("Scan picture anchors failure", e);
        } finally {
            if (pkg != null) {
                try {
                    pkg.close();
                } catch (IOException e) {
                    // closing a read only package best effort, nothing left to clean up
                }
            }
        }
    }

    /**
     * Scan XLSX drawing anchors using an already-opened {@link XSSFReader}.
     * This overload allows the caller to share a single {@link OPCPackage} opening
     * across multiple scanners.
     *
     * @param reader
     *            an already-opened XSSFReader (caller retains ownership)
     * @return sheet index to the set of {@code rowIndex:columnIndex} anchor cells that carry a
     *         picture, never {@code null}
     * @throws IOException
     *             if the drawing parts cannot be read
     */
    static Map<Integer, Set<String>> scanPictureAnchors(XSSFReader reader) throws IOException {
        Map<Integer, Set<String>> anchors = new HashMap<>(16);
        try {
            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator)reader.getSheetsData();
            int sheetIndex = 0;
            while (sheets.hasNext()) {
                InputStream sheetStream = sheets.next();
                try {
                    PackagePart sheetPart = sheets.getSheetPart();
                    if (sheetPart != null) {
                        collectSheetAnchors(sheetPart, anchors, sheetIndex);
                    }
                } finally {
                    sheetStream.close();
                }
                sheetIndex++;
            }
        } catch (IOException e) {
            throw e;
        } catch (Exception e) {
            throw new IOException("Scan picture anchors failure", e);
        }
        return anchors;
    }

    /**
     * Scan legacy XLS (OLE2) drawing anchors by loading the workbook into memory and enumerating
     * every sheet's drawing patriarch. Acceptable because XLS is capped at 65536 rows x 256
     * columns ({@code SpreadsheetVersion.EXCEL97}).
     *
     * @return sheet index to the set of {@code rowIndex:columnIndex} anchor cells that carry a
     *         picture, never {@code null}
     */
    static Map<Integer, Set<String>> scanLegacyPictureAnchors(File file) throws IOException {
        Map<Integer, Set<String>> anchors = new HashMap<>(16);
        try (InputStream in = Files.newInputStream(file.toPath());
             HSSFWorkbook workbook = new HSSFWorkbook(in)) {
            int sheetCount = workbook.getNumberOfSheets();
            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                HSSFPatriarch patriarch = workbook.getSheetAt(sheetIndex).getDrawingPatriarch();
                if (patriarch == null) {
                    continue;
                }
                Set<String> cells = new HashSet<>(16);
                collectHssfPictures(patriarch.getChildren(), cells);
                if (!cells.isEmpty()) {
                    anchors.put(sheetIndex, cells);
                }
            }
        }
        return anchors;
    }

    /**
     * Recursively collect picture anchor positions from a list of HSSF shapes.
     * Shapes may be nested inside {@link HSSFShapeGroup} containers.
     */
    private static void collectHssfPictures(List<HSSFShape> shapes, Set<String> cells) {
        for (HSSFShape shape : shapes) {
            if (shape instanceof HSSFPicture) {
                HSSFAnchor anchor = shape.getAnchor();
                if (anchor instanceof HSSFClientAnchor) {
                    HSSFClientAnchor clientAnchor = (HSSFClientAnchor) anchor;
                    short col1 = clientAnchor.getCol1();
                    int row1 = clientAnchor.getRow1();
                    if (col1 >= 0 && row1 >= 0) {
                        cells.add(row1 + CELL_KEY_SEPARATOR + col1);
                    }
                }
            }
            if (shape instanceof HSSFShapeContainer) {
                collectHssfPictures(((HSSFShapeContainer) shape).getChildren(), cells);
            }
        }
    }

    private static void collectSheetAnchors(PackagePart sheetPart, Map<Integer, Set<String>> anchors,
        int sheetIndex) throws Exception {
        Set<String> cells = null;
        for (PackageRelationship relationship : sheetPart.getRelationships()) {
            if (!XSSFRelation.DRAWINGS.getRelation().equals(relationship.getRelationshipType())) {
                continue;
            }
            PackagePart drawingPart = sheetPart.getRelatedPart(relationship);
            if (cells == null) {
                cells = new HashSet<>(16);
            }
            parseDrawing(drawingPart, cells);
        }
        if (cells != null && !cells.isEmpty()) {
            anchors.put(sheetIndex, cells);
        }
    }

    private static void parseDrawing(PackagePart drawingPart, Set<String> cells) throws Exception {
        SAXParser parser = SAX_FACTORY.newSAXParser();
        XMLReader xmlReader = parser.getXMLReader();
        xmlReader.setContentHandler(new AnchorHandler(cells));
        try (InputStream drawingStream = drawingPart.getInputStream()) {
            xmlReader.parse(new InputSource(drawingStream));
        }
    }

    /**
     * Records the {@code from} cell of every anchor that contains a {@code pic} element.
     */
    private static final class AnchorHandler extends DefaultHandler {
        private final Set<String> cells;
        private final StringBuilder text = new StringBuilder();
        private boolean inFrom;
        private boolean inCol;
        private boolean inRow;
        private boolean sawPic;
        private int currentCol = -1;
        private int currentRow = -1;

        AnchorHandler(Set<String> cells) {
            this.cells = cells;
        }

        @Override
        public void startElement(String uri, String localName, String qName, org.xml.sax.Attributes attributes) {
            if (ELEMENT_TWO_CELL_ANCHOR.equals(localName) || ELEMENT_ONE_CELL_ANCHOR.equals(localName)) {
                sawPic = false;
            } else if (ELEMENT_FROM.equals(localName)) {
                inFrom = true;
                currentCol = -1;
                currentRow = -1;
            } else if (inFrom && ELEMENT_COL.equals(localName)) {
                inCol = true;
                text.setLength(0);
            } else if (inFrom && ELEMENT_ROW.equals(localName)) {
                inRow = true;
                text.setLength(0);
            } else if (ELEMENT_PIC.equals(localName)) {
                sawPic = true;
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (inCol || inRow) {
                text.append(ch, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if (ELEMENT_FROM.equals(localName)) {
                inFrom = false;
            } else if (inCol && ELEMENT_COL.equals(localName)) {
                currentCol = parse(text);
                inCol = false;
            } else if (inRow && ELEMENT_ROW.equals(localName)) {
                currentRow = parse(text);
                inRow = false;
            } else if (ELEMENT_TWO_CELL_ANCHOR.equals(localName)
                || ELEMENT_ONE_CELL_ANCHOR.equals(localName)) {
                if (sawPic && currentRow >= 0 && currentCol >= 0) {
                    cells.add(currentRow + CELL_KEY_SEPARATOR + currentCol);
                }
            }
        }

        private int parse(StringBuilder value) {
            try {
                return Integer.parseInt(value.toString().trim());
            } catch (NumberFormatException e) {
                return -1;
            }
        }
    }
}
