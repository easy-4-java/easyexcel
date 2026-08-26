package com.alibaba.excel.markdown;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

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
 * Scan the drawing parts of an XLSX file in a streaming way and report, per sheet index, the
 * cells where a picture is anchored. The easyexcel SAX read path cannot see drawings, so this
 * scanner is what makes image placeholders possible. Legacy XLS drawings are not covered.
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

    private DrawingAnchorScanner() {}

    /**
     * @return sheet index to the set of {@code rowIndex:columnIndex} anchor cells that carry a
     *         picture, never {@code null}
     */
    static Map<Integer, Set<String>> scanPictureAnchors(File file) throws IOException {
        Map<Integer, Set<String>> anchors = new HashMap<>(16);
        OPCPackage pkg = null;
        try {
            pkg = OPCPackage.open(file, PackageAccess.READ);
            XSSFReader reader = new XSSFReader(pkg);
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
        } finally {
            if (pkg != null) {
                try {
                    pkg.close();
                } catch (IOException e) {
                    // closing a read only package best effort, nothing left to clean up
                }
            }
        }
        return anchors;
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
        SAXParserFactory factory = SAXParserFactory.newInstance();
        factory.setNamespaceAware(true);
        SAXParser parser = factory.newSAXParser();
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
