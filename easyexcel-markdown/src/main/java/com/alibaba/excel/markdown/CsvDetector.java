package com.alibaba.excel.markdown;

import java.io.IOException;
import java.io.Reader;
import java.nio.ByteBuffer;
import java.nio.charset.CharacterCodingException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

/**
 * Sniff the charset and the delimiter of CSV input: BOM first, then strict UTF-8 validation
 * with a GBK fallback (Excel on Chinese Windows writes ANSI, that is GBK); the delimiter is
 * the most frequent candidate seen outside double quotes on the first line.
 *
 * @author wandl
 */
final class CsvDetector {

    static final int SNIFF_LIMIT = 8192;
    private static final byte[] UTF8_BOM = {(byte)0xEF, (byte)0xBB, (byte)0xBF};
    private static final byte[] UTF16LE_BOM = {(byte)0xFF, (byte)0xFE};
    private static final byte[] UTF16BE_BOM = {(byte)0xFE, (byte)0xFF};
    private static final char[] DELIMITER_CANDIDATES = {',', ';', '\t', '|'};
    private static final char DEFAULT_DELIMITER = ',';
    private static final char QUOTE = '"';
    private static final char CR = '\r';
    private static final char LF = '\n';

    private CsvDetector() {}

    static Charset detectCharset(byte[] head) {
        if (startsWith(head, UTF8_BOM)) {
            return StandardCharsets.UTF_8;
        }
        if (startsWith(head, UTF16LE_BOM)) {
            return StandardCharsets.UTF_16LE;
        }
        if (startsWith(head, UTF16BE_BOM)) {
            return StandardCharsets.UTF_16BE;
        }
        try {
            StandardCharsets.UTF_8.newDecoder().decode(ByteBuffer.wrap(head));
            return StandardCharsets.UTF_8;
        } catch (CharacterCodingException e) {
            return Charset.forName("GBK");
        }
    }

    private static boolean startsWith(byte[] head, byte[] bom) {
        if (head.length < bom.length) {
            return false;
        }
        for (int i = 0; i < bom.length; i++) {
            if (head[i] != bom[i]) {
                return false;
            }
        }
        return true;
    }

    static char detectDelimiter(Reader reader) throws IOException {
        reader.mark(SNIFF_LIMIT);
        int[] counts = new int[DELIMITER_CANDIDATES.length];
        boolean quoted = false;
        int character;
        try {
            while ((character = reader.read()) != -1) {
                if (character == LF || character == CR) {
                    break;
                }
                if (character == QUOTE) {
                    quoted = !quoted;
                    continue;
                }
                if (quoted) {
                    continue;
                }
                for (int i = 0; i < DELIMITER_CANDIDATES.length; i++) {
                    if (character == DELIMITER_CANDIDATES[i]) {
                        counts[i]++;
                    }
                }
            }
        } finally {
            reader.reset();
        }
        int best = 0;
        for (int i = 1; i < counts.length; i++) {
            if (counts[i] > counts[best]) {
                best = i;
            }
        }
        return counts[best] > 0 ? DELIMITER_CANDIDATES[best] : DEFAULT_DELIMITER;
    }
}
