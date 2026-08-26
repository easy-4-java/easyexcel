package com.alibaba.excel.markdown;

/**
 * Lightweight value object capturing the subset of Excel cell styles that can be represented
 * in GitHub Flavored Markdown: inline font markers (bold / italic / strikeout) and horizontal
 * alignment for column alignment in the delimiter row.
 * <p>
 * Color, border, background and other rich formatting are intentionally absent because GFM
 * tables cannot represent them.
 *
 * @author wandl
 */
final class CellStyle {

    private static final String HORIZONTAL_GENERAL = "general";

    boolean bold;
    boolean italic;
    boolean strikeout;

    /**
     * Horizontal alignment as a lowercase string: {@code "left"}, {@code "center"},
     * {@code "right"}, {@code "general"}, or {@code null} when not set.
     */
    String horizontal;

    /**
     * Returns {@code true} when this style carries no Markdown-representable formatting:
     * all font flags are false and horizontal alignment is absent or {@code "general"}.
     */
    boolean isPlain() {
        return !bold && !italic && !strikeout
            && (horizontal == null || HORIZONTAL_GENERAL.equals(horizontal));
    }
}
