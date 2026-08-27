package com.alibaba.excel.markdown;

import java.util.List;

/**
 * Lightweight value object capturing the subset of Excel cell styles that can be represented
 * in GitHub Flavored Markdown: inline font markers (bold / italic / strikeout) and horizontal
 * alignment for column alignment in the delimiter row.
 * <p>
 * Color, border, background and other rich formatting are intentionally absent because GFM
 * tables cannot represent them.
 * <p>
 * Rich text support: when a cell contains mixed formatting across runs (e.g. partial bold),
 * the {@link #runs} list carries per-segment style information. When {@code runs} is present
 * and non-empty, the renderer uses run-level wrapping instead of the whole-cell flags
 * ({@code bold}, {@code italic}, {@code strikeout}).
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
     * Optional rich text run segments. When non-empty, the renderer applies per-segment
     * Markdown wrapping instead of whole-cell wrapping. May be {@code null} for plain cells.
     */
    List<RunStyle> runs;

    /**
     * Returns {@code true} when this style carries no Markdown-representable formatting:
     * all font flags are false, no rich text runs carry formatting, and horizontal alignment
     * is absent or {@code "general"}.
     */
    boolean isPlain() {
        if (bold || italic || strikeout) {
            return false;
        }
        if (horizontal != null && !HORIZONTAL_GENERAL.equals(horizontal)) {
            return false;
        }
        if (runs != null) {
            for (int i = 0; i < runs.size(); i++) {
                RunStyle run = runs.get(i);
                if (run.bold || run.italic || run.strikeout) {
                    return false;
                }
            }
        }
        return true;
    }

    /**
     * A single rich text run segment with its own inline style flags and character range.
     * The {@code start} is inclusive and {@code end} is exclusive, matching
     * {@link String#substring(int, int)} semantics.
     */
    static final class RunStyle {
        /** Inclusive start offset within the cell text. */
        final int start;
        /** Exclusive end offset within the cell text. */
        final int end;
        final boolean bold;
        final boolean italic;
        final boolean strikeout;

        RunStyle(int start, int end, boolean bold, boolean italic, boolean strikeout) {
            this.start = start;
            this.end = end;
            this.bold = bold;
            this.italic = italic;
            this.strikeout = strikeout;
        }
    }
}
