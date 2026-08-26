package com.alibaba.excel.markdown;

import java.io.File;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfSystemProperty;
import org.junit.jupiter.api.io.TempDir;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Manual benchmark, skipped by default so that neither CI nor the regular test run pays for it:
 * run it explicitly with {@code ./mvnw -pl easyexcel-markdown test -Dmaven.test.skip=false
 * -Dtest='MarkdownBenchmarkTest' -Dbenchmark=true
 * -DargLine="--add-opens=java.base/java.lang=ALL-UNNAMED"}.
 * <p>
 * Each size is warmed up once, then measured several times and the median is reported, so the
 * numbers are free of first-run JIT noise. Heap deltas are approximate
 * {@code System.gc() + totalMemory - freeMemory} snapshots, good enough to see the order of
 * magnitude of the full in-memory materialization, which is the known scaling limit of the
 * return-everything-as-String API.
 *
 * @author wandl
 */
@EnabledIfSystemProperty(named = "benchmark", matches = "true")
public class MarkdownBenchmarkTest {

    private static final int COLUMNS = 8;
    private static final int[] ROWS_PER_SIZE = {10_000, 50_000, 200_000};
    private static final int WARMUP_ITERATIONS = 1;
    private static final int MEASURED_ITERATIONS = 3;
    private static final int WRITE_BATCH_ROWS = 10_000;

    @TempDir
    File tempDir;

    @Test
    public void conversionThroughputAndHeap() throws Exception {
        System.out.println("[benchmark] | rows | cells | file MB | drawing scan ms | struct ms (median) "
            + "| struct heap MB | markdown ms (median) | markdown MB |");
        for (int rows : ROWS_PER_SIZE) {
            benchmarkSize(rows);
        }
    }

    private void benchmarkSize(int rows) throws Exception {
        File file = generate(rows);

        long scanStart = System.nanoTime();
        DrawingAnchorScanner.scanPictureAnchors(file);
        long scanMillis = (System.nanoTime() - scanStart) / 1_000_000;

        for (int i = 0; i < WARMUP_ITERATIONS; i++) {
            SheetDocument warmup = XlsxToMarkdownConverter.toStructured(file);
            warmup.toMarkdown();
        }

        long[] structMillis = new long[MEASURED_ITERATIONS];
        long[] markdownMillis = new long[MEASURED_ITERATIONS];
        long[] structHeapMb = new long[MEASURED_ITERATIONS];
        long[] markdownLength = new long[MEASURED_ITERATIONS];
        for (int i = 0; i < MEASURED_ITERATIONS; i++) {
            // the iteration keeps its document and markdown in method locals, so the heap
            // snapshot of the next iteration really measures one conversion at a time
            long before = usedHeapMb();
            long structStart = System.nanoTime();
            SheetDocument document = XlsxToMarkdownConverter.toStructured(file);
            structMillis[i] = (System.nanoTime() - structStart) / 1_000_000;
            long afterStruct = usedHeapMb();

            long markdownStart = System.nanoTime();
            String markdown = document.toMarkdown();
            markdownMillis[i] = (System.nanoTime() - markdownStart) / 1_000_000;
            markdownLength[i] = markdown.length();
            structHeapMb[i] = afterStruct - before;
            assertThat(markdown).isNotEmpty();
        }

        SheetDocument sanity = XlsxToMarkdownConverter.toStructured(file);
        assertThat(sanity.sheets).hasSize(1);
        assertThat(sanity.sheets.get(0).rows).hasSize(rows);

        System.out.printf("[benchmark] | %d | %d | %.1f | %d | %d | %d | %d | %.1f |%n",
            rows, (long)rows * COLUMNS, Files.size(file.toPath()) / 1024.0 / 1024.0, scanMillis,
            median(structMillis), median(structHeapMb), median(markdownMillis),
            median(markdownLength) / 1024.0 / 1024.0);
    }

    private File generate(int rows) throws Exception {
        File file = File.createTempFile("markdown-benchmark", ".xlsx", tempDir);
        List<List<String>> head = new ArrayList<>();
        for (int columnIndex = 0; columnIndex < COLUMNS; columnIndex++) {
            head.add(Collections.singletonList("H" + columnIndex));
        }
        ExcelWriter writer = EasyExcel.write(file).head(head).build();
        try {
            for (int from = 0; from < rows; from += WRITE_BATCH_ROWS) {
                List<List<Object>> data = new ArrayList<>();
                for (int rowIndex = from; rowIndex < Math.min(from + WRITE_BATCH_ROWS, rows); rowIndex++) {
                    List<Object> row = new ArrayList<>(COLUMNS);
                    for (int columnIndex = 0; columnIndex < COLUMNS; columnIndex++) {
                        row.add("cell-" + rowIndex + "-" + columnIndex);
                    }
                    data.add(row);
                }
                writer.write(data, EasyExcel.writerSheet(0, "Perf").build());
            }
        } finally {
            writer.finish();
        }
        return file;
    }

    private static long median(long[] values) {
        long[] sorted = values.clone();
        Arrays.sort(sorted);
        return sorted[sorted.length / 2];
    }

    private static long usedHeapMb() {
        for (int i = 0; i < 3; i++) {
            System.gc();
            try {
                Thread.sleep(50);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        }
        Runtime runtime = Runtime.getRuntime();
        return (runtime.totalMemory() - runtime.freeMemory()) / 1024 / 1024;
    }
}
