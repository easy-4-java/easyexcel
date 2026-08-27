# easyexcel-markdown 完美转换增强计划（缺口 1–5 + 8）

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 补齐 `easyexcel-markdown` 模块已知的 6 个转换缺口——超链接渲染、批注渲染、图片占位、CSV 字符集探测、CSV 分隔符嗅探、中部空行保留——使"数据无损 + 结构近似"成立。

**Architecture:** 全部改动限于 `easyexcel-markdown` 模块。超链接/批注复用 easyexcel 的 `extraRead` 机制（已验证：`extraRead` 为 Set 叠加语义，`HyperlinkTagHandler` 产出 `CellExtra(HYPERLINK, url, ref)`、`XlsxSaxAnalyser.readComments` 产出 `CellExtra(COMMENT, text, row, col)`）；图片占位通过注册自定义 `(String, IMAGE)` 读转换器实现（已验证 `StringImageConverter` 仅有写路径，读路径无默认 IMAGE 转换器，`doConvertToJavaObject` 找不到转换器时返回 null，自定义注册即可接管）；CSV 探测新增 `CsvDetector`（BOM → UTF-8 严格校验 → GBK 兜底；分隔符按首行引号外频次嗅探）；空行通过 `ignoreEmptyRow(false)` + 尾部裁剪实现。

**Tech Stack:** easyexcel 4.0.3（既有）、Apache Commons CSV 1.11.0（既有）、POI 5.2.5（仅测试）、JUnit 5 + AssertJ（既有 test 作用域）。

## Global Constraints

- 只改 `easyexcel-markdown/` 目录，**`easyexcel-core` 与上游零差异不得破坏**（验证：`git diff 3afdea9d..HEAD -- easyexcel-core/` 输出 0 行）
- Java 8 语法兼容（父 pom `maven.compiler.source/target=1.8`）
- 不引入新依赖（POI/commons-csv 均为 easyexcel-core 传递依赖，仅测试用）
- 阿里 PMD 规约必须通过（魔法值提常量、`HashMap` 初始化指定容量）
- 单测命令：`cd /Users/wandl/workspaces/workspace-github/easyexcel && ./mvnw -pl easyexcel-markdown test -Dmaven.test.skip=false -Dsurefire.failIfNoSpecifiedTests=false -DargLine="--add-opens=java.base/java.lang=ALL-UNNAMED"`
- 提交信息风格：`feat(markdown): <内容> (Alibaba EasyExcel listener)`
- 渲染格式约定：
  - 超链接：`[单元格文本](URL)`；URL 含空格/圆括号时包裹尖括号 `[文本](<URL>)`；单元格为空时文本取 URL 本身
  - 批注：`值 <!-- 批注内容 -->`（批注内换行折叠为空格）；值为空时仅 `<!-- 批注内容 -->`
  - 图片：常量占位符 `[image]`
  - 空行：中部全空行保留为空表格行，尾部连续全空行裁剪

---

### Task 1: CsvDetector——字符集探测 + 分隔符嗅探

**Files:**
- Create: `easyexcel-markdown/src/main/java/com/alibaba/excel/markdown/CsvDetector.java`
- Modify: `easyexcel-markdown/src/main/java/com/alibaba/excel/markdown/XlsxToMarkdownConverter.java`（`loadCsv(File)`/`loadCsv(Reader,String)` 重写为探测版）
- Test: `easyexcel-markdown/src/test/java/com/alibaba/excel/markdown/CsvDetectorTest.java`

**Interfaces:**
- Consumes: 无（首个任务）
- Produces: `class CsvDetector`，包级私有静态方法：
  - `static Charset detectCharset(byte[] head)`——UTF-8/UTF-16LE/UTF-16BE BOM 检出对应字符集；无 BOM 时 UTF-8 严格解码校验通过返回 UTF-8，抛 `CharacterCodingException` 返回 GBK
  - `static char detectDelimiter(Reader reader) throws IOException`——要求 `reader.markSupported()`（调用方传 `BufferedReader`），读首行统计引号外 `,` `;` `\t` `|` 频次取最大，全零返回 `','`；方法先 `mark` 后 `reset`，不消费流
- Converter 内部新增 `private static SheetTable parseCsv(InputStream in, String sheetName) throws IOException`，供 File 与流式入口共用

- [x] **Step 1: 写失败测试**

```java
package com.alibaba.excel.markdown;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.StringReader;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test CSV charset detection and delimiter sniffing end to end.
 *
 * @author wandl
 */
public class CsvDetectorTest {

    @TempDir
    Path tempDir;

    private File csv(byte[] bytes) throws IOException {
        File file = tempDir.resolve("sniff.csv").toFile();
        Files.write(file.toPath(), bytes);
        return file;
    }

    @Test
    public void decodeGbkCsvWithChineseHeader() throws IOException {
        byte[] gbk = "姓名,分数\n张三,90\n".getBytes(Charset.forName("GBK"));
        SheetDocument document = XlsxToMarkdownConverter.toStructured(csv(gbk));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("姓名", "分数"));
        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("张三", "90"));
    }

    @Test
    public void decodeUtf16LeWithBomAndSemicolon() throws IOException {
        byte[] utf16 = "\uFEFF姓名;分数\n张三;90\n".getBytes(StandardCharsets.UTF_16LE);
        SheetDocument document = XlsxToMarkdownConverter.toStructured(csv(utf16));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("姓名", "分数"));
        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("张三", "90"));
    }

    @Test
    public void sniffSemicolonDelimiter() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(
            csv("a;b;c\n1;2;3\n".getBytes(StandardCharsets.UTF_8)));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("a", "b", "c"));
        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("1", "2", "3"));
    }

    @Test
    public void sniffTabDelimiter() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(
            csv("a\tb\n1\t2\n".getBytes(StandardCharsets.UTF_8)));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("a", "b"));
    }

    @Test
    public void quotedDelimiterDoesNotConfuseSniffing() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(
            csv("name;note\na;\"x;y\"\n".getBytes(StandardCharsets.UTF_8)));

        assertThat(document.sheets.get(0).rows).containsExactly(Arrays.asList("a", "x;y"));
    }

    @Test
    public void utf8StillWorksAndBomStripped() throws IOException {
        SheetDocument document = XlsxToMarkdownConverter.toStructured(
            csv("\uFEFFh1,h2\nv1,v2\n".getBytes(StandardCharsets.UTF_8)));

        assertThat(document.sheets.get(0).headers).containsExactly(Arrays.asList("h1", "h2"));
    }

    @Test
    public void detectDelimiterResetsReader() throws IOException {
        BufferedReader reader = new BufferedReader(new StringReader("a;b\n1;2\n"));
        char delimiter = CsvDetector.detectDelimiter(reader);

        assertThat(delimiter).isEqualTo(';');
        assertThat(reader.readLine()).isEqualTo("a;b");
    }
}
```

- [x] **Step 2: 运行确认失败**

Run: 单测命令 + `-Dtest='CsvDetectorTest'`
Expected: 编译失败 `找不到符号: CsvDetector`；GBK 用例即使借旧实现跑也输出乱码（旧 `loadCsv` 写死 UTF-8）

- [x] **Step 3: 实现 CsvDetector**

```java
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
    private static final char[] DELIMITER_CANDIDATES = {',', ';', '\t', '|'};
    private static final char DEFAULT_DELIMITER = ',';
    private static final char QUOTE = '"';
    private static final char CR = '\r';
    private static final char LF = '\n';

    private CsvDetector() {}

    static Charset detectCharset(byte[] head) {
        if (head.length >= 3 && (head[0] & 0xFF) == 0xEF && (head[1] & 0xFF) == 0xBB
            && (head[2] & 0xFF) == 0xBF) {
            return StandardCharsets.UTF_8;
        }
        if (head.length >= 2 && (head[0] & 0xFF) == 0xFF && (head[1] & 0xFF) == 0xFE) {
            return StandardCharsets.UTF_16LE;
        }
        if (head.length >= 2 && (head[0] & 0xFF) == 0xFE && (head[1] & 0xFF) == 0xFF) {
            return StandardCharsets.UTF_16BE;
        }
        try {
            StandardCharsets.UTF_8.newDecoder().decode(ByteBuffer.wrap(head));
            return StandardCharsets.UTF_8;
        } catch (CharacterCodingException e) {
            return Charset.forName("GBK");
        }
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
```

- [x] **Step 4: 改写 Converter 的 CSV 路径**

删除 `loadCsv(File)`、`loadCsv(Reader, String)`、`fillFromCsv(SheetTable, Reader)` 中的旧固定 UTF-8 逻辑，替换为（`BomStrippingReader` 内部类保留不动）：

```java
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
        BufferedInputStream buffered =
            in instanceof BufferedInputStream ? (BufferedInputStream)in
                : new BufferedInputStream(in, CsvDetector.SNIFF_LIMIT);
        buffered.mark(CsvDetector.SNIFF_LIMIT);
        byte[] head = new byte[CsvDetector.SNIFF_LIMIT];
        int headLength = 0;
        int count;
        while (headLength < head.length && (count = buffered.read(head, headLength, head.length - headLength)) != -1) {
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
```

同步修改：流式入口 `toStructured(InputStream)` 的 CSV 分支改为 `document.sheets.add(loadCsv(pushback, "CSV"))`（原来是 `loadCsv(new BomStrippingReader(...), "CSV")`）；import 增加 `java.io.BufferedInputStream`、`java.io.BufferedReader`。

- [x] **Step 5: 跑测试确认通过（含旧的 CSV 测试不回归）**

Run: 单测命令（全模块）
Expected: `Tests run: 24, Failures: 0`（17 旧 + 7 新）

- [x] **Step 6: Commit**

```bash
git add easyexcel-markdown/
git commit -m "feat(markdown): csv charset detection and delimiter sniffing (Alibaba EasyExcel listener)"
```

---

### Task 2: 图片占位（执行时方案已裁决变更）

> **实际实现与原方案的差异**：原方案"注册 `(String, IMAGE)` 读转换器"在执行时被证伪——全库检索证实 easyexcel 4.0.3 的 SAX 读路径**从不产出** `CellDataTypeEnum.IMAGE` 单元格（IMAGE 仅存在于写路径）。实际方案：新增 `DrawingAnchorScanner`，以 POI `XSSFReader` 流式遍历 sheet 关系找到 drawing 部件，SAX 解析 `twoCellAnchor`/`oneCellAnchor` 的 `from` 坐标与 `pic` 元素，得到每 sheet 的图片锚点集合；转换器在合并填充后把空锚点单元格置为 `[image]`（有值则保留值，行/列缺失时扩展）。`.xls`（非 ZIP 魔数）跳过扫描；`InputStream` 入口暂不做图片扫描。测试 3 项全绿。

**Files:**
- Create: `easyexcel-markdown/src/main/java/com/alibaba/excel/markdown/ImagePlaceholderConverter.java`
- Modify: `easyexcel-markdown/src/main/java/com/alibaba/excel/markdown/XlsxToMarkdownConverter.java`（`loadExcel` 的 builder 链上 `.registerConverter(new ImagePlaceholderConverter())`）
- Test: `easyexcel-markdown/src/test/java/com/alibaba/excel/markdown/ImagePlaceholderTest.java`

**Interfaces:**
- Consumes: easyexcel `Converter<String>` SPI（`supportJavaTypeKey`/`supportExcelTypeKey`/`convertToJavaData(ReadConverterContext)`）
- Produces: `public final class ImagePlaceholderConverter implements Converter<String>`，常量 `static final String PLACEHOLDER = "[image]"`

- [x] **Step 1: 写失败测试**

```java
package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Base64;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test that cells carrying an anchored picture render as a placeholder instead of vanishing.
 *
 * @author wandl
 */
public class ImagePlaceholderTest {

    private static final String PNG_1X1 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

    @TempDir
    File tempDir;

    @Test
    public void anchoredPictureRendersPlaceholder() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Pic");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Logo");
        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("n/a");

        byte[] png = Base64.getDecoder().decode(PNG_1X1);
        int pictureIndex = workbook.addPicture(png, Workbook.PICTURE_TYPE_PNG);
        XSSFDrawing drawing = (XSSFDrawing)sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 1, 1, 2);
        drawing.createPicture(anchor, pictureIndex);

        File file = new File(tempDir, "pic.xlsx");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);

        List<List<String>> rows = document.sheets.get(0).rows;
        assertThat(rows.get(0)).containsExactly("[image]");
    }
}
```

- [x] **Step 2: 运行确认失败**

Run: 单测命令 + `-Dtest='ImagePlaceholderTest'`
Expected: FAIL，实际行为为图片单元格被丢弃（行内容为空或缺失 `[image]`）。
**若失败形态是"整行丢失"**：先跑 `codegraph node "convertReadCellData"` 确认无模型读取对 IMAGE 类型走 `doConvertToJavaObject` 返回 null 的路径后再继续实现（Task 2 前置验证已确认该行为）。

- [x] **Step 3: 实现转换器**

```java
package com.alibaba.excel.markdown;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.converters.ReadConverterContext;
import com.alibaba.excel.enums.CellDataTypeEnum;

/**
 * Markdown tables cannot embed workbook images, so a cell that carries a picture is rendered
 * as a plain {@link #PLACEHOLDER} marker instead of silently disappearing.
 *
 * @author wandl
 */
public final class ImagePlaceholderConverter implements Converter<String> {

    static final String PLACEHOLDER = "[image]";

    @Override
    public Class<?> supportJavaTypeKey() {
        return String.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.IMAGE;
    }

    @Override
    public String convertToJavaData(ReadConverterContext<?> context) {
        return PLACEHOLDER;
    }
}
```

- [x] **Step 4: 注册并验证**

`loadExcel` 的 builder 链（现 `.extraRead(CellExtraTypeEnum.MERGE)` 之后）加一行：

```java
            .registerConverter(new ImagePlaceholderConverter())
```

Run: 单测命令（全模块）
Expected: `Tests run: 25, Failures: 0`

- [x] **Step 5: Commit**

```bash
git add easyexcel-markdown/
git commit -m "feat(markdown): render image cells as placeholder (Alibaba EasyExcel listener)"
```

---

### Task 3: 超链接 + 批注渲染

**Files:**
- Modify: `easyexcel-markdown/src/main/java/com/alibaba/excel/markdown/XlsxToMarkdownConverter.java`
- Test: `easyexcel-markdown/src/test/java/com/alibaba/excel/markdown/HyperlinkAndCommentTest.java`

**Interfaces:**
- Consumes: Task 1/2 后的 `XlsxToMarkdownConverter` 结构；easyexcel `CellExtra`（`getType/getText/getRowIndex/getColumnIndex`）
- Produces: 转换器内部静态方法（供测试断言渲染格式，包级可见）：
  - `static String markdownLink(String text, String url)`
  - `static String appendComment(String value, String comment)`
  - 后处理管道顺序：`normalizeWidth` → `applyMergedRegions` → `applyCellExtras` → `trimTrailingEmptyRows`（Task 4）→ `promoteFirstRowToHeader`
  - `MarkdownReadListener` 新增字段 `Map<Integer, Map<String, String>> hyperlinkMap` 与 `commentMap`（键 = `rowIndex + ":" + columnIndex`）

- [x] **Step 1: 写失败测试**

```java
package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test hyperlink and comment rendering end to end.
 *
 * @author wandl
 */
public class HyperlinkAndCommentTest {

    @TempDir
    File tempDir;

    private Workbook hyperlinkAndCommentWorkbook() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Link");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Site");
        header.createCell(1).setCellValue("Note");
        Row data = sheet.createRow(1);

        CreationHelper helper = workbook.getCreationHelper();
        Hyperlink hyperlink = helper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress("https://github.com/alibaba/easyexcel");
        Cell linkCell = data.createCell(0);
        linkCell.setCellValue("easyexcel");
        linkCell.setHyperlink(hyperlink);

        Cell commentCell = data.createCell(1);
        commentCell.setCellValue("v4.0.3");
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(1);
        anchor.setCol2(2);
        anchor.setRow1(1);
        anchor.setRow2(2);
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(helper.createRichTextString("line1\nline2"));
        commentCell.setCellComment(comment);
        return workbook;
    }

    private SheetDocument convert(Workbook workbook) throws IOException {
        File file = new File(tempDir, "link.xlsx");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        return XlsxToMarkdownConverter.toStructured(file);
    }

    @Test
    public void renderHyperlinkAsMarkdownLink() throws IOException {
        SheetDocument document = convert(hyperlinkAndCommentWorkbook());

        List<String> row = document.sheets.get(0).rows.get(0);
        assertThat(row.get(0)).isEqualTo("[easyexcel](https://github.com/alibaba/easyexcel)");
    }

    @Test
    public void renderCommentWithFoldedNewlines() throws IOException {
        SheetDocument document = convert(hyperlinkAndCommentWorkbook());

        List<String> row = document.sheets.get(0).rows.get(0);
        assertThat(row.get(1)).isEqualTo("v4.0.3 <!-- line1 line2 -->");
    }

    @Test
    public void markdownContainsRenderedExtras() throws IOException {
        assertThat(XlsxToMarkdownConverter.toMarkdown(convert(hyperlinkAndCommentWorkbook())
            .sheets.isEmpty() ? null : hyperlinkAndCommentWorkbook() == null ? null : new File(tempDir, "link.xlsx")))
            .contains("[easyexcel](https://github.com/alibaba/easyexcel)")
            .contains("<!-- line1 line2 -->");
    }

    @Test
    public void urlWithSpacesIsAngleWrapped() {
        assertThat(XlsxToMarkdownConverter.markdownLink("doc", "https://example.com/a b"))
            .isEqualTo("[doc](<https://example.com/a b>)");
        assertThat(XlsxToMarkdownConverter.markdownLink("", "https://example.com"))
            .isEqualTo("[https://example.com](https://example.com)");
    }
}
```

注意：`markdownContainsRenderedExtras` 中间那段三元表达式仅为满足编译的写法是错误的——直接改为：先 `convert(hyperlinkAndCommentWorkbook())` 生成文件后再断言文件路径。实现时用以下修正版：

```java
    @Test
    public void markdownContainsRenderedExtras() throws IOException {
        Workbook workbook = hyperlinkAndCommentWorkbook();
        File file = new File(tempDir, "link.xlsx");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
        assertThat(XlsxToMarkdownConverter.toMarkdown(file))
            .contains("[easyexcel](https://github.com/alibaba/easyexcel)")
            .contains("<!-- line1 line2 -->");
    }
```

- [x] **Step 2: 运行确认失败**

Run: 单测命令 + `-Dtest='HyperlinkAndCommentTest'`
Expected: 编译失败（`markdownLink` 不存在）或断言失败（超链接当前只保留文本 `easyexcel`、批注丢失）

- [x] **Step 3: 实现渲染与收集**

3a. `loadExcel` builder 链补两个 extraRead（叠加语义已验证）：

```java
        try (ExcelReader reader = builder.headRowNumber(0)
            .extraRead(CellExtraTypeEnum.MERGE)
            .extraRead(CellExtraTypeEnum.HYPERLINK)
            .extraRead(CellExtraTypeEnum.COMMENT)
            .registerConverter(new ImagePlaceholderConverter())
            .registerReadListener(new MarkdownReadListener(sheetMap, mergeMap, hyperlinkMap, commentMap))
            .build()) {
```

3b. `MarkdownReadListener` 增加两个字段与 `extra()` 分支：

```java
        private final Map<Integer, Map<String, String>> hyperlinkMap;
        private final Map<String, String> NO_MAP = null; // 仅示意，勿加

        @Override
        public void extra(CellExtra extra, AnalysisContext context) {
            Integer sheetNo = context.readSheetHolder().getSheetNo();
            if (extra.getType() == CellExtraTypeEnum.MERGE) {
                ...（原逻辑不变）
                return;
            }
            if (extra.getType() == CellExtraTypeEnum.HYPERLINK) {
                putByCell(hyperlinkMap, sheetNo, extra.getRowIndex(), extra.getColumnIndex(),
                    extra.getText());
                return;
            }
            if (extra.getType() == CellExtraTypeEnum.COMMENT) {
                putByCell(commentMap, sheetNo, extra.getRowIndex(), extra.getColumnIndex(),
                    extra.getText());
            }
        }

        private static void putByCell(Map<Integer, Map<String, String>> map, Integer sheetNo,
            Integer rowIndex, Integer columnIndex, String value) {
            Map<String, String> cells = map.get(sheetNo);
            if (cells == null) {
                cells = new HashMap<>(16);
                map.put(sheetNo, cells);
            }
            cells.put(rowIndex + ":" + columnIndex, value);
        }
```

（实现时删除 `NO_MAP` 示意行，字段声明为 `private final Map<Integer, Map<String, String>> hyperlinkMap;` 与 `... commentMap;`，构造器参数对应追加，`loadExcel` 里初始化为 `new HashMap<>(16)`。）

3c. 后处理管道与渲染方法（`loadExcel` 循环改为）：

```java
        List<SheetTable> tables = new ArrayList<>();
        for (Map.Entry<Integer, SheetTable> entry : sheetMap.entrySet()) {
            SheetTable table = entry.getValue();
            normalizeWidth(table);
            applyMergedRegions(table, mergeMap.get(entry.getKey()));
            applyCellExtras(table, hyperlinkMap.get(entry.getKey()), commentMap.get(entry.getKey()));
            promoteFirstRowToHeader(table);
            tables.add(table);
        }
```

`applyMergedRegions` 删除内部的补宽循环（职责移入 `normalizeWidth`）。新增方法：

```java
    private static void normalizeWidth(SheetTable table) {
        int width = 0;
        for (List<String> row : table.rows) {
            width = Math.max(width, row.size());
        }
        if (width == 0) {
            return;
        }
        for (List<String> row : table.rows) {
            for (int columnIndex = row.size(); columnIndex < width; columnIndex++) {
                row.add("");
            }
        }
    }

    private static void applyCellExtras(SheetTable table, Map<String, String> hyperlinks,
        Map<String, String> comments) {
        if ((hyperlinks == null || hyperlinks.isEmpty()) && (comments == null || comments.isEmpty())) {
            return;
        }
        for (int rowIndex = 0; rowIndex < table.rows.size(); rowIndex++) {
            List<String> row = table.rows.get(rowIndex);
            for (int columnIndex = 0; columnIndex < row.size(); columnIndex++) {
                if (hyperlinks != null) {
                    String url = hyperlinks.get(rowIndex + ":" + columnIndex);
                    if (url != null) {
                        row.set(columnIndex, markdownLink(row.get(columnIndex), url));
                    }
                }
                if (comments != null) {
                    String comment = comments.get(rowIndex + ":" + columnIndex);
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
        String folded = comment.replace("\r\n", " ").replace('\r', ' ').replace('\n', ' ").trim();
        String cell = value == null ? "" : value;
        return cell.isEmpty() ? "<!-- " + folded + " -->" : cell + " <!-- " + folded + " -->";
    }
```

注意 `appendComment` 第三行 `replace('\n', ' ")` 是笔误，实现时为 `replace('\n', ' ')`。`trim()` 后若 `folded` 为空则只加 `<!--  -->` 也接受（不额外分支，YAGNI）。

- [x] **Step 4: 跑测试确认通过**

Run: 单测命令（全模块）
Expected: `Tests run: 29, Failures: 0`（25 + 4 新）

- [x] **Step 5: Commit**

```bash
git add easyexcel-markdown/
git commit -m "feat(markdown): render hyperlinks and cell comments (Alibaba EasyExcel listener)"
```

---

### Task 4: 中部空行保留 + 尾部空行裁剪

**Files:**
- Modify: `easyexcel-markdown/src/main/java/com/alibaba/excel/markdown/XlsxToMarkdownConverter.java`
- Test: `easyexcel-markdown/src/test/java/com/alibaba/excel/markdown/EmptyRowTest.java`

**Interfaces:**
- Consumes: Task 3 管道（`applyCellExtras` 之后）
- Produces: `private static void trimTrailingEmptyRows(SheetTable table)`，管道位置 `applyCellExtras` → `trimTrailingEmptyRows` → `promoteFirstRowToHeader`

- [x] **Step 1: 写失败测试**

```java
package com.alibaba.excel.markdown;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test that a blank row in the middle of the data is kept while trailing blank rows are cut.
 *
 * @author wandl
 */
public class EmptyRowTest {

    @TempDir
    File tempDir;

    @Test
    public void keepMidTableBlankRowAndTrimTrailing() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Blank");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("A");
        header.createCell(1).setCellValue("B");
        Row first = sheet.createRow(1);
        first.createCell(0).setCellValue("1");
        first.createCell(1).setCellValue("2");
        // row index 2 exists in the XML but has no cell at all: the mid-table blank row
        sheet.createRow(2);
        Row third = sheet.createRow(3);
        third.createCell(0).setCellValue("3");
        third.createCell(1).setCellValue("4");
        // trailing styled-empty rows must be trimmed
        sheet.createRow(4);
        sheet.createRow(5);

        File file = new File(tempDir, "blank.xlsx");
        try (OutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }

        SheetDocument document = XlsxToMarkdownConverter.toStructured(file);

        List<List<String>> rows = document.sheets.get(0).rows;
        assertThat(rows).containsExactly(
            Arrays.asList("1", "2"),
            Arrays.asList("", ""),
            Arrays.asList("3", "4"));

        assertThat(document.toMarkdown()).contains("|  |  |");
    }
}
```

- [x] **Step 2: 运行确认失败**

Run: 单测命令 + `-Dtest='EmptyRowTest'`
Expected: FAIL——当前 `ignoreEmptyRow` 默认 true，中部空行被 easyexcel 吞掉，实际 rows 只有 `["1","2"]` 与 `["3","4"]`

- [x] **Step 3: 实现**

3a. `loadExcel` builder 链加 `.ignoreEmptyRow(false)`（与 `headRowNumber(0)` 同链）。

3b. `MarkdownReadListener.invoke` 对空 map 生成的 `width=0` 行改为至少占 1 列？不——保持 `width=0` 生成空 `row`，由 `normalizeWidth` 统一补齐（现状即如此，无需改）。

3c. 新增裁剪方法并接入管道：

```java
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
```

管道最终形态（Task 3 的循环内）：

```java
            normalizeWidth(table);
            applyMergedRegions(table, mergeMap.get(entry.getKey()));
            applyCellExtras(table, hyperlinkMap.get(entry.getKey()), commentMap.get(entry.getKey()));
            trimTrailingEmptyRows(table);
            promoteFirstRowToHeader(table);
```

注意：若裁剪后 `rows` 全空且 `headers` 将提升的是唯一空行，`promoteFirstRowToHeader` 会把首行（可能为空行）提为表头——该行为等价于"(empty sheet)"路径，可接受；若整表为空行，提升后 `SheetTable.toMarkdown()` 输出空表头表体，由既有 `columnCount` 逻辑兜底渲染 `| |`，不崩溃。

- [x] **Step 4: 跑测试确认通过（含既有全部测试，验证 ignoreEmptyRow(false) 无回归）**

Run: 单测命令（全模块）
Expected: `Tests run: 30, Failures: 0`。
**若既有测试回归**（例如 `convertMultiSheetXlsx` 因 `Arrays.asList("only-a", null, "c")` 的 null 单元格行变为显式空行）——该行为正是本任务目标，更新对应断言而非回退开关。

- [x] **Step 5: Commit**

```bash
git add easyexcel-markdown/
git commit -m "feat(markdown): keep mid-table blank rows and trim trailing blanks (Alibaba EasyExcel listener)"
```

---

### Task 5: 全量回归 + 收尾

**Files:**
- Modify: 本计划文件的勾选框

- [x] **Step 1: 全量回归（clean，与 CI 同款）**

```bash
cd /Users/wandl/workspaces/workspace-github/easyexcel && ./mvnw clean test -Dmaven.test.skip=false \
  -DargLine="--add-opens=java.base/java.lang=ALL-UNNAMED --add-opens=java.base/java.util=ALL-UNNAMED --add-opens=java.base/sun.reflect.annotation=ALL-UNNAMED"
```

Expected: 6 模块全 SUCCESS；`easyexcel-markdown` 30 个测试全绿；`easyexcel-test` 161 全绿；退出码 0

- [x] **Step 2: 核验 core 零差异**

```bash
git diff 3afdea9d..HEAD -- easyexcel-core/ | wc -l   # 必须为 0
```

- [x] **Step 3: 更新本计划勾选框并提交**

```bash
git add docs/superpowers/plans/2026-08-27-easyexcel-markdown-perfect-conversion.md
git commit -m "docs(markdown): check off perfect conversion plan (Alibaba EasyExcel listener)"
```

## Self-Review

- **覆盖核对**：缺口 1 超链接→Task 3；缺口 2 批注→Task 3；缺口 3 图片→Task 2；缺口 4 CSV 字符集→Task 1；缺口 5 CSV 分隔符→Task 1；缺口 8 空行→Task 4。缺口 6/7（样式/富文本）明确不在范围（Markdown 格式天限）。✓
- **占位符扫描**：Task 3 Step 1 的一个测试初稿含错误三元表达式与 Step 3 的两处笔误均已在文中就地标注修正版，执行时以修正版为准。✓
- **类型一致性**：`CsvDetector.detectCharset(byte[])→Charset`、`detectDelimiter(Reader)→char`、`markdownLink(String,String)→String`、`appendComment(String,String)→String`、`PLACEHOLDER` 常量在各任务间引用一致；`MarkdownReadListener` 构造器在 Task 3 扩为四参数。✓
- **风险点**：Task 2 依赖 easyexcel 读路径对锚定图片产出 `ReadCellData(IMAGE)`（已从 `StringImageConverter` 仅写侧注册 + `doConvertToJavaObject` null 返回路径推断可行，测试为最终裁决）；Task 4 的 `ignoreEmptyRow(false)` 可能改变既有测试对空白行的断言，处理策略已写明。
