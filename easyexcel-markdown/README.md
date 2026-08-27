# easyexcel-markdown

基于 Alibaba EasyExcel 流式读取的 XLSX / XLS / CSV 到多 Sheet Markdown 转换器。
支持 100MB+ 大文件，逐 Sheet 处理不全量加载。

## 快速开始

Maven 依赖（版本随仓库 `${revision}` 统一发布，无需手动指定）：

```xml
<dependency>
    <groupId>com.alibaba</groupId>
    <artifactId>easyexcel-markdown</artifactId>
    <version>${revision}</version>
</dependency>
```

### 用法一：全量 String API

```java
// 返回完整 Markdown 字符串（所有 Sheet 拼接）
String md = XlsxToMarkdownConverter.toMarkdown(new File("report.xlsx"));

// 返回结构化对象，可按 Sheet 遍历
SheetDocument doc = XlsxToMarkdownConverter.toStructured(new File("report.xlsx"));
for (SheetTable sheet : doc.sheets) {
    System.out.println(sheet.sheetName + ": " + sheet.rows.size() + " rows");
}
```

### 用法二：流式 Writer API

单 Sheet 内存占用有界，适合写入文件或 HTTP 响应流。CSV 走两遍法逐行写入。

```java
try (Writer writer = new BufferedWriter(new FileWriter("output.md"))) {
    XlsxToMarkdownConverter.toMarkdown(new File("huge.xlsx"), writer);
}
```

### 用法三：多行分组表头

```java
// 前 2 行作为表头，空格拼接合并（如 "基本信息 姓名"）
SheetDocument doc = XlsxToMarkdownConverter.toStructured(new File("grouped.xlsx"), 2);
```

## 能力矩阵

| 能力 | 状态 | 说明 |
|------|------|------|
| 多 Sheet 保序与空 Sheet 占位 | 已支持 | TreeMap 按 sheetNo 排序，空 Sheet 输出 `(empty sheet)` |
| 合并单元格左上值回填 | 已支持 | 通过 EasyExcel MERGE extra 回填 |
| 超链接 `[文本](URL)` | 已支持 | 通过 HYPERLINK extra 渲染为 Markdown 链接 |
| 批注 `值 <!-- 注 -->` | 已支持 | 通过 COMMENT extra 附加 HTML 注释 |
| 图片占位 `[image]` | 已支持 | XLSX/XLS DrawingAnchorScanner 扫描锚点；流式入口经临时文件 spool |
| 行内样式 `**粗**/*斜*/~~删~~` | 已支持 | XLSX SAX 解析 styles.xml；XLS 通过 HSSFWorkbook 读取字体 |
| Rich text run 级样式 | 已支持 | XLSX sharedStrings.xml / XLS HSSFRichTextString 逐段包裹 |
| 列对齐 `:---:` / `---:` | 已支持 | 从表头单元格 horizontal alignment 推导 |
| CSV 字符集探测 | 已支持 | BOM 优先 -> UTF-8 严格校验 -> GBK 兜底（兼容中文 Windows Excel） |
| CSV 分隔符嗅探 | 已支持 | 首行引号外 `,;|\t` 频次取最高 |
| 格式保真 | 已支持 | 日期/百分比/千分位按 Excel 显示值输出（EasyExcel 层面） |
| 空行策略 | 已支持 | 中部保留、尾部裁剪 |
| 隐藏行列 | 原样输出 | EasyExcel SAX 路径不读取 hidden 属性，隐藏行列照常渲染 |

## 已知边界

- **颜色 / 背景 / 边框 / 字号**：GFM 表格无对应语法，有意丢弃（见 `CellStyle` javadoc）。
- **XLS 样式**：`CellStyleScanner.scanLegacyStyles()` 加载整个 `HSSFWorkbook` 枚举单元格，支持 bold / italic / strikeout / 对齐 / rich text run；因 XLS 上限 65536 行 x 256 列，全量加载可接受。
- **XLS 图片**：通过 `HSSFPatriarch` 枚举 `HSSFPicture` + `HSSFClientAnchor`，与 XLSX SAX 路径对齐。

## 性能参考

手动基准命令（需 JDK 17+，默认跳过）：

```bash
./mvnw -pl easyexcel-markdown test -Dmaven.test.skip=false \
  -Dtest='MarkdownBenchmarkTest' -Dbenchmark=true \
  -DargLine="--add-opens=java.base/java.lang=ALL-UNNAMED"
```

最近一次参考数字（单机非 JMH，仅量级参考）：

| 行数 | struct 耗时 (中位) | struct 堆增量 | Markdown 输出大小 |
|------|-------------------|--------------|------------------|
| 10,000 | ~0.4s | ~15 MB | ~1.5 MB |
| 50,000 | ~1.8s | ~50 MB | ~7.5 MB |
| 200,000 | ~4.6s | ~102 MB | ~30 MB |

> 测试环境：Apple Silicon, JDK 17, 8 列 x N 行 XLSX。数字仅供量级参考，不代表生产环境表现。

## 许可证

Apache License 2.0，与 easyexcel 主仓库一致。
