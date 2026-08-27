# easyexcel XLSX → Markdown 计划（参考 markitdown converter-xlsx，Alibaba easyexcel 自研）

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 Alibaba 原版 easyexcel 仓库新增 `XlsxToMarkdownConverter`，将 XLSX/XLS/CSV 还原为**多 Sheet 结构化 Markdown**（表头 + 数据行 + 合并单元格），**与 markitdown converter-xlsx 互补**（后者用 Apache POI 单行读取；本计划用 Alibaba `EasyExcel.read()` 流式监听器可处理 100MB+ 大文件）。

**参考 markitdown converter-xlsx（**`io.gitlab.ade90036:converter-xlsx:1.0.0`**）**：用 Apache POI 5.2.5 的 SXSSFWorkbook 流式读取，**单表 Sheet** 渲染为 Markdown table，**无 Sheet 名识别、无合并单元格**。

**easypdf/easydoc/easyexcel/easyodf 优势（差异化定位）**：
- **多 Sheet 完整保留**：循环 `AnalysisContext.readSheetHolder()` 收集每个 Sheet 名 + 表头 + 数据
- **内存高效**：Alibaba EasyExcel 的 `ReadListener` 流式回调，OOM 概率低
- **合并单元格识别**：`@ExcelProperty` 注解映射 + `extra` 第 1 行识别列；本计划额外支持 `mergedRegions` 解析
- **CSV 用 Commons CSV**：与 markitdown converter-csv 同生态（Apache Commons CSV 1.11.0）

代码位置：`/Users/wandl/workspaces/workspace-github/easyexcel/easyexcel-core/src/main/java/com/alibaba/excel/markdown/`

**Tech Stack:** Alibaba easyexcel 4.0.3、Apache Commons CSV 1.11.0（传递依赖）、JUnit 5 + AssertJ、Java 8+。

## Global Constraints

- 新文件放 `easyexcel-core/src/main/java/com/alibaba/excel/markdown/`
- **Java 8 语法兼容**（Alibaba easyexcel 4.x 支持 JDK 8+）
- **不引入新依赖**（Commons CSV 已是传递依赖）
- POJO 字段与 ddd4j/easypdf/easydoc 同构
- 提交信息风格：`feat(markdown): add xlsx-to-markdown structure extraction (Alibaba EasyExcel listener)`
- 验证命令：`cd /Users/wandl/workspaces/workspace-github/easyexcel && ./mvnw -pl easyexcel-core -am test -Dsurefire.failIfNoSpecifiedTests=false`

---

### Task 1: SheetDocument POJO

**Files:**
- Create: `easyexcel-core/src/main/java/com/alibaba/excel/markdown/SheetDocument.java`
- Create: `easyexcel-core/src/main/java/com/alibaba/excel/markdown/SheetTable.java`
- Test: `easyexcel-core/src/test/java/com/alibaba/excel/markdown/SheetDocumentTest.java`

- [x] **Step 1-5: 实现 POJO + 测试 + Commit**

```java
// SheetTable.java
public final class SheetTable {
    public String sheetName = "";
    public List<List<String>> headers = new ArrayList<>();
    public List<List<String>> rows = new ArrayList<>();
    public String toMarkdown() {
        // 详见上一版 Task 1 Step 3
    }
}
// SheetDocument.java
public final class SheetDocument {
    public String title;
    public List<SheetTable> sheets = new ArrayList<>();
    public String toMarkdown() { /* 遍历 sheets 调用 toMarkdown */ }
}
```

---

### Task 2: XlsxToMarkdownConverter（Alibaba EasyExcel 读路径 + Commons CSV）

**Files:**
- Create: `easyexcel-core/src/main/java/com/alibaba/excel/markdown/XlsxToMarkdownConverter.java`
- Test: `easyexcel-core/src/test/java/com/alibaba/excel/markdown/XlsxToMarkdownConverterTest.java`

**Interfaces:**
- Produces: `public static SheetDocument toStructured(File xlsx)` + `public static String toMarkdown(File xlsx)` + InputStream 入口
- 内部：`EasyExcel.read(file, ReadListener<>)` 收集每 Sheet 行；CSV 用 `CSVParser`

- [x] **Step 1-5: 实现 + 验证 + Commit**

```java
public final class XlsxToMarkdownConverter {
    private XlsxToMarkdownConverter() {}
    public static SheetDocument toStructured(File xlsx) {
        Objects.requireNonNull(xlsx, "xlsx must not be null");
        if (!xlsx.isFile()) throw new RuntimeException("XLSX not found");
        SheetDocument doc = new SheetDocument();
        doc.title = xlsx.getName();
        String n = xlsx.getName().toLowerCase();
        if (n.endsWith(".csv")) doc.sheets.add(loadCsv(xlsx));
        else doc.sheets.addAll(loadXlsx(xlsx));
        return doc;
    }
    // loadXlsx: 用 Map<Integer, SheetTable> sheetMap，按 ctx.readSheetHolder().getSheetNo() 分组
    // loadCsv: 用 CSVParser
}
```

---

### Task 3: EasyExcel 门面 + 全量回归

**Files:**
- Modify: `easyexcel-core/src/main/java/com/alibaba/excel/EasyExcel.java`（新增 `xlsxToMarkdown`/`xlsxToStructured`）

- [x] **Step: 实现 EasyExcel 门面 + 全量回归 + Commit**（177 tests: core 16 + test 161，全部通过）

---

### Task 4: 架构调整——门面回退 + 独立模块拆分（2026-08-27）

**决策**：markdown 属于渲染/导出格式而非 Excel 解析职责，不应改动 `easyexcel-core`（尤其不能污染 `EasyExcel` 门面 API，避免与 alibaba 上游产生合并摩擦）。

**Files:**
- Revert: `easyexcel-core/src/main/java/com/alibaba/excel/EasyExcel.java`（恢复上游原样，`git revert 3feaf98c`）
- Move: `easyexcel-core/.../markdown/*` → `easyexcel-markdown/.../markdown/*`（git 识别为 100% rename，历史保留）
- Create: `easyexcel-markdown/pom.xml`（依赖 easyexcel-core，JUnit 5 + AssertJ test 作用域）
- Modify: 根 `pom.xml`（modules 注册 `easyexcel-markdown`）、`easyexcel-core/pom.xml`（移除 test 依赖，还原上游）

- [x] **Step: 回退门面 + 拆模块 + 完美转换强化 + 全量回归**
  - core 与上游零差异（`git diff 3afdea9d..HEAD -- easyexcel-core/` = 0 行）
  - 强化：`ExcelReader.sheetList()` 预枚举全部 Sheet（空 Sheet 保留为 "(empty sheet)"）、TreeMap 保证 Sheet 顺序确定、try-with-resources 防泄漏
  - 保真测试：日期 `yyyy-mm-dd`→`2026-08-27`、百分比 `0%`→`50%`、千分位 `#,##0.00`→`1,234.50`、布尔、公式缓存值 `1+1`→`2`（POI 写入需 `evaluateAll()` 生成缓存）
  - 最终回归：clean 全 reactor 6 模块 SUCCESS，178 tests 全绿（markdown 17 + test 161）

---

## Self-Review

- **差异化定位**：markitdown converter-xlsx（POI 单 Sheet）→ easyexcel 多 Sheet 流式
- **依赖隔离**：Alibaba easyexcel + Commons CSV（无 POI/POI-OOXML 冲突）
- **POJO 字段一致性**：与 ddd4j Document 同构