# Excel数据驱动Word文档自动生成工具 - 技术实现说明

## 目录

1. [架构设计](#架构设计)
2. [核心模块](#核心模块)
3. [关键技术实现](#关键技术实现)
4. [数据流程](#数据流程)
5. [代码结构](#代码结构)
6. [扩展开发指南](#扩展开发指南)

---

## 架构设计

### 整体架构

```
┌─────────────────┐
│  ExcelToWordTool │  主程序入口
│  (命令行解析)    │
└────────┬────────┘
         │
         ├─────────────────┐
         │                 │
┌────────▼────────┐  ┌─────▼────────┐
│  ExcelReader    │  │ WordProcessor │
│  (数据读取)     │  │ (文档处理)    │
└────────┬────────┘  └─────┬────────┘
         │                 │
         │                 │
┌────────▼────────┐  ┌─────▼────────┐
│  ModuleData     │  │  XWPFDocument │
│  TestCase       │  │  (Apache POI) │
└─────────────────┘  └──────────────┘
```

### 设计原则

1. **单一职责原则**：每个类只负责一个功能模块
2. **开闭原则**：对扩展开放，对修改关闭
3. **依赖倒置**：依赖抽象而非具体实现
4. **格式自适应**：自动提取并应用模板格式

---

## 核心模块

### 1. ExcelReader（Excel数据读取模块）

**职责**：读取Excel文件，解析测试用例数据，按模块分组

**关键方法**：
- `readExcel(String excelPath)`: 读取Excel文件并返回模块数据Map
- `parseRow(Row row, Map<String, Integer> columnIndexMap)`: 解析单行数据

**技术要点**：
- 使用Apache POI XSSFWorkbook读取.xlsx文件
- 支持动态列名识别（不依赖列顺序）
- 自动过滤空行
- 支持多行文本（换行符处理）

**数据模型**：
```java
ModuleData {
    String moduleNumber;      // 模块编号，如 "5.2"
    List<TestCase> testCases; // 测试用例列表
}

TestCase {
    String id;                // 测试用例ID
    String testName;          // 测试名称
    Map<String, String> columnData; // 列数据（动态列）
}
```

### 2. WordProcessor（Word文档处理模块）

**职责**：处理Word模板，定位章节，插入表格和子标题

**关键方法**：
- `processWord(String templatePath, String outputPath, Map<String, ModuleData> moduleDataMap)`: 主处理流程
- `findSectionParagraph(XWPFDocument document, String sectionNumber)`: 查找章节段落
- `createSubSectionParagraph(...)`: 创建子章节段落
- `fillTableData(XWPFTable table, TestCase testCase)`: 填充表格数据
- `extractSubSectionFormat(XWPFParagraph templatePara)`: 提取子章节格式
- `extractCaptionFormat(XWPFParagraph templateCaption)`: 提取Caption格式

**技术要点**：
- 使用Apache POI XWPFDocument处理.docx文件
- 通过样式ID识别章节（Heading 2样式ID=3）
- 通过样式ID识别子章节（Heading 3样式ID=4）
- 通过样式ID识别Caption（Caption样式ID=11）
- 自动跳过目录（TOC）项（样式ID 22, 25, 16）
- 格式自动提取和应用

**格式提取机制**：
1. 从模板第一个子章节提取格式
2. 识别编号部分和内容部分的格式差异
3. 缓存格式信息供后续使用
4. 应用到所有新生成的子章节和Caption

### 3. ConfigLoader（配置加载模块）

**职责**：加载配置文件或解析命令行参数

**关键方法**：
- `loadFromFile()`: 从config.properties加载配置
- `parseArguments(String[] args)`: 解析命令行参数

**技术要点**：
- 使用Apache Commons CLI解析命令行参数
- 支持配置文件方式运行
- 参数验证和错误提示

### 4. FileUtil（文件工具类）

**职责**：文件路径处理和输出文件名生成

**关键方法**：
- `generateOutputFileName(String templatePath, String outputDir)`: 生成输出文件名

**技术要点**：
- 自动创建输出目录
- 生成带时间戳的输出文件名
- 处理相对路径和绝对路径

---

## 关键技术实现

### 1. Word文档结构解析

#### 章节定位策略

**问题**：Word模板中章节标题可能没有编号，只有名称（如"功能测试"），而目录（TOC）中有完整编号（如"5.2 功能测试"）

**解决方案**：双查找策略
1. 从TOC中提取章节编号和名称的映射
2. 在正文中通过名称+样式定位章节段落

```java
// 1. 从TOC提取章节信息
String chapterNameFromToc = extractChapterNameFromToc(document, sectionNumber);

// 2. 在正文中查找（通过名称或编号）
XWPFParagraph sectionPara = findSectionParagraph(document, sectionNumber, chapterNameFromToc);
```

#### 子章节插入位置

**问题**：需要准确找到章节的结束位置，以便插入新的子章节

**解决方案**：`findLastElementInSection`方法
1. 从章节段落开始遍历
2. 遇到下一个同级或上级章节时停止
3. 返回最后一个相关元素（段落或表格）

```java
private XWPFParagraph findLastElementInSection(
    XWPFDocument document, 
    XWPFParagraph sectionPara, 
    String sectionNumber
) {
    // 遍历段落，找到章节的最后一个元素
    // 停止条件：遇到下一个主章节或不同父级的子章节
}
```

### 2. 格式自动提取和应用

#### 格式提取

**子章节格式提取**：
```java
private SubSectionFormat extractSubSectionFormat(XWPFParagraph templatePara) {
    // 1. 获取样式ID
    String styleId = templatePara.getStyle();
    
    // 2. 分析Run的格式
    // 对于Heading 3样式，默认格式：
    // - 编号：黑体 10号
    // - 内容：黑体 12pt
    // 如果Run中有格式，则提取；否则使用样式默认值
}
```

**Caption格式提取**：
```java
private CaptionFormat extractCaptionFormat(XWPFParagraph templateCaption) {
    // 对于Caption样式，默认格式：黑体 12pt
    // 如果Run中有格式，则提取；否则使用样式默认值
}
```

#### 格式应用

**创建子章节时应用格式**：
```java
private XWPFParagraph createSubSectionParagraph(...) {
    // 1. 使用模板格式或默认格式
    SubSectionFormat subFmt = templateSubSectionFormat != null 
        ? templateSubSectionFormat 
        : new SubSectionFormat();
    
    // 2. 设置样式
    para.setStyle(subFmt.styleId);
    
    // 3. 禁用自动编号（避免双编号）
    numId.setVal(BigInteger.ZERO);
    
    // 4. 创建编号Run（使用numberFormat）
    XWPFRun numRun = para.createRun();
    numRun.setFontFamily(subFmt.numberFormat.fontFamily);
    numRun.setFontSize(subFmt.numberFormat.fontSize);
    
    // 5. 创建内容Run（使用contentFormat）
    XWPFRun contentRun = para.createRun();
    contentRun.setFontFamily(subFmt.contentFormat.fontFamily);
    contentRun.setFontSize(subFmt.contentFormat.fontSize);
}
```

### 3. 表格处理

#### 表格模板复制

**问题**：需要复制模板表格结构，但清空数据列

**解决方案**：
1. 使用`copyTable`方法复制CTTbl对象
2. 遍历表格行，清空数据列（保留标签列）
3. 根据表格列数判断布局（2列或4列）

```java
private CTTbl copyTable(XWPFDocument document, CTP afterPara, CTTbl sourceTable) {
    // 1. 克隆CTTbl对象
    CTTbl newTable = (CTTbl) sourceTable.copy();
    
    // 2. 插入到指定位置
    CTBody body = document.getDocument().getBody();
    int insertIndex = findInsertIndex(body, afterPara);
    body.insertTbl(insertIndex, newTable);
    
    return newTable;
}
```

#### 表格数据填充

**问题**：不同表格布局（2列或4列）需要不同的填充策略

**解决方案**：
```java
private void fillTableData(XWPFTable table, TestCase testCase) {
    int columnCount = table.getRow(0).getTableCells().size();
    
    if (columnCount == 2) {
        // 2列布局：第一列标签，第二列数据
        fillTwoColumnTable(table, testCase);
    } else if (columnCount == 4) {
        // 4列布局：第一列标签，第二列数据，第三列标签，第四列数据
        fillFourColumnTable(table, testCase);
    }
}
```

### 4. 自动编号处理

**问题**：Word的Heading样式可能包含自动编号，导致双编号（如"5.2.2 5.2.1 登录功能测试"）

**解决方案**：显式禁用自动编号
```java
// 在段落属性中设置numId=0
CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
CTNumPr numPr = ppr.isSetNumPr() ? ppr.getNumPr() : ppr.addNewNumPr();
CTDecimalNumber numId = numPr.isSetNumId() ? numPr.getNumId() : numPr.addNewNumId();
numId.setVal(BigInteger.ZERO);  // numId=0 表示不使用编号
```

### 5. TOC（目录）识别和跳过

**问题**：目录中的章节标题可能被误识别为正文章节

**解决方案**：通过样式ID识别TOC项
```java
// TOC样式ID: 22=toc1, 25=toc2, 16=toc3
boolean isToc = styleName != null && 
    (styleName.equals("22") || styleName.equals("25") || styleName.equals("16") ||
     styleName.toLowerCase().startsWith("toc"));
```

---

## 数据流程

### 完整处理流程

```
1. 命令行参数解析
   ↓
2. 文件路径验证
   ↓
3. Excel数据读取
   ├─ 读取Excel文件
   ├─ 解析列名映射
   ├─ 按模块分组
   └─ 构建ModuleData对象
   ↓
4. Word模板处理
   ├─ 打开Word文档
   ├─ 扫描章节编号（从TOC）
   ├─ 提取模板格式（从第一个子章节）
   ├─ 遍历模块数据
   │   ├─ 查找章节段落
   │   ├─ 查找已存在的子章节
   │   ├─ 替换或创建子章节
   │   ├─ 创建或更新Caption
   │   ├─ 复制表格模板
   │   └─ 填充表格数据
   └─ 保存输出文件
   ↓
5. 输出结果
```

### 子章节处理流程

```
对于每个测试用例：
1. 检查是否有已存在的子章节模板
   ├─ 有：替换标题和填充表格
   └─ 无：创建新子章节
2. 创建或更新Caption
3. 复制表格模板（如果不存在）
4. 填充表格数据
5. 更新插入点位置
```

---

## 代码结构

### 包结构

```
pub.developers.docautogenbyexcel/
├── ExcelToWordTool.java          # 主程序入口
├── config/
│   └── ConfigLoader.java         # 配置加载器
├── model/
│   ├── ModuleData.java           # 模块数据模型
│   └── TestCase.java             # 测试用例模型
├── reader/
│   └── ExcelReader.java          # Excel读取器
├── processor/
│   └── WordProcessor.java        # Word处理器（核心模块）
└── util/
    └── FileUtil.java             # 文件工具类
```

### 核心类说明

#### WordProcessor

**主要内部类**：
- `RunFormat`: 保存Run的格式信息（字体、字号、加粗）
- `SubSectionFormat`: 保存子章节格式（编号格式、内容格式、样式ID）
- `CaptionFormat`: 保存Caption格式
- `PlaceholderInfo`: 占位符信息（已废弃）
- `SubSectionInfo`: 子章节信息（已废弃）

**关键方法**：
- `processWord()`: 主处理流程
- `scanWordSections()`: 扫描Word文档中的章节编号
- `findSectionParagraph()`: 查找章节段落
- `findExistingSubSectionsInSection()`: 查找已存在的子章节
- `createSubSectionParagraph()`: 创建子章节段落
- `updateParagraphText()`: 更新段落文本
- `createTableCaption()`: 创建表格标题
- `updateTableCaption()`: 更新表格标题
- `copyTable()`: 复制表格
- `fillTableData()`: 填充表格数据
- `extractSubSectionFormat()`: 提取子章节格式
- `extractCaptionFormat()`: 提取Caption格式

#### ExcelReader

**关键方法**：
- `readExcel()`: 读取Excel文件
- `parseRow()`: 解析单行数据
- `getColumnIndex()`: 获取列索引

---

## 扩展开发指南

### 1. 添加新的Excel列支持

**步骤**：
1. 在`ExcelReader`的`parseRow`方法中添加新列的解析
2. 在`WordProcessor`的`fillTableData`方法中添加新列的数据填充逻辑

**示例**：
```java
// 在ExcelReader中
String newColumn = getCellValue(row, columnIndexMap.get("新列名"));

// 在WordProcessor中
if (columnCount == 2) {
    // 在第二列填充新列数据
    setCellValue(table, rowIndex, 1, testCase.getColumnValue("新列名"));
}
```

### 2. 支持新的表格布局

**步骤**：
1. 在`fillTableData`方法中添加新的布局判断
2. 实现对应的填充方法

**示例**：
```java
private void fillTableData(XWPFTable table, TestCase testCase) {
    int columnCount = table.getRow(0).getTableCells().size();
    
    if (columnCount == 2) {
        fillTwoColumnTable(table, testCase);
    } else if (columnCount == 4) {
        fillFourColumnTable(table, testCase);
    } else if (columnCount == 6) {
        // 新增6列布局支持
        fillSixColumnTable(table, testCase);
    }
}
```

### 3. 自定义格式提取逻辑

**步骤**：
1. 修改`extractSubSectionFormat`或`extractCaptionFormat`方法
2. 添加自定义格式识别逻辑

**示例**：
```java
private SubSectionFormat extractSubSectionFormat(XWPFParagraph templatePara) {
    // 自定义格式提取逻辑
    // 例如：根据特定标记识别格式
    String text = templatePara.getText();
    if (text.contains("[特殊标记]")) {
        // 使用特殊格式
    }
}
```

### 4. 添加新的输出格式

**步骤**：
1. 创建新的生成器类（如`PDFGenerator`）
2. 实现对应的生成方法
3. 在主程序中集成

**示例**：
```java
public class PDFGenerator {
    public void generatePDF(ModuleData moduleData, String outputPath) {
        // PDF生成逻辑
    }
}
```

### 5. 性能优化建议

1. **批量处理**：对于大量数据，考虑分批处理
2. **内存管理**：及时释放不需要的对象
3. **缓存机制**：缓存格式信息，避免重复提取
4. **异步处理**：对于多文件处理，考虑使用多线程

---

## 技术栈

- **Java**: JDK 8+
- **Apache POI**: 5.2.5
  - XSSFWorkbook: Excel读取
  - XWPFDocument: Word处理
- **Apache Commons CLI**: 1.5.0
  - 命令行参数解析
- **Maven**: 项目构建和依赖管理

---

## 注意事项

### 1. Word文档格式限制

- 仅支持.docx格式，不支持.doc格式
- 章节标题必须是独立段落
- 表格模板必须存在且结构正确

### 2. Excel格式限制

- 仅支持.xlsx格式，不支持.xls格式
- 第一行必须是列名
- 必填列不能为空

### 3. 性能考虑

- 单次处理建议不超过100条测试用例
- 大文件处理时注意内存占用
- 建议使用SSD存储以提高IO性能

### 4. 错误处理

- 所有文件操作都有异常处理
- 提供详细的错误提示信息
- 建议在生产环境中添加日志记录

---

*文档版本：v1.0*  
*最后更新：2025-12-10*

