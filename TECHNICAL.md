# Excel数据驱动Word文档自动生成工具 - 技术实现说明

## 📋 目录

1. [架构设计](#架构设计)
2. [核心模块](#核心模块)
3. [关键技术实现](#关键技术实现)
4. [数据流程](#数据流程)
5. [代码结构](#代码结构)
6. [扩展开发指南](#扩展开发指南)

---

## 🏗️ 架构设计

### 整体架构

```
┌─────────────────────┐
│  ExcelToWordTool    │  主程序入口（命令行解析）
└──────────┬──────────┘
           │
           ├──────────────────────┐
           │                      │
    ┌──────▼──────┐      ┌────────▼────────┐
    │ ExcelReader │      │  WordProcessor   │
    │ (数据读取)  │      │  (文档处理)      │
    └──────┬──────┘      └────────┬─────────┘
           │                      │
           │                      │
    ┌──────▼──────┐      ┌────────▼─────────┐
    │ TableData   │      │  TableFill       │
    │ Reader      │      │  Processor       │
    └─────────────┘      └──────────────────┘
           │                      │
           │                      │
    ┌──────▼──────┐      ┌────────▼─────────┐
    │ TableConfig │      │  XWPFDocument    │
    │ (配置管理)  │      │  (Apache POI)    │
    └─────────────┘      └──────────────────┘
```

### 设计原则

1. **单一职责原则**：每个类只负责一个功能模块
2. **开闭原则**：对扩展开放，对修改关闭
3. **依赖倒置**：依赖抽象而非具体实现
4. **配置驱动**：通过配置文件管理识别规则，避免硬编码

---

## 🔧 核心模块

### 1. ExcelReader（Excel数据读取模块）

**职责**：读取Excel文件，解析测试用例数据，并按模块分组

**关键方法**：
- `readExcel(String excelPath)`: 读取Excel文件，自动搜索包含 `模块编号` 列的Sheet

**实现要点**：
- 自动识别测试用例Sheet（包含配置的必填列）
- 支持动态列名，所有列数据都会被读取
- 按模块编号分组数据

**代码位置**：`src/main/java/pub/developers/docautogenbyexcel/reader/ExcelReader.java`

### 2. TableDataReader（表格数据读取模块）

**职责**：读取基本信息和列表型表格数据

**关键方法**：
- `readBasicInfo(String excelPath)`: 读取基本信息表格数据
- `readAllListTableData(String excelPath)`: 读取所有列表型表格数据

**实现要点**：
- 自动识别基本信息Sheet（列结构：`表格名称 | 字段名 | 字段值`）
- 自动识别列表型Sheet（第一列为 `表格名称`）
- 跳过测试用例Sheet（包含 `模块编号` 列）

**代码位置**：`src/main/java/pub/developers/docautogenbyexcel/reader/TableDataReader.java`

### 3. WordProcessor（Word文档处理模块）

**职责**：处理Word模板，生成测试用例表格

**关键方法**：
- `processWord(String wordPath, String outputPath, Map<String, ModuleData> moduleDataMap)`: 主处理方法

**实现要点**：
- 自动定位章节位置（通过样式和内容匹配）
- 自动生成子章节编号（如 `5.2.1`、`5.2.2`）
- 自动提取并应用模板格式
- 处理表格插入和Caption更新

**代码位置**：`src/main/java/pub/developers/docautogenbyexcel/processor/WordProcessor.java`

### 4. TableFillProcessor（表格填充处理模块）

**职责**：填充基本信息和列表型表格

**关键方法**：
- `fillBasicInfoTables(XWPFDocument document, Map<String, BasicInfoData> basicInfoMap)`: 填充基本信息表格
- `fillListTables(XWPFDocument document, Map<String, ListTableData> listDataMap)`: 填充列表型表格

**实现要点**：
- 通过Caption匹配表格
- 通过列名匹配数据列
- 支持动态行数扩展

**代码位置**：`src/main/java/pub/developers/docautogenbyexcel/processor/TableFillProcessor.java`

### 5. TableConfig（配置管理模块）

**职责**：管理表格识别和填充的配置参数

**关键方法**：
- `getTestCaseRequiredColumn()`: 获取测试用例Sheet的必填列名
- `getBasicInfoTableNameColumn()`: 获取基本信息Sheet的表格名称列名
- `isTestCaseSheet(List<String> headerColumns)`: 判断Sheet是否是测试用例Sheet
- `isBasicInfoSheet(List<String> headerColumns)`: 判断Sheet是否是基本信息Sheet
- `isListDataSheet(List<String> headerColumns)`: 判断Sheet是否是列表型Sheet

**实现要点**：
- 从配置文件读取参数（`table-config.properties`）
- 支持UTF-8编码
- 提供默认值

**代码位置**：`src/main/java/pub/developers/docautogenbyexcel/config/TableConfig.java`

---

## 🔑 关键技术实现

### 1. Sheet类型自动识别

**实现原理**：

```java
// 1. 读取Sheet的表头列名
List<String> headerColumns = new ArrayList<>();
for (int c = 0; c < headerRow.getLastCellNum(); c++) {
    String colName = getCellValue(headerRow.getCell(c));
    headerColumns.add(colName != null ? colName.trim() : "");
}

// 2. 根据列名判断Sheet类型
if (headerColumns.contains(config.getTestCaseRequiredColumn())) {
    // 测试用例Sheet
} else if (isBasicInfoSheet(headerColumns)) {
    // 基本信息Sheet
} else if (isListDataSheet(headerColumns)) {
    // 列表型Sheet
}
```

**优势**：
- 不依赖Sheet名称，提高灵活性
- 通过配置文件管理识别规则，易于扩展

### 2. Word文档章节定位

**实现原理**：

```java
// 1. 从目录（TOC）中提取章节编号和名称
// 2. 在正文中查找匹配的章节标题
// 3. 通过样式（Heading 2）确认章节位置
XWPFParagraph sectionPara = findSectionParagraph(document, moduleNumber);
```

**关键点**：
- 区分目录（TOC）和正文内容
- 通过样式ID识别章节（`2` = Heading 1, `3` = Heading 2, `4` = Heading 3）
- 支持章节名称匹配（正文中可能没有编号）

### 3. 格式提取和应用

**实现原理**：

```java
// 1. 从模板段落中提取格式
SubSectionFormat format = extractSubSectionFormat(templatePara);

// 2. 应用到新生成的段落
XWPFRun run = paragraph.createRun();
run.setFontFamily(format.getFontFamily());
run.setFontSize(format.getFontSize());
run.setBold(format.isBold());
```

**关键点**：
- 从 `XWPFRun` 和 `CTRPr` 中提取格式信息
- 支持字体、字号、加粗等属性
- 自动应用模板格式，保持一致性

### 4. 表格Caption匹配

**实现原理**：

```java
// 1. 遍历文档中的所有段落
// 2. 查找Caption样式的段落
// 3. 提取Caption文本并匹配
if (para.getStyle() != null && para.getStyle().equals("Caption")) {
    String captionText = para.getText();
    if (captionText.contains(targetTableName)) {
        // 找到匹配的表格
    }
}
```

**关键点**：
- 通过样式ID识别Caption（`11` = Caption）
- 支持部分匹配（Caption可能包含额外文本）
- 定位Caption后的表格

### 5. 动态编号生成

**实现原理**：

```java
// 1. 查找现有子章节的最大编号
int maxSubSection = findMaxSubSectionNumber(sectionPara);

// 2. 生成新的子章节编号
String newSubSectionNumber = moduleNumber + "." + (maxSubSection + 1);

// 3. 创建子章节段落
XWPFParagraph subSectionPara = createSubSectionParagraph(
    document, afterPara, newSubSectionNumber, testName
);
```

**关键点**：
- 自动计算下一个编号
- 处理编号格式（如 `5.2.1`、`5.2.2`）
- 确保编号连续性

### 6. 插入位置管理

**实现原理**：

```java
// 1. 查找当前章节的最后一个元素（段落或表格）
XWPFParagraph lastElement = findLastElementInSection(document, sectionPara);

// 2. 查找下一个主章节的边界
XWPFParagraph nextMainSection = findNextMainSection(document, sectionPara);

// 3. 确保插入位置在边界内
int insertIndex = calculateInsertIndex(lastElement, nextMainSection);
```

**关键点**：
- 准确识别章节边界
- 避免插入到错误的章节
- 处理表格和段落的混合结构

---

## 📊 数据流程

### 测试用例表格生成流程

```
Excel文件
    ↓
ExcelReader.readExcel()
    ↓
Map<模块编号, ModuleData>
    ↓
WordProcessor.processWord()
    ↓
1. 查找章节位置
2. 提取模板格式
3. 生成子章节和表格
4. 填充数据
    ↓
输出Word文档
```

### 基本信息/列表型表格填充流程

```
Excel文件
    ↓
TableDataReader.readBasicInfo() / readAllListTableData()
    ↓
Map<表格名称, BasicInfoData> / Map<表格名称, ListTableData>
    ↓
TableFillProcessor.fillBasicInfoTables() / fillListTables()
    ↓
1. 查找表格（通过Caption匹配）
2. 匹配列名
3. 填充数据
    ↓
输出Word文档
```

---

## 📁 代码结构

```
src/main/java/pub/developers/docautogenbyexcel/
├── ExcelToWordTool.java          # 主程序入口
├── config/
│   ├── ConfigLoader.java         # 配置文件加载器
│   └── TableConfig.java          # 表格配置管理
├── model/
│   ├── ModuleData.java           # 模块数据模型
│   └── TestCase.java             # 测试用例数据模型
├── reader/
│   ├── ExcelReader.java          # Excel数据读取
│   └── TableDataReader.java      # 表格数据读取
├── processor/
│   ├── WordProcessor.java        # Word文档处理
│   └── TableFillProcessor.java   # 表格填充处理
└── util/
    └── FileUtil.java             # 文件工具类

src/main/resources/
└── table-config.properties       # 表格配置文件
```

---

## 🛠️ 扩展开发指南

### 1. 添加新的Sheet类型

**步骤**：

1. 在 `table-config.properties` 中添加配置：
```properties
newtype.column.required=必填列名
```

2. 在 `TableConfig.java` 中添加识别方法：
```java
public boolean isNewTypeSheet(List<String> headerColumns) {
    String requiredColumn = getProperty("newtype.column.required", "默认值");
    return headerColumns.contains(requiredColumn);
}
```

3. 在 `TableDataReader.java` 中添加读取方法：
```java
public Map<String, NewTypeData> readNewTypeData(String excelPath) {
    // 实现读取逻辑
}
```

### 2. 添加新的表格填充逻辑

**步骤**：

1. 在 `TableFillProcessor.java` 中添加填充方法：
```java
public int fillNewTypeTables(XWPFDocument document, Map<String, NewTypeData> dataMap) {
    // 实现填充逻辑
}
```

2. 在 `ExcelToWordTool.java` 中调用：
```java
int newTypeCount = tableFillProcessor.fillNewTypeTables(document, newTypeMap);
```

### 3. 修改格式提取逻辑

**步骤**：

1. 在 `WordProcessor.java` 中修改 `extractSubSectionFormat` 方法：
```java
private SubSectionFormat extractSubSectionFormat(XWPFParagraph para) {
    // 修改格式提取逻辑
}
```

### 4. 添加新的命令行参数

**步骤**：

1. 在 `ExcelToWordTool.java` 的 `parseArguments` 方法中添加选项：
```java
options.addOption("newparam", true, "新参数说明");
```

2. 在 `ConfigLoader.java` 中添加对应的属性：
```java
private String newParam;

public String getNewParam() {
    return newParam;
}
```

---

## 🔍 调试技巧

### 1. 启用调试日志

在 `table-config.properties` 中设置：
```properties
debug.enabled=true
```

### 2. 检查Sheet识别

在 `TableDataReader.java` 中添加日志：
```java
System.out.println("检查Sheet: " + sheetName);
System.out.println("列名: " + headerColumns);
System.out.println("识别结果: " + isTestCaseSheet(headerColumns));
```

### 3. 检查Word文档结构

在 `WordProcessor.java` 中添加日志：
```java
System.out.println("段落样式: " + para.getStyle());
System.out.println("段落文本: " + para.getText());
System.out.println("样式ID: " + para.getCTP().getPPr().getPStyle().getVal());
```

---

## 📝 注意事项

1. **文件格式**：仅支持 `.docx` 和 `.xlsx` 格式，不支持 `.doc` 和 `.xls`
2. **编码问题**：配置文件使用UTF-8编码，确保正确读取中文
3. **样式匹配**：Word模板必须使用标准样式（Heading 2、Heading 3、Caption）
4. **Caption匹配**：表格Caption必须与Excel中的 `表格名称` 完全匹配
5. **列名匹配**：列表型表格的列名必须与Word表格的表头列名匹配

---

## 📚 相关技术

- **Apache POI**：Java API for Microsoft Office文档处理
- **XWPFDocument**：POI的Word文档处理类
- **XSSFWorkbook**：POI的Excel文档处理类
- **Commons CLI**：命令行参数解析

---

## 📄 许可证

本项目采用 MIT 许可证。

