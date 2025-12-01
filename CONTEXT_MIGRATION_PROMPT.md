# Excel数据驱动Word多模块动态表格生成工具 - 对话上下文迁移Prompt

## 项目概述
这是一个Java开发的工具，用于从Excel读取测试用例数据，自动填充到Word模板的表格中。核心功能是：
- 读取Excel文件（支持动态列名）
- 识别Word模板中的章节和表格
- 复制模板表格格式并填充数据
- 保持原有表格结构，不创建新行列

## 技术栈
- Java 17
- Apache POI 5.2.5（处理Word和Excel）
- Apache Commons CLI（命令行参数解析）
- Maven构建

## 核心问题与解决方案

### 1. 表格复制与填充
**问题**：需要复制Word模板中的表格格式，为每个子章节创建相同格式的表格并填充数据。

**解决方案**：
- 实现`copyTable()`方法，复制模板表格的完整结构（包括单元格合并、样式等）
- 在复制时只保留标签列内容，清空数据列等待填充
- 使用`insertTestCaseTableAfterParagraph()`方法，优先查找已存在表格，不存在则复制模板

### 2. 动态列名匹配
**问题**：Word表格中的标签列名需要动态匹配Excel的列名，不能硬编码。

**解决方案**：
- 实现`findMatchingColumn()`方法，支持完全匹配、包含匹配、去除空格匹配
- 从Word表格左列读取标签，在Excel列名集合中智能查找
- 移除了所有硬编码的标签映射关系

### 3. 单元格合并处理
**问题**：Word模板表格使用了单元格合并（gridSpan），导致逻辑单元格和物理单元格不一致。

**关键发现**：
- 模板表格行1有4个逻辑单元格，但只有2个物理单元格
- 单元格0和1指向同一个物理单元格（标签列）
- 单元格2和3指向同一个物理单元格（数据列）
- 填充数据时应该填充单元格2（或3），而不是单元格1

**解决方案**：
- 在`fillTwoColumnTable()`和`fillMultiColumnTable()`中，判断单元格数量
- 如果>=3列，填充单元格2；否则填充单元格1

### 4. 章节顺序和插入点
**问题**：子章节和表格需要按正确顺序插入，确保每个大章节内的所有小章节完成后才进入下一个大章节。

**解决方案**：
- 使用顺序插入而非倒序插入
- 每次创建表格后，在表格后创建空段落作为下次插入点
- 检查插入点是否属于同一主章节，遇到其他主章节立即停止查找

### 5. 模板表格复用
**问题**：4.3章节有表格模板，6.2章节没有模板，需要复用4.3的模板。

**解决方案**：
- 保存第一个找到的表格作为全局模板（`globalTemplateTable`）
- 对于没有模板的章节，使用全局模板复制

## 关键代码结构

### 主要方法
1. `processWord()` - 主处理流程
   - 扫描Word章节和占位符
   - 处理已存在的子章节
   - 处理主章节（复制模板表格）
   - 处理没有模板的章节（使用全局模板）

2. `insertTestCaseTableAfterParagraph()` - 插入/填充表格
   - 查找已存在的表格
   - 如果不存在且有模板，复制模板表格
   - 填充数据

3. `copyTable()` - 复制表格
   - 复制表格结构、属性、网格
   - 只复制标签列内容，数据列留空

4. `fillTwoColumnTable()` / `fillMultiColumnTable()` - 填充数据
   - 从Word读取标签
   - 智能匹配Excel列名
   - 填充到正确的单元格（考虑合并）

5. `findMatchingColumn()` - 智能列名匹配
   - 完全匹配
   - 包含匹配
   - 去除空格匹配

### 数据模型
- `TestCase` - 测试用例，包含`Map<String, String> columnData`存储动态列数据
- `ModuleData` - 模块数据，包含多个TestCase

### Excel读取
- `ExcelReader.readExcel()` - 动态读取所有列，不硬编码列名
- 列名存储在`TestCase.columnData`的key中

## 当前状态
✅ 已完成：
- 表格复制功能
- 动态列名匹配
- 单元格合并处理
- 章节顺序控制
- 模板表格复用

⚠️ 待修复：
- 子章节标题显示为"测试"而不是完整名称（如"登录功能测试"）
- 问题原因：`getTestName()`方法查找"testName"列，但Excel列名是"测试项名称"

## 下一步工作
1. 修复`getTestName()`方法，使其能够从`columnData`中查找"测试项名称"列
2. 或者修改调用处，直接使用`getColumnValue("测试项名称")`

## 对话风格
- 用户经常说"继续"，表示继续当前工作
- 用户会提供截图或详细描述问题
- 需要快速定位问题并修复
- 重视代码质量和可维护性
- 使用中文交流

## 重要文件路径
- `/Users/chuang.yan/code/idea/DocAutoGenByExcel/`
- 主代码：`src/main/java/pub/developers/docautogenbyexcel/processor/WordProcessor.java`
- 数据模型：`src/main/java/pub/developers/docautogenbyexcel/model/TestCase.java`
- Excel读取：`src/main/java/pub/developers/docautogenbyexcel/reader/ExcelReader.java`

## 测试数据
- Excel：`test_data.xlsx` - 包含模块编号、测试项名称、标识、测试内容等列
- Word模板：`test_template.docx` - 包含4.3章节的表格模板（6行4列）
- 输出：`./output/test_template_*.docx`

