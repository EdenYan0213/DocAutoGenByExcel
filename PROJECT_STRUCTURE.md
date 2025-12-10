# 项目文件结构说明

本文档说明项目的核心文件结构和文档组织。

## 一、核心文件清单

### 1.1 源代码文件

**主程序**：
- `src/main/java/pub/developers/docautogenbyexcel/ExcelToWordTool.java` - 主程序入口

**数据模型**：
- `src/main/java/pub/developers/docautogenbyexcel/model/TestCase.java` - 测试用例模型
- `src/main/java/pub/developers/docautogenbyexcel/model/ModuleData.java` - 模块数据模型
- `src/main/java/pub/developers/docautogenbyexcel/model/Requirement.java` - 需求模型
- `src/main/java/pub/developers/docautogenbyexcel/model/Traceability.java` - 追溯关系模型

**读取器**：
- `src/main/java/pub/developers/docautogenbyexcel/reader/ExcelReader.java` - 测试用例Excel读取器
- `src/main/java/pub/developers/docautogenbyexcel/reader/RequirementExcelReader.java` - 需求Excel读取器

**处理器**：
- `src/main/java/pub/developers/docautogenbyexcel/processor/WordProcessor.java` - Word文档处理器

**管理器**：
- `src/main/java/pub/developers/docautogenbyexcel/manager/RequirementManager.java` - 需求管理器
- `src/main/java/pub/developers/docautogenbyexcel/manager/TraceabilityManager.java` - 追溯管理器

**生成器**：
- `src/main/java/pub/developers/docautogenbyexcel/generator/RequirementTreeWordGenerator.java` - 需求树Word生成器

**配置和工具**：
- `src/main/java/pub/developers/docautogenbyexcel/config/ConfigLoader.java` - 配置加载器
- `src/main/java/pub/developers/docautogenbyexcel/util/FileUtil.java` - 文件工具类
- `src/main/java/pub/developers/docautogenbyexcel/util/TestDataGenerator.java` - 测试数据生成器

**示例代码**：
- `src/main/java/pub/developers/docautogenbyexcel/example/QuickStartExample.java` - 快速开始示例
- `src/main/java/pub/developers/docautogenbyexcel/example/RequirementTraceabilityExample.java` - 需求追溯示例
- `src/main/java/pub/developers/docautogenbyexcel/example/RequirementTreeWordExample.java` - 需求树Word生成示例

### 1.2 测试代码

- `src/test/java/pub/developers/docautogenbyexcel/RequirementTraceabilityTest.java` - 功能测试（SpringBootTest）

### 1.3 配置文件

- `pom.xml` - Maven配置文件
- `src/main/resources/application.properties` - 应用配置
- `src/main/resources/config.properties.example` - 配置文件示例

### 1.4 说明文档

- `README.md` - 项目主说明文档
- `REQUIREMENT_TRACEABILITY_GUIDE.md` - 需求追溯功能使用指南
- `REQUIREMENT_TREE_WORD_GUIDE.md` - 需求树Word生成指南
- `example_input_output.md` - 输入输出示例文档
- `TEST_README.md` - 测试说明文档

### 1.5 示例文件

- `test_template.docx` - Word模板示例
- `test_data.xlsx` - 测试数据示例
- `output/requirement_tree.docx` - 需求树Word文档示例
- `output/requirement_tree_with_trace.docx` - 带追溯信息的需求树Word文档示例

---

## 二、项目结构

```
DocAutoGenByExcel/
├── README.md                          # 项目主说明文档
├── REQUIREMENT_TRACEABILITY_GUIDE.md  # 需求追溯功能使用指南
├── REQUIREMENT_TREE_WORD_GUIDE.md    # 需求树Word生成指南
├── example_input_output.md            # 输入输出示例
├── TEST_README.md                     # 测试说明文档
├── PROJECT_STRUCTURE.md              # 本文件：项目结构说明
├── pom.xml                            # Maven配置文件
├── test_template.docx                 # Word模板示例
├── test_data.xlsx                     # 测试数据示例
├── src/
│   ├── main/
│   │   ├── java/pub/developers/docautogenbyexcel/
│   │   │   ├── ExcelToWordTool.java          # 主程序入口
│   │   │   ├── DocAutoGenByExcelApplication.java # Spring Boot应用
│   │   │   ├── model/                        # 数据模型
│   │   │   │   ├── TestCase.java
│   │   │   │   ├── ModuleData.java
│   │   │   │   ├── Requirement.java
│   │   │   │   └── Traceability.java
│   │   │   ├── reader/                       # Excel读取器
│   │   │   │   ├── ExcelReader.java
│   │   │   │   └── RequirementExcelReader.java
│   │   │   ├── processor/                    # Word处理器
│   │   │   │   └── WordProcessor.java
│   │   │   ├── manager/                      # 管理器
│   │   │   │   ├── RequirementManager.java
│   │   │   │   └── TraceabilityManager.java
│   │   │   ├── generator/                    # 文档生成器
│   │   │   │   └── RequirementTreeWordGenerator.java
│   │   │   ├── config/                       # 配置
│   │   │   │   └── ConfigLoader.java
│   │   │   ├── util/                         # 工具类
│   │   │   │   ├── FileUtil.java
│   │   │   │   └── TestDataGenerator.java
│   │   │   └── example/                      # 示例代码
│   │   │       ├── QuickStartExample.java
│   │   │       ├── RequirementTraceabilityExample.java
│   │   │       └── RequirementTreeWordExample.java
│   │   └── resources/
│   │       ├── application.properties
│   │       └── config.properties.example
│   └── test/
│       └── java/pub/developers/docautogenbyexcel/
│           └── RequirementTraceabilityTest.java # 功能测试
└── output/                            # 输出目录
    ├── requirement_tree.docx          # 需求树示例
    └── requirement_tree_with_trace.docx # 带追溯信息的需求树示例
```

---

## 三、文档说明

### 3.1 README.md
项目主说明文档，包含：
- 项目简介和核心功能
- 快速开始指南
- Excel和Word格式要求
- 使用方法和示例
- 项目结构和技术栈

### 3.2 REQUIREMENT_TRACEABILITY_GUIDE.md
需求追溯功能详细使用指南，包含：
- 需求分解功能说明
- 需求追溯功能说明
- Excel集成方法
- 完整示例代码
- 输入输出说明

### 3.3 REQUIREMENT_TREE_WORD_GUIDE.md
需求树Word文档生成指南，包含：
- 功能概述
- 使用方法
- Word文档格式说明
- 自定义格式方法

### 3.4 example_input_output.md
输入输出详细示例，包含：
- 代码方式输入输出示例
- Excel方式输入输出示例
- 完整输入输出示例

### 3.5 TEST_README.md
测试说明文档，包含：
- 测试方法列表
- 运行方法（Maven和IDE）
- 测试内容详解
- 快速参考

---

## 四、快速导航

### 4.1 新用户入门
1. 阅读 `README.md` 了解项目
2. 查看 `example_input_output.md` 了解输入输出
3. 运行 `QuickStartExample.java` 快速体验

### 4.2 使用需求分解和追溯功能
1. 阅读 `REQUIREMENT_TRACEABILITY_GUIDE.md`
2. 运行 `RequirementTraceabilityExample.java`
3. 参考示例代码进行开发

### 4.3 生成需求树Word文档
1. 阅读 `REQUIREMENT_TREE_WORD_GUIDE.md`
2. 运行 `RequirementTreeWordExample.java`
3. 查看生成的Word文档示例

### 4.4 运行测试
1. 阅读 `TEST_README.md`
2. 使用Maven或IDE运行测试
3. 查看测试输出

---

*文档版本：v1.0*  
*最后更新：2025-12-01*

