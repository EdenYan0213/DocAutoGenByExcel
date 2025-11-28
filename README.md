# Excel数据驱动Word多模块动态表格生成工具

## 项目简介

本工具可以根据Excel中的测试用例数据，自动在Word测试报告的指定章节生成带动态编号的测试表格，无需修改原始Word模板。

## 功能特性

- ✅ 零侵入性：完全适配原始Word模板，无需修改模板结构
- ✅ 精准定位：按Word章节标题自动定位填充位置
- ✅ 动态编号：模块内编号自动生成（如5.3.1、5.3.2）
- ✅ 批量处理：支持一次处理多个模块数据
- ✅ 格式保留：完全保留原始模板的格式

## 系统要求

- JDK 8 及以上
- Maven 3.6 及以上（用于构建）
- Word 2016及以上（.docx格式）
- Excel 2016及以上（.xlsx格式）

## 快速开始

### 1. 构建项目

```bash
mvn clean package
```

构建完成后，在 `target` 目录下会生成 `DocAutoGenByExcel-0.0.1-SNAPSHOT.jar` 文件。

### 2. 准备数据文件

#### Excel数据格式

Excel文件必须包含以下列（列名不可修改，顺序可调整）：

| 列名 | 是否必填 | 说明 | 示例 |
|------|---------|------|------|
| 模块编号 | 是 | Word章节对应前缀 | 5.3、6.2 |
| testName | 是 | 测试名称 | 功能1、接口A |
| id | 是 | 测试用例唯一标识 | F001、I002 |
| content | 是 | 测试内容 | 验证功能1的登录逻辑 |
| strategy | 是 | 测试策略与方法 | 1) 模拟正常登录；2) 输入错误密码 |
| criteria | 是 | 判定准则 | 1) 登录成功跳转首页 |
| stopCondition | 是 | 测试终止条件 | 测试用例执行完成 |
| trace | 是 | 追踪关系 | 需求文档V1.0 |

**注意事项：**
- 第一行为列名
- 支持空行自动过滤
- 单元格数据超过500字符正常读取，Word中自动换行

#### Word模板格式

Word模板中需要包含章节标题，格式为：`X.X 模块名称`

例如：
- `5.3 功能测试`
- `6.2 接口测试`

章节标题必须是独立段落，文本内容完全匹配上述格式。

### 3. 运行工具

#### 方式一：命令行参数

```bash
java -jar DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -excel "D:\data.xlsx" -word "D:\template.docx" -out "D:\output"
```

参数说明：
- `-excel`: Excel文件路径（必填）
- `-word`: Word模板文件路径（必填）
- `-out`: 输出目录路径（可选，默认为Excel文件同目录）
- `-h`: 显示帮助信息

#### 方式二：配置文件

1. 复制 `src/main/resources/config.properties.example` 为 `config.properties`
2. 编辑配置文件，填写实际路径：

```properties
excel.path=D:/data.xlsx
word.path=D:/template.docx
output.path=D:/output
```

3. 运行：

```bash
java -jar DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -config
```

或直接运行（无参数时自动使用配置文件）：

```bash
java -jar DocAutoGenByExcel-0.0.1-SNAPSHOT.jar
```

### 4. 输出结果

生成的Word文件命名规则：`原始模板名_生成时间.docx`

例如：`测试报告模板_202511281530.docx`

## 生成的表格格式

工具会在Word指定章节下生成以下格式的表格：

| 测试项名称 | testName值 | 标识 | id值 |
|-----------|-----------|------|------|
| 测试内容 | content值 | | |
| 测试策略与方法 | strategy值 | | |
| 判定准则 | criteria值 | | |
| 测试终止条件 | stopCondition值 | | |
| 追踪关系 | trace值 | | |

每个测试用例会生成：
1. 子标题：`模块编号.序号 测试名称测试`（如：`5.3.1 功能1测试`）
2. 测试表格（6行4列）
3. 空行（最后一个测试用例后无空行）

## 常见问题

### 1. Excel文件读取失败

**问题**：提示"Excel文件路径错误或文件损坏"

**解决**：
- 确认文件路径正确
- 确认文件格式为.xlsx（不支持.xls）
- 确认文件未被其他程序占用

### 2. Word模板中未找到模块

**问题**：提示"Word模板中未找到模块: X.X"

**解决**：
- 确认Word模板中存在对应章节标题
- 确认章节标题格式为：`X.X 模块名称`（X为数字）
- 确认章节标题为独立段落

### 3. Excel缺少必填列

**问题**：提示"Excel缺少必填列: XXX"

**解决**：
- 确认Excel第一行为列名
- 确认列名拼写正确（区分大小写）
- 参考"Excel数据格式"章节检查列名

### 4. 生成的表格格式不正确

**问题**：表格边框、对齐方式不符合要求

**解决**：
- 确认使用Word 2016及以上版本打开
- 如仍有问题，请检查Word模板格式

## 项目结构

```
DocAutoGenByExcel/
├── src/main/java/pub/developers/docautogenbyexcel/
│   ├── ExcelToWordTool.java          # 主程序入口
│   ├── model/                        # 数据模型
│   │   ├── TestCase.java            # 测试用例模型
│   │   └── ModuleData.java          # 模块数据模型
│   ├── reader/                       # Excel读取模块
│   │   └── ExcelReader.java
│   ├── processor/                    # Word处理模块
│   │   └── WordProcessor.java
│   ├── config/                       # 配置模块
│   │   └── ConfigLoader.java
│   └── util/                         # 工具类
│       └── FileUtil.java
└── src/main/resources/
    └── config.properties.example    # 配置文件示例
```

## 技术栈

- **Java**: JDK 8+
- **Apache POI**: 5.2.5（Excel和Word处理）
- **Apache Commons CLI**: 1.5.0（命令行参数解析）
- **Spring Boot**: 4.0.0（框架支持，可选）

## 性能指标

- 数据处理能力：支持单次处理100条以内测试用例
- 处理时间：100条数据处理时间≤10秒
- 内存占用：运行时内存占用≤512MB

## 许可证

本项目采用 Apache License 2.0 许可证。

## 联系方式

如有问题或建议，请联系开发团队。

