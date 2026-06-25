# Excel数据驱动Word文档自动生成工具

## 项目简介

本工具是一个基于Excel数据自动填充Word文档的企业级工具，支持测试说明（STD）、测试报告（STR）、需求规格说明等多种文档类型的自动生成。适用于软件测试、需求管理、文档自动化等场景。

### 核心能力

- **测试用例表格自动生成** — 根据Excel数据自动在Word文档中生成测试用例表格及子章节
- **基本信息表格填充** — 自动填充Word文档中的基本信息表格（如软件基本信息、被测软件信息等）
- **列表型表格填充** — 自动填充各种列表型表格（如接口信息、测试环境、软件项等）
- **智能Sheet识别** — 根据Excel列内容自动识别Sheet类型（测试用例、基本信息、列表型），无需固定命名
- **格式自适应** — 自动提取并应用Word模板的字体、字号、样式、对齐等格式
- **动态编号** — 自动生成子章节编号（如 5.2.1、5.2.2）和表格Caption编号
- **目录自动更新** — 自动更新Word目录，删除占位符章节（包含"XX"的章节）
- **文档类型支持** — 支持生成软件测试说明（STD）和软件测试报告（STR）
- **测试统计分析** — STR模式下自动统计测试通过率、需求覆盖率和缺陷汇总
- **追溯矩阵** — 自动生成需求-测试用例追溯矩阵，检测"孤儿需求"
- **需求基线** — 支持从SRS文档（Word）中智能解析需求并建立基线
- **导出Excel** — 支持将解析后的需求、测试用例、测试结果导出到Excel
- **需求分解与追溯** — 支持需求树分解、测试用例与需求的追溯关系管理
- **Web管理界面** — 提供现代化的Web界面，支持文件上传、在线预览、下载管理
- **REST API** — 提供完整的RESTful API，支持集成到DevOps/测试管理平台
- **S3云存储**（可选） — 支持将输出文档存储到AWS S3（默认本地存储）

## 系统要求

| 组件 | 最低版本 |
|------|---------|
| JDK | 17+ |
| Maven | 3.6+ |
| Word格式 | .docx |
| Excel格式 | .xlsx |
| 浏览器（Web模式） | Chrome/Firefox/Edge 最新版 |

> **注意**：不支持旧版 `.doc` 和 `.xls` 格式，请先另存为 `.docx` / `.xlsx` 再使用。

## 快速开始

### 1. 克隆并构建

```bash
# 克隆项目
git clone <repo-url>
cd DocAutoGenByExcel

# 构建（跳过测试以加速）
mvn clean package -DskipTests
```

构建完成后，在 `target` 目录下生成 `DocAutoGenByExcel-0.0.1-SNAPSHOT.jar`。

### 2. 准备输入文件

准备两份文件：

- **Excel数据文件**（.xlsx）— 包含要填充的数据
- **Word模板文件**（.docx）— 包含章节结构、表格样式和目标格式

### 3. 运行

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
  -excel "2号文档完整填充模板2.0.xlsx" \
  -word "2-XX软件配置项测试报告(公开）.docx" \
  -out "output"
```

## 使用方式

本工具支持两种使用方式：**命令行模式**和**Web界面模式**。

### 方式一：命令行模式

#### 命令行参数

| 参数 | 说明 | 必填 | 默认值 |
|------|------|------|--------|
| `-excel <路径>` | Excel数据文件路径 | ✅ | — |
| `-word <路径>` | Word模板文件路径 | ✅ | — |
| `-out <路径>` | 输出目录路径 | ❌ | Excel文件所在目录 |
| `-docType <STD\|STR>` | 文档类型 | ❌ | STD |
| `-config` | 使用配置文件 | ❌ | — |
| `-h, --help` | 显示帮助信息 | ❌ | — |

#### 基本示例

**生成软件测试说明（STD，默认）：**

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
  -excel "2号文档完整填充模板2.0.xlsx" \
  -word "2-XX软件配置项测试报告(公开）.docx" \
  -out "output"
```
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar  -excel "2号文档完整填充模板2.0.xlsx"  -word "2-XX软件配置项测试报告(公开）.docx" -out "output"

**生成软件测试报告（STR）：**

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
  -excel "测试数据.xlsx" \
  -word "测试报告模板.docx" \
  -out "output" \
  -docType STR
```

**生成测试大纲：**

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
  -excel "test_data_enhanced2.0.xlsx" \
  -word "1-XX测试大纲（公开）_副本.docx" \
  -out "output"
```

**生成需求规格说明：**

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
  -excel "需求规格说明数据示例.xlsx" \
  -word "副本软件需求规格说明框架（公开）.docx" \
  -out "output"
```

#### 使用配置文件

创建 `config.properties` 文件：

```properties
excel.path=D:/data/测试数据.xlsx
word.path=D:/data/模板文件.docx
output.path=D:/output
```

运行：

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -config
```

配置文件加载顺序：命令行参数 > `config.properties`（当前目录）> 默认值。

#### 显示帮助

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -h
```

### 方式二：Web界面模式

#### 启动Web服务

```bash
# 默认端口8080启动
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar --spring

# 指定端口
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar --spring --server.port=9090
```

启动成功后访问：**http://localhost:8080**

#### Web界面功能

1. **文件上传** — 拖拽或点击上传Excel数据文件和Word模板文件
2. **文档处理** — 点击"开始处理"，自动填充并生成文档
3. **在线预览** — 支持PDF预览和Word文档预览
4. **历史管理** — 查看、下载、删除历史处理记录
5. **自动清理** — 自动清理7天前的旧文件（可配置）

## 配置说明

### 应用配置

文件位置：`src/main/resources/application.properties`

```properties
# 应用名称
spring.application.name=DocAutoGenByExcel

# 服务器端口
server.port=8080

# 文件上传大小限制
spring.servlet.multipart.max-file-size=50MB
spring.servlet.multipart.max-request-size=100MB

# 存储类型：local（默认）或 s3
storage.type=local
```

### 表格识别配置

文件位置：`src/main/resources/table-config.properties`

```properties
# 测试用例Sheet识别列名（包含此列即识别为测试用例Sheet）
testcase.required.column=模块编号

# 基本信息Sheet三列结构
basicinfo.column.tablename=表格名称
basicinfo.column.fieldname=字段名
basicinfo.column.fieldvalue=字段值

# 列表型Sheet第一列名称
listdata.column.tablename=表格名称

# 调试日志开关
debug.enabled=false
```

可以根据实际数据格式修改列名，修改后重新编译即可。

### 环境变量

| 变量名 | 说明 | 默认值 |
|--------|------|--------|
| `DOCAUTOGEN_PYTHON` | Python可执行文件路径（PDF转换用） | 自动检测 `.venv/Scripts/python.exe` 或 `python` |

### 存储配置

#### 本地存储（默认）

```
storage/
├── uploads/    # 上传的文件
└── outputs/    # 生成的文档
```

#### S3云存储（可选）

修改 `application.properties`：

```properties
storage.type=s3
aws.s3.bucket=my-bucket
aws.s3.region=us-east-1
aws.s3.access-key=AKIA...
aws.s3.secret-key=...
# aws.s3.endpoint=https://s3.custom-endpoint.com  # 可选，兼容S3的服务
```

同时需要在 `pom.xml` 中取消注释AWS SDK依赖：

```xml
<dependency>
    <groupId>software.amazon.awssdk</groupId>
    <artifactId>s3</artifactId>
    <version>2.20.0</version>
</dependency>
```

## Excel数据格式

### 1. 测试用例Sheet

**识别规则**：包含 `模块编号` 列的Sheet（Sheet名称可任意）。

| 模块编号 | 测试项名称 | 标识 | 测试内容 | 测试策略与方法 | 判定准则 | 测试终止条件 | 追踪关系 |
|---------|-----------|------|---------|--------------|---------|------------|---------|
| 5.2 | 登录功能 | F001 | 验证用户登录功能 | 1) 输入正确的用户名和密码 | 正常登录应跳转到首页 | 用例执行完成 | 需求文档V1.0 |
| 5.3 | 密码重置 | F002 | 验证密码重置功能 | 1) 输入已注册的邮箱 | 应发送重置链接邮件 | 用例执行完成 | 需求文档V1.0 |

- `模块编号` 必填，标识数据所属章节（如 5.2、5.3）
- 其他列名自定义，所有列数据填充到Word表格
- 同一模块编号的多行数据生成多个子章节（5.2.1、5.2.2）

### 2. 测试步骤Sheet（可选）

**识别规则**：Sheet名为 `测试步骤` 或包含 `测试用例标识` 列。

| 测试用例标识 | 步骤序号 | 输入及操作 | 期望结果与评估标准 |
|------------|---------|-----------|------------------|
| T_LOGIN | 1 | 打开浏览器，输入URL | 正确显示登录页面 |
| T_LOGIN | 2 | 输入正确的用户名 | 输入框显示内容 |

- `测试用例标识` 必须与测试用例Sheet中的 `标识` 列匹配
- 会自动扩展表格行数填充测试步骤

### 3. 基本信息Sheet

**识别规则**：列结构为 `表格名称 | 字段名 | 字段值` 的Sheet。

| 表格名称 | 字段名 | 字段值 |
|---------|--------|--------|
| 表1.1 被测软件基本信息 | 软件等级 | D |
| 表1.1 被测软件基本信息 | 类型 | 嵌入式 |
| 表1.1 被测软件基本信息 | 行数 | 15000 |

- `表格名称` 必须与Word文档中的表格Caption完全匹配

### 4. 列表型Sheet

**识别规则**：第一列为 `表格名称`，后续为数据列。

| 表格名称 | 序号 | 接口类型 | 方向（被测件而言） | 说明 |
|---------|------|---------|------------------|------|
| 表1.2 被测软件接口信息 | 1 | RS485串口 | 输入 | 接收控制指令 |
| 表6.1 测试环境软件项 | 1 | test1 | v1.0 | test |

- `表格名称` 必须与Word文档中的表格Caption完全匹配
- 后续列名必须与Word表格的表头列名匹配

### 5. 需求Sheet（SRS/需求基线）

**识别规则**：包含 `ReqID` 或 `需求ID` 列的Sheet。

| ReqID | ReqTitle | Description | Priority |
|-------|---------|-------------|----------|
| REQ-001 | 用户登录功能 | 系统应支持用户通过用户名密码登录 | 高 |
| REQ-002 | 用户注册功能 | 系统应支持新用户注册 | 中 |

- 支持 `ReqID`/`需求ID`/`RequirementId` 多种列名
- Priority支持：高/中/低 或 HIGH/MEDIUM/LOW

### 6. 测试结果Sheet（STR）

**识别规则**：包含 `TCID`/`测试用例标识`/`标识` 列的Sheet。

| TCID | ExecDate | Result | DefectID |
|------|---------|--------|----------|
| TC-001 | 2026-04-23 | 通过 | |
| TC-002 | 2026-04-23 | 失败 | BUG-001 |

- Result支持：通过/失败/阻塞 或 PASS/FAIL/BLOCKED
- DefectID 可选，用于缺陷汇总统计

### 7. Meta配置Sheet（可选）

**识别规则**：包含 `Key`/`键` 列的Sheet。用于存储全局配置项。

## Word模板格式

### 测试用例章节结构

```
5.2 功能测试                          ← Heading 2样式
    5.2.1 登录功能测试                ← Heading 3样式（或已被占位符5.x替换）
    表5.2.1 登录功能测试              ← Caption样式
    [表格内容]
    5.2.2 注册功能测试                ← 程序自动创建
    表5.2.2 注册功能测试
    [表格内容]
```

- **主章节**：使用 `Heading 2` 样式，如 `5.2 功能测试`
- **子章节**：使用 `Heading 3` 样式，程序自动生成或替换
- **表格Caption**：使用 `Caption` 样式，如 `表5.2.1 XX测试`
- **占位符**：使用 `5.x` 或 `5.3.x` 格式，程序自动展开为实际子章节

### 基本信息表格

```
表1.1 被测软件基本信息           ← 带Caption的表格
[表格内容]
```

Caption必须与Excel数据中的 `表格名称` 完全匹配。

### 列表型表格

```
表1.2 被测软件接口信息           ← 带Caption的表格
[表头行 + 数据行]
```

Caption必须与Excel数据中的 `表格名称` 完全匹配，列名必须匹配。

## REST API

启动Web服务后，提供以下REST API接口。基础URL：`http://localhost:8080`

### 文档处理API

所有文档处理API通过 `DocumentController` 提供，路径前缀 `/api/documents`。

#### POST /api/documents/process — 处理文档

上传Excel和Word文件，生成填充后的文档。

**请求**：`multipart/form-data`

| 字段 | 类型 | 说明 | 必填 |
|------|------|------|------|
| excel | File | Excel数据文件（.xlsx/.xls） | ✅ |
| word | File | Word模板文件（.docx/.doc） | ✅ |
| docType | String | 文档类型：STD 或 STR（默认STD） | ❌ |

**响应示例**：

```json
{
  "success": true,
  "message": "成功处理 4 个模块",
  "outputId": "73c6034c_20251223111858",
  "outputFileName": "2-XX软件配置项测试报告(公开）_20251223111858.docx",
  "docType": "STD",
  "moduleCount": 4,
  "downloadUrl": "/api/documents/download/2-XX软件配置项测试报告(公开）_20251223111858.docx",
  "downloadUrlById": "/api/documents/download/id/73c6034c_20251223111858"
}
```

**cURL示例**：

```bash
curl -X POST http://localhost:8080/api/documents/process \
  -F "excel=@测试数据.xlsx" \
  -F "word=@模板文件.docx" \
  -F "docType=STR"
```

#### GET /api/documents/list — 获取文档列表

```json
{
  "success": true,
  "total": 2,
  "documents": [
    {
      "id": "73c6034c_20251223111858",
      "fileName": "2-XX软件配置项测试报告(公开）_20251223111858.docx",
      "fileSize": 118032,
      "createdAt": "2025-12-23T11:18:58.123Z"
    }
  ]
}
```

#### GET /api/documents/download/{fileName} — 下载文档（按文件名）

#### GET /api/documents/download/id/{outputId} — 下载文档（按输出ID）

推荐使用按输出ID下载，不受文件名变化影响。

#### GET /api/documents/preview/{fileName} — 预览文档

设置 Content-Type 为 `application/vnd.openxmlformats-officedocument.wordprocessingml.document`，浏览器会尝试在线预览。

#### GET /api/documents/preview/pdf/{fileName} — PDF预览

调用 Python 脚本将 docx 转换为 PDF 后返回，需要 Python 环境和 docx 转 PDF 依赖。

#### GET /api/documents/preview/id/{outputId} — 按输出ID预览

#### GET /api/documents/preview/pdf/id/{outputId} — 按输出ID的PDF预览

#### DELETE /api/documents/{fileName} — 删除文档（按文件名）

```json
{ "success": true, "message": "文档已删除" }
```

#### DELETE /api/documents/id/{outputId} — 删除文档（按输出ID）

#### POST /api/documents/cleanup?daysToKeep=7 — 清理旧文件

```json
{ "success": true, "deletedCount": 5, "message": "已清理 5 个旧文件" }
```

### 需求管理API

所有需求管理API通过 `RequirementController` 提供，路径前缀 `/api/requirements`。

#### POST /api/requirements/parse — 解析SRS文档

上传Word格式的SRS需求规格说明文档，智能解析需求条目。

**请求**：`multipart/form-data`

| 字段 | 类型 | 说明 | 必填 |
|------|------|------|------|
| srs | File | SRS文档（.docx） | ✅ |
| projectName | String | 项目名称 | ❌ |

**响应**：返回解析后的需求列表和校验报告。

#### POST /api/requirements/confirm — 确认并导出需求

将确认后的需求写入Excel。

**请求**：`multipart/form-data`

| 字段 | 类型 | 说明 |
|------|------|------|
| requirements | String (JSON) | 需求列表JSON数组 |
| excel | File (可选) | 已有的Excel文件（追加写入） |

**需求JSON格式**：

```json
[
  {
    "reqId": "REQ-001",
    "title": "用户登录功能",
    "description": "系统应支持用户通过用户名密码登录",
    "priority": "高"
  }
]
```

#### GET /api/requirements/download/{fileName} — 下载需求Excel

### 需求基线API

通过 `BaselineController` 提供，路径前缀 `/api/baseline`。

#### POST /api/baseline/build — 建立需求基线

从SRS文档解析需求并写入Excel，建立需求基线。

| 参数 | 说明 | 必填 |
|------|------|------|
| srsPath | SRS文档本地路径 | ✅ |
| excelPath | 目标Excel路径 | ✅ |

```json
{ "success": true, "requirementCount": 25, "message": "需求基线建立完成" }
```

### JavaScript使用示例

```javascript
// 处理文档
async function processDocument(excelFile, wordFile, docType = 'STD') {
  const formData = new FormData();
  formData.append('excel', excelFile);
  formData.append('word', wordFile);
  formData.append('docType', docType);

  const response = await fetch('http://localhost:8080/api/documents/process', {
    method: 'POST',
    body: formData
  });
  return await response.json();
}

// 获取文档列表
async function getDocumentList() {
  const response = await fetch('http://localhost:8080/api/documents/list');
  const data = await response.json();
  return data.documents;
}

// 下载文档
function downloadDocument(fileName) {
  window.open(`/api/documents/download/${encodeURIComponent(fileName)}`);
}

// 解析SRS文档
async function parseSrs(srsFile, projectName) {
  const formData = new FormData();
  formData.append('srs', srsFile);
  if (projectName) formData.append('projectName', projectName);

  const response = await fetch('/api/requirements/parse', {
    method: 'POST',
    body: formData
  });
  return await response.json();
}
```

## 功能特性详解

### 1. 智能Sheet识别

不依赖Sheet名称，根据列内容自动匹配：

| Sheet类型 | 识别规则 |
|-----------|---------|
| 测试用例 | 包含 `模块编号` 列 |
| 测试步骤 | Sheet名为"测试步骤"或包含 `测试用例标识` 列 |
| 基本信息 | 列结构为 `表格名称 | 字段名 | 字段值` |
| 列表型数据 | 第一列为 `表格名称` |
| 需求数据 | 包含 `ReqID` 或 `需求ID` 列 |
| 测试结果 | 包含 `TCID` 列及结果列 |

### 2. 格式自适应

- 自动提取Word模板中子章节标题（Heading 3）的字体、字号、加粗格式
- 自动提取Caption（题注）的样式和格式
- 生成的文档格式与模板保持一致
- 自动处理表格对齐、边框、行高

### 3. 动态编号

- 自动生成子章节编号：`主章节.序号`（如 5.2.1）
- 自动生成表格Caption编号：`表主章节.序号`（如 表5.2.1）
- 支持占位符展开：`5.x` → `5.1 XX测试`、`5.2 YY测试`
- 自动更新Word目录（TOC），含占位符章节清理

### 4. 表格自动填充

| 表格类型 | 填充策略 |
|----------|---------|
| 测试用例表格 | 按模块编号分组，每行生成一个子章节+表格 |
| 测试步骤表格 | 按测试用例标识匹配，动态扩展行数 |
| 基本信息表格 | 按Caption匹配，按字段名填充 |
| 列表型表格 | 按Caption+列名匹配，逐行填充 |
| 追溯表 | 根据追踪关系字段自动生成（如 表9.1） |

### 5. 测试统计（STR模式）

STR生成器自动计算并输出：

- **测试统计概览**：总用例数、已执行数、通过数、失败数、阻塞数、通过率
- **需求通过率**：每个需求的测试覆盖率、通过率百分比
- **缺陷汇总**：按缺陷ID聚合受影响的用例和需求
- **测试结论建议**：根据通过率和需求覆盖情况给出建议

### 6. 追溯矩阵（STD模式）

STD生成器自动生成：

- 需求→测试用例的双向追溯映射
- "孤儿需求"检测（无任何测试用例覆盖的需求）
- 追溯矩阵统计（总需求数、已覆盖数、覆盖率）

### 7. 需求管理

- **需求解析**：从Word格式的SRS文档中智能提取需求条目（编号、标题、描述）
- **需求基线**：将解析的需求写入Excel，建立可追踪的基线版本
- **需求分解**：支持根需求→子需求→叶子需求的多级分解
- **追溯关系**：建立需求与测试用例的双向追溯，支持覆盖率计算
- **需求树导出**：将需求树结构输出为Word文档

### 8. 目录更新

- 自动更新Word文档目录（Table of Contents）
- 为新增的子章节添加目录条目
- 删除占位符章节（包含"XX"的章节）
- 标记目录字段为"需要更新"，打开文档时自动刷新

## 构建与开发

### 构建命令

```bash
# 完整构建（含测试）
mvn clean package

# 跳过测试快速构建
mvn clean package -DskipTests

# 仅编译
mvn compile

# 运行测试
mvn test
```

### 项目结构

```
DocAutoGenByExcel/
├── pom.xml                          # Maven构建配置
├── config.properties                # 运行配置文件
├── src/
│   └── main/
│       ├── java/pub/developers/docautogenbyexcel/
│       │   ├── DocAutoGenByExcelApplication.java   # Spring Boot启动入口
│       │   ├── ExcelToWordTool.java                # CLI主入口
│       │   ├── baseline/                           # 需求基线（SRS解析）
│       │   ├── builder/                            # Word文档构建门面
│       │   ├── config/                             # 配置加载类
│       │   ├── controller/                         # REST API控制器
│       │   ├── dto/                                # 数据传输对象
│       │   ├── example/                            # 快速开始示例
│       │   ├── generator/                          # 文档生成器（STD/STR/需求树）
│       │   ├── hub/                                # 数据中枢（Excel读写）
│       │   ├── importer/                           # 测试用例/结果导入
│       │   ├── manager/                            # 需求/追溯管理器
│       │   ├── model/                              # 数据模型
│       │   ├── processor/                          # Word/表格处理核心
│       │   ├── reader/                             # Excel读取器
│       │   ├── service/                            # 业务服务层
│       │   ├── util/                               # 工具类
│       │   └── writer/                             # 数据写入器
│       └── resources/
│           ├── application.properties              # Spring Boot配置
│           ├── table-config.properties             # 表格识别配置
│           ├── config.properties.example           # 配置文件示例
│           └── static/index.html                   # Web前端界面
├── storage/
│   ├── uploads/     # 上传文件目录
│   └── outputs/     # 输出文件目录
└── target/          # 构建输出（jar包）
```

### 技术架构

```
┌─────────────────────────────────────────────────────────────┐
│                       用户入口                                │
│  ┌──────────┐  ┌──────────┐  ┌───────────────────────────┐  │
│  │ CLI模式  │  │ Web界面  │  │ REST API（第三方集成）     │  │
│  └────┬─────┘  └────┬─────┘  └───────────┬───────────────┘  │
├───────┴──────────────┴─────────────────────┴────────────────┤
│                     控制器/入口层                            │
│    ExcelToWordTool    DocumentController    BaselineController│
├─────────────────────────────────────────────────────────────┤
│                     业务服务层                               │
│    DocumentService         RequirementService              │
├─────────────────────────────────────────────────────────────┤
│                   文档生成引擎                                │
│  ┌──────────┐  ┌──────────┐  ┌─────────────────────────┐   │
│  │STDGenerator│  │STRGenerator│  │RequirementTreeGenerator│  │
│  └─────┬────┘  └─────┬────┘  └───────────┬─────────────┘   │
│        └──────────────┴───────────────────┘                  │
│                      AbstractDocumentGenerator               │
├─────────────────────────────────────────────────────────────┤
│                    Word文档构建层                             │
│    WordDocumentBuilder → WordProcessor → POI API            │
│    TableFillProcessor → XWPFDocument                        │
├─────────────────────────────────────────────────────────────┤
│                    数据中枢层                                │
│    DataHub ← ExcelDataHub → ExcelReader / TableDataReader  │
│         → Apache POI (XSSFWorkbook)                         │
├─────────────────────────────────────────────────────────────┤
│                   存储层                                     │
│    本地文件系统 (storage/)    或     AWS S3 (可选)           │
└─────────────────────────────────────────────────────────────┘
```

### 处理流程

```
Excel数据文件                         Word模板文件
     │                                    │
     ▼                                    ▼
┌──────────────────────────────────────────────┐
│          1. 数据提取 (extractData)             │
│  ┌──────────┐ ┌────────┐ ┌────────────────┐  │
│  │模块数据  │ │基本信息│ │列表型/需求/结果│  │
│  └──────────┘ └────────┘ └────────────────┘  │
├──────────────────────────────────────────────┤
│          2. 内容生成 (generateContent)         │
│  ┌────────────────────────────────────────┐   │
│  │ WordProcessor → 扫描章节/占位符/子章节 │   │
│  │ → 创建新子章节 → 复制模板表格 → 填充  │   │
│  │ → 更新Caption → 更新目录               │   │
│  └────────────────────────────────────────┘   │
├──────────────────────────────────────────────┤
│          3. 后处理保存 (save)                  │
│  ┌────────────────────────────────────────┐   │
│  │ 填充基本信息表/列表型表/追溯表          │   │
│  │ 生成统计信息/追溯矩阵（STD/STR）       │   │
│  └────────────────────────────────────────┘   │
├──────────────────────────────────────────────┤
│                    ▼                           │
│           输出文件 (.docx)                     │
└──────────────────────────────────────────────┘
```

## 常见问题

### Q1: 提示"未找到包含'模块编号'列的Sheet"

检查：
1. Excel中至少有一个Sheet包含 `模块编号` 列
2. 列名与 `table-config.properties` 中的 `testcase.required.column` 一致
3. 如果使用了自定义列名，需修改配置文件后重新编译

### Q2: 表格没有填充数据

检查：
1. Excel中的 `表格名称` 是否与Word文档中表格的Caption完全匹配（包括空格）
2. 列表型表格的列名是否与Word表格的表头列名匹配
3. Word文档中的表格Caption是否使用了 `Caption` 样式
4. Word模板中表格的格式是否标准（第一行为表头）

### Q3: 生成的子章节位置不对

检查Word模板：
- 主章节使用 `Heading 2` 样式（如 `5.2 功能测试`）
- 子章节使用 `Heading 3` 样式（如 `5.2.1 XX测试`）
- 占位符使用 `5.x` 格式（不限定样式，会被程序识别）

### Q4: 格式不一致

程序会自动提取模板格式。如果出现不一致：
1. 确保模板中的格式已正确设置（字体、字号、加粗等）
2. 确保使用了正确的样式（Heading 2、Heading 3、Caption）
3. 更新目录后打开Word文档时会自动刷新

### Q5: 支持.doc格式吗？

不支持。仅支持 `.docx` 格式。`.doc` 文件需要在Word中"另存为"→"Word文档(*.docx)"。

### Q6: Web服务无法启动

```bash
# 检查端口占用
netstat -ano | findstr :8080

# 检查JDK版本
java -version   # 需要17+

# 查看详细日志
java -jar app.jar --spring --debug
```

### Q7: 文件上传失败

- 文件大小是否超过50MB（可在 `application.properties` 中调整）
- 文件格式是否为 .xlsx/.docx
- `storage/uploads/` 目录是否存在、可写

### Q8: 如何修改服务端口？

```bash
# 方法1：启动时指定
java -jar app.jar --spring --server.port=9090

# 方法2：修改配置文件
# 编辑 application.properties
server.port=9090
```

### Q9: SRS需求解析不准确？

需求解析使用启发式规则结合LLM（大语言模型）。解析不准确时：
1. 检查SRS文档中的需求编号格式是否规范（如 REQ-001、3.1.1）
2. 确保需求使用了标题样式（Heading 1/2/3）
3. 解析后通过 `/api/requirements/confirm` 接口可以手动调整

### Q10: PDF预览失败？

```bash
# 确保Python环境可用
python --version

# 需要安装pdf转换库
pip install python-docx reportlab

# 或通过环境变量指定Python路径
set DOCAUTOGEN_PYTHON=C:\Python39\python.exe
```

PDF预览依赖于 `scripts/convert_docx_to_pdf.py` 脚本，如果该脚本不存在则PDF预览不可用。

## 许可证

本项目采用 MIT 许可证。
