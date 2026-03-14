# Excel数据驱动Word文档自动生成工具

## 📖 项目简介

本工具是一个基于Excel数据自动填充Word文档的工具，支持：

- ✅ **测试用例表格自动生成**：根据Excel数据自动在Word文档中生成测试用例表格
- ✅ **基本信息表格填充**：自动填充Word文档中的基本信息表格（如软件基本信息）
- ✅ **列表型表格填充**：自动填充各种列表型表格（如接口信息、测试环境等）
- ✅ **智能Sheet识别**：根据Excel内容自动识别Sheet类型，无需固定命名
- ✅ **格式自适应**：自动提取并应用Word模板的格式
- ✅ **Web界面**：提供友好的Web界面，支持文件上传、在线预览和下载
- ✅ **REST API**：提供完整的REST API接口，支持集成到其他系统
- ✅ **数据库存储**：支持将处理后的文档保存到数据库，支持从数据库查询和下载

## 🚀 快速开始

### 系统要求

- JDK 17 及以上
- Maven 3.6 及以上（用于构建）
- Word 2016及以上（.docx格式）
- Excel 2016及以上（.xlsx格式）

### 构建项目

```bash
mvn clean package
```

构建完成后，在 `target` 目录下会生成 `DocAutoGenByExcel-0.0.1-SNAPSHOT.jar` 文件。

## 📝 使用方法

本工具支持两种使用方式：**命令行模式**和**Web界面模式**。

### 方式一：命令行模式

#### 基本用法

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
  -excel "test_data.xlsx" \
  -word "template.docx" \
  -out "output"
```

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
  -excel "2号文档完整填充模板.xlsx" \
  -word "2-XX软件配置项测试报告(公开）.docx" \
  -out "output"
```
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -excel "2号文档完整填充模板.xlsx" -word "2-XX软件配置项测试报告(公开）.docx" -out "output"
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -excel "test_data_enhanced2.0.xlsx" -word "1-XX测试大纲（公开）_副本.docx" -out "output"

#### 命令行参数说明

| 参数 | 说明 | 必填 | 默认值 |
|------|------|------|--------|
| `-excel <路径>` | Excel数据文件路径 | ✅ | - |
| `-word <路径>` | Word模板文件路径 | ✅ | - |
| `-out <路径>` | 输出目录路径 | ❌ | Excel文件同目录 |
| `-config` | 使用配置文件 | ❌ | - |
| `-h, --help` | 显示帮助信息 | ❌ | - |

#### 使用示例

**示例1：基本使用**

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
  -excel "完整测试用例模板.xlsx" \
  -word "2-XX软件配置项测试报告(公开）.docx" \
  -out "output"
```
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
-excel "需求规格说明数据示例.xlsx" \
-word "副本软件需求规格说明框架（公开）.docx" \
-out "output"
**示例2：使用配置文件**

创建 `config.properties` 文件：

```properties
excel.path=test_data.xlsx
word.path=template.docx
output.path=output
```

运行：

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -config
```

**示例3：显示帮助信息**

```bash
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -h
```

### 方式二：Web界面模式

#### 启动Web服务

```bash
# 方式1：直接运行（作为Spring Boot应用）
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar --spring

# 方式2：指定端口
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar --spring --server.port=8080
```

启动成功后，访问：**http://localhost:8080**

#### Web界面功能

1. **文件上传**
   - 支持拖拽上传或点击选择
   - 上传Excel数据文件和Word模板文件
   - 实时显示已选择的文件名称

2. **文档处理**
   - 点击"开始处理"按钮自动填充文档
   - 显示处理进度和结果
   - 自动保存处理后的文档到本地

3. **在线预览**
   - 点击"在线预览"查看生成的文档
   - 支持下载文档到本地

4. **历史记录**
   - 查看所有已处理的文档列表
   - 显示文件大小和创建时间
   - 支持预览、下载和删除操作

5. **文档管理**
   - 刷新文档列表
   - 删除不需要的文档
   - 自动清理旧文件（7天以上）

#### Web界面截图说明

- **上传区域**：两个上传框分别用于Excel和Word文件
- **处理按钮**：选择文件后自动启用
- **结果区域**：处理成功后显示下载和预览链接
- **历史文档**：显示所有已处理的文档，支持批量管理

### REST API接口

Web服务提供完整的REST API，支持集成到其他系统。

#### 基础URL

```
http://localhost:8080/api/documents
```

#### 接口列表

##### 1. 处理文档

**请求**

```http
POST /api/documents/process
Content-Type: multipart/form-data

excel: [Excel文件]
word: [Word文件]
```

**响应**

```json
{
  "success": true,
  "message": "成功处理 4 个模块",
  "outputId": "73c6034c_20251223111858",
  "outputFileName": "2-XX软件配置项测试报告(公开）_20251223111858.docx",
  "moduleCount": 4,
  "downloadUrl": "/api/documents/download/2-XX软件配置项测试报告(公开）_20251223111858.docx"
}
```

**cURL示例**

```bash
curl -X POST http://localhost:8080/api/documents/process \
  -F "excel=@test_data.xlsx" \
  -F "word=@template.docx"
```

##### 2. 获取文档列表

**请求**

```http
GET /api/documents/list
```

**响应**

```json
{
  "success": true,
  "total": 2,
  "documents": [
    {
      "id": "73c6034c_20251223111858",
      "fileName": "2-XX软件配置项测试报告(公开）_20251223111858.docx",
      "originalExcel": "",
      "originalWord": "",
      "fileSize": 118032,
      "createdAt": "2025-12-23T11:18:58.123Z"
    }
  ]
}
```

##### 3. 下载文档

**请求**

```http
GET /api/documents/download/{fileName}
```

**响应**

返回文件流，浏览器会自动下载。

**示例**

```bash
curl -O http://localhost:8080/api/documents/download/2-XX软件配置项测试报告(公开）_20251223111858.docx
```

##### 4. 预览文档

**请求**

```http
GET /api/documents/preview/{fileName}
```

**响应**

返回文件流，Content-Type设置为Word文档类型，浏览器会尝试在线预览。

##### 5. 根据输出ID下载（推荐）

**请求**

```http
GET /api/documents/download/id/{outputId}
```

**响应**

返回文件流，浏览器会自动下载。

**说明**：处理文档时会返回 `outputId`，使用ID下载更可靠，不受文件名变化影响。

##### 6. 删除文档（根据文件名）

**请求**

```http
DELETE /api/documents/{fileName}
```

**响应**

```json
{
  "success": true,
  "message": "文档已删除"
}
```

##### 7. 删除文档（根据输出ID）

**请求**

```http
DELETE /api/documents/id/{outputId}
```

**响应**

```json
{
  "success": true,
  "message": "文档已删除"
}
```

##### 8. 清理旧文件

**请求**

```http
POST /api/documents/cleanup?daysToKeep=7
```

**响应**

```json
{
  "success": true,
  "deletedCount": 5,
  "message": "已清理 5 个旧文件"
}
```

#### API使用示例（JavaScript）

```javascript
// 处理文档
async function processDocument(excelFile, wordFile) {
  const formData = new FormData();
  formData.append('excel', excelFile);
  formData.append('word', wordFile);
  
  const response = await fetch('http://localhost:8080/api/documents/process', {
    method: 'POST',
    body: formData
  });
  
  const result = await response.json();
  console.log('处理结果:', result);
  return result;
}

// 获取文档列表
async function getDocumentList() {
  const response = await fetch('http://localhost:8080/api/documents/list');
  const data = await response.json();
  console.log('文档列表:', data.documents);
  return data.documents;
}

// 下载文档
function downloadDocument(fileName) {
  window.open(`http://localhost:8080/api/documents/download/${encodeURIComponent(fileName)}`);
}
```

## 📋 Excel数据格式

### 1. 测试用例Sheet

**识别规则**：包含 `模块编号` 列的Sheet（Sheet名称可任意）

**格式示例**：

| 模块编号 | 测试项名称 | 标识 | 测试内容 | 测试策略与方法 | 判定准则 | 测试终止条件 | 追踪关系 |
|---------|-----------|------|---------|--------------|---------|------------|---------|
| 5.2 | 登录功能 | F001 | 验证用户登录功能 | 1) 输入正确的用户名和密码 | 正常登录应跳转到首页 | 测试用例执行完成 | 需求文档V1.0 |
| 5.2 | 注册功能 | F002 | 验证用户注册功能 | 1) 输入有效的邮箱和密码 | 正常注册应创建账户 | 测试用例执行完成 | 需求文档V1.0 |
| 5.3 | 密码重置 | F003 | 验证用户密码重置功能 | 1) 输入已注册的邮箱 | 应发送包含重置链接的邮件 | 测试用例执行完成 | 需求文档V1.0 |

**说明**：
- `模块编号` 列是必填的，用于标识数据属于哪个章节（如 `5.2`、`5.3`）
- 其他列名可以自定义，所有列的数据都会被填充到Word表格中
- 同一模块编号的多行数据会生成多个子章节（如 `5.2.1`、`5.2.2`）

### 2. 测试步骤Sheet（可选）

**识别规则**：Sheet名称为 `测试步骤` 或包含 `测试用例标识` 列

**格式示例**：

| 测试用例标识 | 步骤序号 | 输入及操作 | 期望结果与评估标准 | 实测结果 |
|------------|---------|-----------|------------------|---------|
| T_LOGIN | 1 | 打开浏览器，输入系统URL | 正确显示登录页面 | 与预期一致 |
| T_LOGIN | 2 | 输入正确的用户名 | 用户名输入框显示内容 | 与预期一致 |
| T_LOGIN | 3 | 输入正确的密码 | 密码以星号显示 | 与预期一致 |

**说明**：
- `测试用例标识` 必须与测试用例Sheet中的 `标识` 列匹配
- 支持多个测试步骤，会自动扩展表格行数
- 测试步骤会填充到对应测试用例的"测试步骤"子表格中

### 3. 基本信息Sheet

**识别规则**：列结构为 `表格名称 | 字段名 | 字段值` 的Sheet（Sheet名称可任意）

**格式示例**：

| 表格名称 | 字段名 | 字段值 |
|---------|--------|--------|
| 表1.1 被测软件基本信息 | 软件等级 | D |
| 表1.1 被测软件基本信息 | 类型 | 嵌入式 |
| 表1.1 被测软件基本信息 | 行数 | 15000 |
| 表1.1 被测软件基本信息 | 中断源 | 3 |

**说明**：
- `表格名称` 必须与Word文档中表格的Caption（表格标题）完全匹配
- 程序会自动在Word文档中查找对应Caption的表格并填充数据

### 4. 列表型Sheet

**识别规则**：第一列为 `表格名称`，后续列为数据列的Sheet（Sheet名称可任意）

**格式示例**：

| 表格名称 | 序号 | 接口类型 | 方向（被测件而言） | 说明 |
|---------|------|---------|------------------|------|
| 表1.2 被测软件接口信息 | 1 | RS485串口 | 输入 | 接收控制指令 |
| 表1.2 被测软件接口信息 | 2 | 以太网 | 输出 | 发送状态数据 |
| 表1.2 被测软件接口信息 | 3 | CAN总线 | 双向 | 与外部设备通信 |

**格式示例2（测试环境）**：

| 表格名称 | 序号 | 名称 | 版本标识 | 用途 |
|---------|------|------|---------|------|
| 表6.1 测试环境软件项 | 1 | test1 | v1.0 | test |
| 表6.1 测试环境软件项 | 2 | test2 | v1.1 | test |
| 表6.1 测试环境软件项 | 3 | test3 | v1.2 | test |

**说明**：
- `表格名称` 必须与Word文档中表格的Caption（表格标题）完全匹配
- 后续列名必须与Word表格的表头列名匹配
- 程序会自动匹配列名并填充数据

## 📄 Word模板格式

### 1. 测试用例表格

Word模板中需要包含章节标题（如 `5.2 功能测试`），程序会在该章节下自动生成子章节和表格。

**模板示例**：

```
5.2 功能测试

5.2.1 XX测试

表5.2.1 XX测试

[表格内容]
```

**说明**：
- 章节标题使用 `Heading 2` 样式（如 `5.2 功能测试`）
- 子章节标题使用 `Heading 3` 样式（如 `5.2.1 XX测试`）
- 表格Caption使用 `Caption` 样式（如 `表5.2.1 XX测试`）
- 如果模板中已有子章节（如 `5.2.1 XX测试`），程序会替换其内容
- 如果Excel中有更多数据，程序会自动创建新的子章节（如 `5.2.2`、`5.2.3`）

### 2. 基本信息表格和列表型表格

Word模板中需要包含带Caption的表格，Caption必须与Excel中的 `表格名称` 完全匹配。

**模板示例**：

```
表1.1 被测软件基本信息

[表格内容]
```

### 3. 目录更新

程序会自动更新Word文档的目录（Table of Contents），包括：
- 更新现有章节的页码
- 添加新生成的章节（如 `5.2.2`、`5.2.3`）
- 删除占位符章节（包含"XX"的章节）
- 标记目录字段为"需要更新"，打开文档时会自动刷新

## ⚙️ 配置说明

### 应用配置

配置文件位置：`src/main/resources/application.properties`

```properties
# 应用名称
spring.application.name=DocAutoGenByExcel

# 服务器端口
server.port=8080

# 文件上传配置
spring.servlet.multipart.enabled=true
spring.servlet.multipart.max-file-size=50MB
spring.servlet.multipart.max-request-size=100MB
```

### 表格配置

配置文件位置：`src/main/resources/table-config.properties`

```properties
# 测试用例Sheet识别（只要包含此列即为测试用例Sheet）
testcase.required.column=模块编号

# 基本信息Sheet识别（第一列，第二列，第三列）
basicinfo.column.tablename=表格名称
basicinfo.column.fieldname=字段名
basicinfo.column.fieldvalue=字段值

# 列表型Sheet识别（第一列名称）
listdata.column.tablename=表格名称

# 是否启用调试日志
debug.enabled=false
```

**说明**：
- 可以根据实际需求修改列名配置
- 修改后需要重新编译项目

### 存储配置

#### 数据库存储（默认）

文档内容存储在数据库中（MySQL）：
- ✅ 数据集中管理，易于备份和迁移
- ✅ 支持事务，数据一致性好
- ✅ 支持根据ID或文件名查询和下载
- ⚠️ 注意：大文件（>10MB）建议使用本地或S3存储

**MySQL配置**：
- 数据库名：`docautogen`
- 默认用户：`root`（生产环境建议使用专用用户）
- 详细配置说明：参见 [MYSQL_SETUP.md](MYSQL_SETUP.md)

#### 本地存储

文件存储在本地目录：
- 上传文件：`storage/uploads/`
- 输出文件：`storage/outputs/`

配置：
```properties
storage.type=local
storage.save-to-database=true  # 同时保存元数据到数据库
```

#### 云存储（S3 - 可选）

支持AWS S3云存储，配置方式见 [云存储配置](#云存储配置) 章节。

配置：
```properties
storage.type=s3
storage.save-to-database=true  # 同时保存元数据到数据库
```

## ✨ 功能特性

### 1. 智能Sheet识别

- 不依赖Sheet名称，根据内容自动识别Sheet类型
- 测试用例Sheet：包含 `模块编号` 列
- 测试步骤Sheet：包含 `测试用例标识` 列
- 基本信息Sheet：列结构为 `表格名称 | 字段名 | 字段值`
- 列表型Sheet：第一列为 `表格名称`

### 2. 格式自适应

- 自动提取Word模板的格式（字体、字号、样式等）
- 生成的文档格式与模板保持一致
- 支持自定义格式配置
- 自动保持表格格式（对齐、边框、行高等）

### 3. 动态编号

- 自动生成子章节编号（如 `5.2.1`、`5.2.2`）
- 自动生成表格Caption编号（如 `表5.2.1`、`表5.2.2`）
- 支持多模块批量处理
- 自动更新目录（TOC）

### 4. 表格自动填充

- 测试用例表格：根据模块编号自动生成
- 测试步骤表格：支持动态扩展行数，保持格式一致
- 基本信息表格：根据Caption匹配填充
- 列表型表格：根据Caption和列名匹配填充

### 5. Web界面

- 现代化的用户界面
- 拖拽上传文件
- 实时处理状态显示
- 在线预览和下载
- 历史文档管理

### 6. REST API

- 完整的RESTful API
- 支持文件上传和处理
- 文档列表查询
- 文档下载和预览
- 文档删除和清理

## ❓ 常见问题

### Q1: 提示"未找到包含'模块编号'列的Sheet"

**A**: 请确保Excel中至少有一个Sheet包含 `模块编号` 列，列名必须与配置文件中的 `testcase.required.column` 一致。

### Q2: 表格没有填充数据

**A**: 请检查：
1. Excel中的 `表格名称` 是否与Word文档中表格的Caption完全匹配
2. 列表型表格的列名是否与Word表格的表头列名匹配
3. Word文档中的表格Caption是否使用了 `Caption` 样式

### Q3: 生成的子章节位置不对

**A**: 请确保Word模板中的章节标题使用了正确的样式：
- 主章节使用 `Heading 2` 样式
- 子章节使用 `Heading 3` 样式

### Q4: 格式不一致

**A**: 程序会自动提取模板格式，如果格式不一致，请检查：
1. 模板中的格式是否正确设置
2. 是否使用了正确的样式（Heading 2、Heading 3、Caption）

### Q5: 支持.doc格式吗？

**A**: 不支持，仅支持 `.docx` 格式。`.doc` 文件需要先转换为 `.docx` 格式。

### Q6: Web服务无法启动

**A**: 请检查：
1. 端口8080是否被占用：`lsof -i :8080`
2. JDK版本是否为17及以上：`java -version`
3. 查看启动日志中的错误信息

### Q7: 文件上传失败

**A**: 请检查：
1. 文件大小是否超过50MB限制
2. 文件格式是否正确（.xlsx/.xls 和 .docx/.doc）
3. 服务器磁盘空间是否充足

### Q8: 如何修改服务端口？

**A**: 有两种方式：
1. 修改 `application.properties` 中的 `server.port`
2. 启动时指定：`java -jar app.jar --server.port=9090`

## 🔧 云存储配置

### AWS S3配置（可选）

本工具支持将处理后的文档保存到AWS S3云存储。目前代码已实现，但默认未启用。

#### 启用S3存储

1. 添加AWS SDK依赖（已在pom.xml中，但被注释）

2. 配置S3参数（在 `application.properties` 中）：

```properties
# S3存储配置（可选）
storage.type=s3
aws.s3.bucket=your-bucket-name
aws.s3.region=us-east-1
aws.s3.access-key=your-access-key
aws.s3.secret-key=your-secret-key
aws.s3.endpoint=https://s3.amazonaws.com
```

3. 修改 `DocumentService` 以使用S3存储（代码已预留接口）

#### S3存储优势

- 文件持久化存储
- 支持多地域部署
- 自动备份和版本控制
- 可配置访问权限
- 支持CDN加速

**注意**：S3存储功能已实现但默认未启用，需要时取消注释相关代码并配置参数。

## 💾 数据库功能

项目支持将处理后的文档保存到数据库，支持从数据库查询和下载。

### 快速开始

1. **默认使用H2数据库**（无需配置）：
   - 数据库文件自动创建在 `./data/docautogen.mv.db`
   - 可通过 H2 控制台查看：http://localhost:8080/h2-console

2. **使用MySQL/PostgreSQL**（生产环境）：
   - 在 `pom.xml` 中取消注释对应数据库驱动
   - 在 `application.properties` 中配置数据库连接

### 主要功能

- ✅ 自动保存处理后的文档到数据库
- ✅ 支持根据文件名或输出ID查询和下载
- ✅ 支持文档列表查询（按创建时间排序）
- ✅ 支持文档删除（同时删除数据库记录和文件）
- ✅ 支持自动清理旧文档

### 详细说明

参见 [DATABASE.md](DATABASE.md) 文档，包含：
- 数据库配置说明
- 表结构说明
- API接口说明
- 使用示例
- 性能优化建议

## 📚 更多信息

- 技术实现说明：参见 [TECHNICAL.md](TECHNICAL.md)
- 数据库功能说明：参见 [DATABASE.md](DATABASE.md)
- 配置文件说明：参见 `src/main/resources/table-config.properties`
- API文档：启动Web服务后访问 http://localhost:8080

## 📄 许可证

本项目采用 MIT 许可证。
