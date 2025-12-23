# 数据库功能说明

## 概述

DocAutoGen 现在支持将处理后的文档保存到数据库中，支持从数据库查询和下载文档。

## 数据库配置

### MySQL 配置（默认）

项目使用 MySQL 数据库，适合生产环境：

```properties
spring.datasource.url=jdbc:mysql://localhost:3306/docautogen?useUnicode=true&characterEncoding=utf8&useSSL=false&serverTimezone=Asia/Shanghai
spring.datasource.driver-class-name=com.mysql.cj.jdbc.Driver
spring.datasource.username=root
spring.datasource.password=your-password
spring.jpa.database-platform=org.hibernate.dialect.MySQLDialect
spring.jpa.hibernate.ddl-auto=update
spring.jpa.show-sql=false
```

**重要**：请将 `your-password` 替换为实际的数据库密码。

### 快速设置

1. **创建数据库**：
```sql
CREATE DATABASE docautogen CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

2. **配置应用**：编辑 `application.properties`，设置数据库连接信息

3. **启动应用**：应用会自动创建表结构

详细配置说明：参见 [MYSQL_SETUP.md](MYSQL_SETUP.md)

### H2 数据库（可选，用于开发测试）

如果需要使用 H2 数据库进行开发测试：

1. 在 `pom.xml` 中取消注释 H2 依赖
2. 在 `application.properties` 中注释 MySQL 配置，启用 H2 配置

H2 控制台：http://localhost:8080/h2-console

### 使用 MySQL（生产环境）

1. 在 `pom.xml` 中取消注释 MySQL 驱动依赖
2. 在 `application.properties` 中配置：

```properties
spring.datasource.url=jdbc:mysql://localhost:3306/docautogen?useUnicode=true&characterEncoding=utf8&useSSL=false&serverTimezone=Asia/Shanghai
spring.datasource.driver-class-name=com.mysql.cj.jdbc.Driver
spring.datasource.username=root
spring.datasource.password=your-password
spring.jpa.database-platform=org.hibernate.dialect.MySQLDialect
```

### 使用 PostgreSQL

1. 在 `pom.xml` 中取消注释 PostgreSQL 驱动依赖
2. 在 `application.properties` 中配置：

```properties
spring.datasource.url=jdbc:postgresql://localhost:5432/docautogen
spring.datasource.driver-class-name=org.postgresql.Driver
spring.datasource.username=postgres
spring.datasource.password=your-password
spring.jpa.database-platform=org.hibernate.dialect.PostgreSQLDialect
```

## 存储类型

### database（默认）

文档内容直接存储在数据库中（BLOB 字段）：

- ✅ 优点：数据集中管理，易于备份和迁移
- ✅ 支持事务，数据一致性好
- ⚠️ 注意：大文件可能影响数据库性能

### local

文档存储在本地文件系统，数据库只保存元数据：

```properties
storage.type=local
storage.save-to-database=true  # 同时保存元数据到数据库
```

### s3

文档存储在 AWS S3，数据库只保存元数据：

```properties
storage.type=s3
storage.save-to-database=true  # 同时保存元数据到数据库
```

## 数据库表结构

### documents 表

| 字段 | 类型 | 说明 |
|------|------|------|
| id | BIGINT | 主键，自增 |
| output_id | VARCHAR(100) | 输出文档ID（业务标识），唯一 |
| output_file_name | VARCHAR(500) | 输出文件名 |
| original_excel_name | VARCHAR(500) | 原始Excel文件名 |
| original_word_name | VARCHAR(500) | 原始Word文件名 |
| content | BLOB | 文档内容（当 storage_type='database' 时） |
| file_size | BIGINT | 文件大小（字节） |
| module_count | INT | 处理的模块数量 |
| message | VARCHAR(1000) | 处理消息 |
| created_at | TIMESTAMP | 创建时间 |
| updated_at | TIMESTAMP | 更新时间 |
| local_file_path | VARCHAR(1000) | 本地文件路径（当 storage_type='local' 时） |
| s3_key | VARCHAR(1000) | S3存储路径（当 storage_type='s3' 时） |
| storage_type | VARCHAR(20) | 存储类型：database, local, s3 |

## API 接口

### 1. 处理文档（自动保存到数据库）

```http
POST /api/documents/process
Content-Type: multipart/form-data

excel: [Excel文件]
word: [Word文件]
```

响应包含 `outputId`，可用于后续查询和下载。

### 2. 根据文件名下载

```http
GET /api/documents/download/{fileName}
```

### 3. 根据输出ID下载（推荐）

```http
GET /api/documents/download/id/{outputId}
```

### 4. 获取文档列表

```http
GET /api/documents/list
```

响应示例：

```json
{
  "success": true,
  "total": 2,
  "documents": [
    {
      "id": "73c6034c_20251223111858",
      "fileName": "2-XX软件配置项测试报告(公开）_20251223111858.docx",
      "originalExcel": "test_data.xlsx",
      "originalWord": "template.docx",
      "fileSize": 118032,
      "createdAt": "2025-12-23T11:18:58.123Z"
    }
  ]
}
```

### 5. 根据文件名删除

```http
DELETE /api/documents/{fileName}
```

### 6. 根据输出ID删除

```http
DELETE /api/documents/id/{outputId}
```

## 使用示例

### Java 代码示例

```java
// 处理文档（自动保存到数据库）
ProcessResult result = documentService.processDocuments(
    excelStream, "data.xlsx",
    wordStream, "template.docx"
);

// 从数据库获取文档
byte[] content = documentService.getOutputDocumentById(result.outputId());

// 获取文档列表
List<DocumentInfo> documents = documentService.listOutputDocuments();

// 删除文档
documentService.deleteDocumentById(result.outputId());
```

### JavaScript 示例

```javascript
// 处理文档
const formData = new FormData();
formData.append('excel', excelFile);
formData.append('word', wordFile);

const response = await fetch('/api/documents/process', {
  method: 'POST',
  body: formData
});

const result = await response.json();
console.log('Output ID:', result.outputId);

// 根据ID下载
window.open(`/api/documents/download/id/${result.outputId}`);
```

## 数据迁移

### 从文件系统迁移到数据库

如果需要将现有的文件系统存储迁移到数据库：

1. 使用 `listOutputDocuments()` 获取所有文档
2. 对于每个文档，读取文件内容
3. 创建 `DocumentEntity` 并保存到数据库

### 从数据库导出到文件系统

如果需要将数据库中的文档导出到文件系统：

1. 使用 `documentRepository.findAll()` 获取所有文档
2. 对于每个文档，读取 `content` 字段
3. 写入到文件系统

## 性能优化建议

1. **大文件处理**：
   - 对于大于 10MB 的文件，建议使用 `local` 或 `s3` 存储类型
   - 数据库存储适合小于 10MB 的文件

2. **索引优化**：
   - `output_id` 字段已建立唯一索引
   - `created_at` 字段已建立索引，查询列表时性能更好

3. **清理策略**：
   - 定期使用 `cleanupOldFiles()` 清理旧文档
   - 建议保留最近 30 天的文档

## 注意事项

1. **数据库备份**：定期备份数据库文件（H2）或使用数据库备份工具（MySQL/PostgreSQL）

2. **存储空间**：使用数据库存储时，注意数据库文件大小，及时清理旧数据

3. **事务管理**：所有数据库操作都使用 `@Transactional` 注解，确保数据一致性

4. **并发访问**：H2 支持多连接访问（AUTO_SERVER=TRUE），但生产环境建议使用 MySQL 或 PostgreSQL

