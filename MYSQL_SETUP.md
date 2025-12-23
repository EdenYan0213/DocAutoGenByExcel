# MySQL 数据库配置指南

## 快速开始

### 1. 安装 MySQL

确保已安装 MySQL 8.0 或更高版本。

**macOS (使用 Homebrew)**:
```bash
brew install mysql
brew services start mysql
```

**Linux (Ubuntu/Debian)**:
```bash
sudo apt-get update
sudo apt-get install mysql-server
sudo systemctl start mysql
```

**Windows**:
从 [MySQL官网](https://dev.mysql.com/downloads/mysql/) 下载安装程序。

### 2. 创建数据库

登录 MySQL：
```bash
mysql -u root -p
```

创建数据库和用户：
```sql
-- 创建数据库
CREATE DATABASE docautogen CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

-- 创建用户（可选，也可以直接使用root）
CREATE USER 'docautogen'@'localhost' IDENTIFIED BY 'your-password';

-- 授予权限
GRANT ALL PRIVILEGES ON docautogen.* TO 'docautogen'@'localhost';
FLUSH PRIVILEGES;

-- 退出
EXIT;
```

### 3. 配置应用

编辑 `src/main/resources/application.properties`：

```properties
# MySQL Configuration
spring.datasource.url=jdbc:mysql://localhost:3306/docautogen?useUnicode=true&characterEncoding=utf8&useSSL=false&serverTimezone=Asia/Shanghai
spring.datasource.driver-class-name=com.mysql.cj.jdbc.Driver
spring.datasource.username=docautogen
spring.datasource.password=your-password
spring.jpa.database-platform=org.hibernate.dialect.MySQLDialect
```

**重要**：请将 `your-password` 替换为实际的数据库密码。

### 4. 启动应用

```bash
mvn clean package
java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar --spring
```

应用启动时会自动创建表结构（`spring.jpa.hibernate.ddl-auto=update`）。

## 验证配置

### 检查数据库连接

启动应用后，查看日志，应该看到类似以下信息：
```
HikariPool-1 - Starting...
HikariPool-1 - Start completed.
```

### 查看数据库表

登录 MySQL：
```bash
mysql -u docautogen -p docautogen
```

查看表：
```sql
SHOW TABLES;
-- 应该看到 documents 表

DESCRIBE documents;
-- 查看表结构
```

## 常见问题

### 1. 连接失败：Access denied

**原因**：用户名或密码错误

**解决**：
1. 检查 `application.properties` 中的用户名和密码
2. 确认 MySQL 用户有访问数据库的权限

### 2. 连接失败：Unknown database

**原因**：数据库不存在

**解决**：
```sql
CREATE DATABASE docautogen CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

### 3. 时区错误

**原因**：MySQL 时区设置不正确

**解决**：
在连接URL中添加时区参数（已包含）：
```
serverTimezone=Asia/Shanghai
```

或者设置 MySQL 时区：
```sql
SET GLOBAL time_zone = '+8:00';
```

### 4. 字符编码问题

**原因**：数据库字符集不是 utf8mb4

**解决**：
确保数据库使用 utf8mb4：
```sql
ALTER DATABASE docautogen CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

### 5. 表不存在

**原因**：JPA 自动创建表失败

**解决**：
1. 检查 `spring.jpa.hibernate.ddl-auto=update` 配置
2. 查看应用启动日志中的错误信息
3. 手动创建表（参考 `DATABASE.md` 中的表结构）

## 性能优化

### 1. 连接池配置

在 `application.properties` 中添加：
```properties
# 连接池配置
spring.datasource.hikari.maximum-pool-size=10
spring.datasource.hikari.minimum-idle=5
spring.datasource.hikari.connection-timeout=30000
spring.datasource.hikari.idle-timeout=600000
spring.datasource.hikari.max-lifetime=1800000
```

### 2. 索引优化

表结构已包含以下索引：
- `idx_created_at`：按创建时间查询
- `idx_output_id`：按输出ID查询（唯一索引）

### 3. BLOB 存储优化

对于大文件（>10MB），建议：
- 使用 `storage.type=local` 或 `storage.type=s3`
- 数据库只保存元数据，不保存文件内容

## 备份和恢复

### 备份数据库

```bash
mysqldump -u docautogen -p docautogen > docautogen_backup.sql
```

### 恢复数据库

```bash
mysql -u docautogen -p docautogen < docautogen_backup.sql
```

## 生产环境建议

1. **使用专用数据库用户**：不要使用 root 用户
2. **设置强密码**：使用复杂的密码
3. **启用 SSL**：在生产环境中启用 SSL 连接
4. **定期备份**：设置自动备份任务
5. **监控性能**：监控数据库连接数和查询性能
6. **关闭自动创建表**：生产环境使用 `spring.jpa.hibernate.ddl-auto=validate`

## 从 H2 迁移到 MySQL

如果之前使用 H2 数据库，需要迁移数据：

1. **导出 H2 数据**：
   - 使用 H2 控制台导出数据为 SQL 文件

2. **转换 SQL 语句**：
   - 将 H2 的 SQL 语法转换为 MySQL 语法
   - 注意数据类型差异

3. **导入到 MySQL**：
   ```bash
   mysql -u docautogen -p docautogen < converted_data.sql
   ```

或者使用应用的数据迁移功能（如果已实现）。

