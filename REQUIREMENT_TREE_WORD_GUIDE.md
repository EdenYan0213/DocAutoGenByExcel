# 需求树Word文档生成使用指南

## 一、功能概述

本功能可以将需求树以树状结构输出到Word文档，支持：
- ✅ 基础需求树结构（树状格式，带缩进）
- ✅ 带追溯信息的需求树（显示每个需求对应的测试用例）
- ✅ 统计信息（总需求数、根需求数、叶子需求数、最大层级深度等）
- ✅ 追溯统计（已追溯需求数、追溯覆盖率）

---

## 二、快速开始

### 2.1 生成基础需求树Word文档

```java
// 1. 创建需求管理器并构建需求树
RequirementManager reqManager = new RequirementManager("REQ");
Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
reqManager.addRequirement(rootReq);

List<String> childNames = Arrays.asList("用户登录功能", "用户注册功能");
reqManager.autoDecomposeRequirement("REQ-001", childNames);

// 2. 生成Word文档
RequirementTreeWordGenerator generator = new RequirementTreeWordGenerator();
generator.generateRequirementTreeWord(reqManager, "output/requirement_tree.docx");
```

### 2.2 生成带追溯信息的需求树Word文档

```java
// 1. 创建需求管理器和追溯管理器
RequirementManager reqManager = new RequirementManager("REQ");
TraceabilityManager traceManager = new TraceabilityManager(reqManager);

// 2. 构建需求树并建立追溯关系
// ... 创建需求和测试用例 ...
traceManager.establishTraceability("REQ-001.1", "TC-001");

// 3. 生成Word文档
RequirementTreeWordGenerator generator = new RequirementTreeWordGenerator();
generator.generateRequirementTreeWordWithTraceability(
    reqManager, traceManager, "output/requirement_tree_with_trace.docx");
```

---

## 三、Word文档格式说明

### 3.1 树状结构格式

**示例输出**：
```
REQ-001 用户管理系统 [功能需求] (高)
    实现完整的用户管理功能，包括登录、注册、信息管理
├─ REQ-001.1 用户登录功能 [功能需求] (高)
    实现用户登录功能，支持用户名密码登录
    → 测试用例: TC-001
    ├─ REQ-001.1.1 用户名密码验证 [功能需求] (高)
        验证用户输入的用户名和密码是否正确
        → 测试用例: TC-002
    └─ REQ-001.1.2 登录状态管理 [功能需求] (高)
        管理用户的登录状态，包括登录、登出
├─ REQ-001.2 用户注册功能 [功能需求] (高)
    实现用户注册功能，支持邮箱注册
    → 测试用例: TC-003
└─ REQ-001.3 用户信息管理 [功能需求] (中)
    实现用户信息管理功能，包括查看和修改
```

### 3.2 格式特点

| 元素 | 格式说明 |
|------|---------|
| 根需求 | 加粗、12号字体 |
| 子需求 | 普通字体、带缩进 |
| 树状符号 | ├─（有子节点）、└─（叶子节点） |
| 需求描述 | 斜体、灰色、10号字体 |
| 追溯信息 | 蓝色、10号字体 |
| 统计信息 | 11号字体、左缩进 |

### 3.3 统计信息

文档末尾包含统计信息：
- 总需求数
- 根需求数
- 叶子需求数
- 非叶子需求数
- 最大层级深度
- 追溯统计（如果提供追溯管理器）

---

## 四、完整示例

### 4.1 代码示例

```java
package pub.developers.docautogenbyexcel.example;

import pub.developers.docautogenbyexcel.generator.RequirementTreeWordGenerator;
import pub.developers.docautogenbyexcel.manager.RequirementManager;
import pub.developers.docautogenbyexcel.manager.TraceabilityManager;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;

import java.util.Arrays;
import java.util.List;

public class RequirementTreeWordExample {
    public static void main(String[] args) {
        try {
            // 1. 创建需求管理器
            RequirementManager reqManager = new RequirementManager("REQ");
            
            // 2. 构建需求树
            Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
            rootReq.setDescription("实现完整的用户管理功能");
            reqManager.addRequirement(rootReq);
            
            List<String> childNames = Arrays.asList("用户登录功能", "用户注册功能");
            reqManager.autoDecomposeRequirement("REQ-001", childNames);
            
            // 3. 生成基础需求树Word文档
            RequirementTreeWordGenerator generator = new RequirementTreeWordGenerator();
            generator.generateRequirementTreeWord(reqManager, "output/requirement_tree.docx");
            
            // 4. 建立追溯关系并生成带追溯信息的Word文档
            TraceabilityManager traceManager = new TraceabilityManager(reqManager);
            traceManager.establishTraceability("REQ-001.1", "TC-001");
            
            generator.generateRequirementTreeWordWithTraceability(
                reqManager, traceManager, "output/requirement_tree_with_trace.docx");
            
            System.out.println("Word文档生成成功！");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
```

### 4.2 运行示例

```bash
# 编译项目
mvn clean package -DskipTests

# 运行示例
mvn exec:java -Dexec.mainClass="pub.developers.docautogenbyexcel.example.RequirementTreeWordExample"
```

---

## 五、API说明

### 5.1 RequirementTreeWordGenerator类

#### 方法1：generateRequirementTreeWord()

**功能**：生成基础需求树Word文档

**参数**：
- `requirementManager` - 需求管理器
- `outputPath` - 输出文件路径

**示例**：
```java
RequirementTreeWordGenerator generator = new RequirementTreeWordGenerator();
generator.generateRequirementTreeWord(reqManager, "output/requirement_tree.docx");
```

#### 方法2：generateRequirementTreeWordWithTraceability()

**功能**：生成带追溯信息的需求树Word文档

**参数**：
- `requirementManager` - 需求管理器
- `traceabilityManager` - 追溯管理器（可选，传null则不显示追溯信息）
- `outputPath` - 输出文件路径

**示例**：
```java
RequirementTreeWordGenerator generator = new RequirementTreeWordGenerator();
generator.generateRequirementTreeWordWithTraceability(
    reqManager, traceManager, "output/requirement_tree_with_trace.docx");
```

---

## 六、输出文档内容

### 6.1 基础需求树文档内容

1. **标题**：需求树结构（居中、加粗、18号字体）
2. **需求树**：
   - 树状结构展示
   - 需求编号、名称
   - 需求类型、优先级
   - 需求描述（如果有）
3. **统计信息**：
   - 总需求数
   - 根需求数
   - 叶子需求数
   - 非叶子需求数
   - 最大层级深度

### 6.2 带追溯信息的需求树文档内容

在基础需求树的基础上，增加：
1. **追溯信息**：每个需求下方显示对应的测试用例ID
2. **追溯统计**：
   - 已追溯需求数
   - 追溯覆盖率

---

## 七、自定义格式

### 7.1 修改树状符号

在 `RequirementTreeWordGenerator` 类中修改：
```java
// 当前使用：├─、└─
// 可以改为：│  ├─、│  └─ 或其他符号
```

### 7.2 修改缩进距离

```java
// 当前：每级缩进400 twips（约0.2英寸）
ind.setLeft(BigInteger.valueOf(level * 400));

// 可以调整为其他值，如：
ind.setLeft(BigInteger.valueOf(level * 600)); // 更大的缩进
```

### 7.3 修改字体和颜色

```java
// 修改字体大小
run.setFontSize(12); // 改为其他大小

// 修改颜色
run.setColor("0066CC"); // 改为其他颜色代码
```

---

## 八、常见问题

### Q1: 生成的Word文档打不开？

**A**: 确保：
- 使用Word 2016及以上版本
- 文件没有被其他程序占用
- 文件路径正确且有写入权限

### Q2: 树状结构显示不正确？

**A**: 
- 检查需求树是否正确构建
- 确认需求之间的父子关系是否正确
- 查看控制台输出的需求树结构

### Q3: 如何添加更多信息到Word文档？

**A**: 修改 `RequirementTreeWordGenerator` 类，在 `addRequirementNode()` 方法中添加更多内容。

---

## 九、完整示例文件

参考 `src/main/java/pub/developers/docautogenbyexcel/example/RequirementTreeWordExample.java`

---

*文档版本：v1.0*  
*最后更新：2025-12-01*

