# 需求分解和需求追溯功能使用指南

## 一、功能概述

本工具实现了两个核心功能：
1. **需求分解**：支持将高层需求分解为多个子需求，形成需求树结构
2. **需求追溯**：建立需求与测试用例之间的追溯关系，支持追溯矩阵生成和验证

### 1.1 输入输出说明

**输入**：
- **需求数据**：可以通过代码创建或从Excel文件读取
  - 需求基本信息：ID、名称、描述、类型、优先级等
  - 需求层级关系：父需求ID（用于建立需求树）
- **测试用例数据**：可以通过代码创建或从Excel文件读取
  - 测试用例基本信息：ID、名称、内容等
  - 关联需求ID：用于建立追溯关系

**输出**：
- **需求树结构**：文本格式的树形结构，展示需求的层级关系
- **追溯矩阵**：需求ID到测试用例ID列表的映射关系
- **追溯验证报告**：列出未追溯的需求和测试用例
- **追溯覆盖率报告**：需求覆盖率和测试用例覆盖率的百分比

**详细输入输出示例请参考**：`example_input_output.md`

---

## 二、需求分解功能

### 2.1 核心类

- **`Requirement`**：需求数据模型
- **`RequirementManager`**：需求管理器，提供需求分解、查询、管理功能

### 2.2 基本使用

#### 创建需求

```java
RequirementManager manager = new RequirementManager("REQ");

// 创建根需求
Requirement rootReq = new Requirement("REQ-001", "用户登录功能");
rootReq.setDescription("实现用户登录功能");
rootReq.setType(Requirement.RequirementType.FUNCTIONAL);
rootReq.setPriority(Requirement.Priority.HIGH);
manager.addRequirement(rootReq);
```

#### 手动分解需求

```java
// 创建子需求
List<Requirement> children = new ArrayList<>();

Requirement child1 = new Requirement();
child1.setRequirementName("用户名密码验证");
child1.setDescription("验证用户输入的用户名和密码");
children.add(child1);

Requirement child2 = new Requirement();
child2.setRequirementName("登录状态管理");
child2.setDescription("管理用户的登录状态");
children.add(child2);

// 执行分解
List<Requirement> decomposed = manager.decomposeRequirement("REQ-001", children);
```

#### 自动分解需求

```java
// 根据名称列表自动创建子需求
List<String> childNames = Arrays.asList("记住密码功能", "自动登录功能");
List<Requirement> autoDecomposed = manager.autoDecomposeRequirement("REQ-001", childNames);
```

### 2.3 需求查询

```java
// 获取所有需求
List<Requirement> allReqs = manager.getAllRequirements();

// 获取根需求
List<Requirement> rootReqs = manager.getRootRequirements();

// 获取子需求
List<Requirement> children = manager.getChildren("REQ-001");

// 获取所有后代需求
List<Requirement> descendants = manager.getDescendants("REQ-001");

// 搜索需求
List<Requirement> results = manager.searchRequirements("登录");

// 获取叶子需求（没有子需求的需求）
List<Requirement> leafReqs = manager.getLeafRequirements();
```

### 2.4 需求树可视化

```java
// 获取需求树
RequirementManager.RequirementTree tree = manager.getRequirementTree();

// 打印需求树
tree.printTree();
```

输出示例：
```
REQ-001 用户登录功能
  REQ-001.1 用户名密码验证
  REQ-001.2 登录状态管理
  REQ-001.3 错误处理
    REQ-001.3.1 账户锁定处理
```

---

## 三、需求追溯功能

### 3.1 核心类

- **`Traceability`**：追溯关系数据模型
- **`TraceabilityManager`**：追溯管理器，提供追溯关系建立、查询、验证功能

### 3.2 基本使用

#### 建立追溯关系

```java
RequirementManager reqManager = new RequirementManager("REQ");
TraceabilityManager traceManager = new TraceabilityManager(reqManager);

// 建立单个追溯关系
Traceability trace = traceManager.establishTraceability("REQ-001", "TC-001");

// 批量建立追溯关系
Map<String, List<String>> requirementToTestCases = new HashMap<>();
requirementToTestCases.put("REQ-001", Arrays.asList("TC-001", "TC-002"));
requirementToTestCases.put("REQ-002", Arrays.asList("TC-003"));
List<Traceability> traces = traceManager.establishTraceabilities(requirementToTestCases);
```

#### 查询追溯关系

```java
// 根据需求ID获取所有追溯关系
List<Traceability> traces = traceManager.getTraceabilitiesByRequirement("REQ-001");

// 根据测试用例ID获取所有追溯关系
List<Traceability> traces = traceManager.getTraceabilitiesByTestCase("TC-001");

// 获取需求对应的所有测试用例ID
List<String> testCaseIds = traceManager.getTestCaseIdsByRequirement("REQ-001");

// 获取测试用例对应的所有需求ID
List<String> requirementIds = traceManager.getRequirementIdsByTestCase("TC-001");
```

### 3.3 追溯矩阵生成

```java
// 生成基本追溯矩阵（需求ID -> 测试用例ID列表）
Map<String, List<String>> matrix = traceManager.generateTraceMatrix();

// 生成详细追溯矩阵（需求编号 -> 测试用例名称列表）
Map<String, List<String>> detailedMatrix = traceManager.generateTraceMatrixWithDetails();

// 打印追溯矩阵
for (Map.Entry<String, List<String>> entry : matrix.entrySet()) {
    System.out.println("需求 " + entry.getKey() + " -> 测试用例 " + entry.getValue());
}
```

### 3.4 追溯关系验证

```java
// 验证追溯关系完整性
List<Requirement> requirements = ...;  // 所有需求
List<TestCase> testCases = ...;      // 所有测试用例

TraceabilityManager.TraceabilityValidationResult validation = 
    traceManager.validateTraceability(requirements, testCases);

// 检查验证结果
if (validation.isValid()) {
    System.out.println("所有追溯关系有效");
} else {
    // 打印验证报告
    validation.printReport();
}
```

验证报告会显示：
- 没有测试用例的需求列表
- 没有需求关联的测试用例列表
- 无效的追溯关系（需求或测试用例不存在）

### 3.5 追溯覆盖率计算

```java
// 计算追溯覆盖率
TraceabilityManager.TraceabilityCoverage coverage = 
    traceManager.calculateCoverage(requirements, testCases);

// 打印覆盖率报告
coverage.printReport();
```

覆盖率报告包括：
- 需求覆盖率：有测试用例的需求数量 / 总需求数量
- 测试用例覆盖率：有需求关联的测试用例数量 / 总测试用例数量

---

## 四、Excel集成

### 4.1 从Excel读取需求

#### Excel格式要求

| 列名 | 是否必填 | 说明 | 示例 |
|------|---------|------|------|
| 需求ID | 是 | 需求唯一标识 | REQ-001 |
| 需求名称 | 是 | 需求名称 | 用户登录功能 |
| 需求编号 | 否 | 需求编号（自动生成） | REQ-001 |
| 需求描述 | 否 | 需求详细描述 | 实现用户登录功能... |
| 需求类型 | 否 | 功能需求/性能需求等 | 功能需求 |
| 优先级 | 否 | 高/中/低 | 高 |
| 父需求ID | 否 | 父需求ID（用于建立层级关系） | REQ-001 |
| 状态 | 否 | 草稿/评审中/已批准等 | 已批准 |

#### 读取代码

```java
RequirementExcelReader reader = new RequirementExcelReader();
List<Requirement> requirements = reader.readRequirements("requirements.xlsx");

// 添加到需求管理器
RequirementManager manager = new RequirementManager("REQ");
for (Requirement req : requirements) {
    manager.addRequirement(req);
}

// 建立需求树关系（如果Excel中有父需求ID）
for (Requirement req : requirements) {
    if (req.getParentRequirementId() != null) {
        Requirement parent = manager.getRequirement(req.getParentRequirementId());
        if (parent != null) {
            parent.addChild(req);
        }
    }
}
```

### 4.2 从测试用例Excel建立追溯关系

如果测试用例Excel中包含"需求ID"或"关联需求"列，可以自动建立追溯关系：

```java
// 读取测试用例（使用现有的ExcelReader）
ExcelReader excelReader = new ExcelReader();
Map<String, ModuleData> moduleDataMap = excelReader.readExcel("test_cases.xlsx");

// 读取需求
RequirementManager reqManager = new RequirementManager("REQ");
// ... 添加需求 ...

// 建立追溯关系
TraceabilityManager traceManager = new TraceabilityManager(reqManager);

for (ModuleData moduleData : moduleDataMap.values()) {
    for (TestCase testCase : moduleData.getTestCases()) {
        // 从测试用例的"关联需求"或"需求ID"字段获取需求ID
        String requirementId = testCase.getColumnValue("需求ID");
        if (requirementId != null && !requirementId.trim().isEmpty()) {
            traceManager.establishTraceability(requirementId.trim(), testCase.getId());
        }
    }
}
```

---

## 五、完整示例

### 5.1 代码示例

参考 `RequirementTraceabilityExample.java` 文件，包含三个完整示例：
1. 需求分解示例
2. 需求追溯示例
3. Excel集成示例

运行示例：
```bash
java -cp target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar \
     pub.developers.docautogenbyexcel.example.RequirementTraceabilityExample
```

### 5.2 输入输出示例

**输入示例**：

```java
// 输入1：创建根需求
Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
reqManager.addRequirement(rootReq);

// 输入2：分解需求（输入子需求名称列表）
List<String> childNames = Arrays.asList("用户登录功能", "用户注册功能");
List<Requirement> children = reqManager.autoDecomposeRequirement("REQ-001", childNames);

// 输入3：创建测试用例并关联需求
TestCase tc1 = new TestCase();
tc1.setId("TC-001");
tc1.setTestName("用户登录功能测试");
tc1.addColumnData("需求ID", "REQ-001.1");  // 关联到需求REQ-001.1

// 输入4：建立追溯关系
traceManager.establishTraceability("REQ-001.1", "TC-001");
```

**输出示例**：

```
需求树结构：
REQ-001 用户管理系统
  REQ-001.1 用户登录功能
  REQ-001.2 用户注册功能

追溯矩阵：
  REQ-001.1 -> [TC-001]

追溯覆盖率：
  需求覆盖率: 1/3 (33.33%)
  测试用例覆盖率: 1/1 (100.00%)
```

**详细输入输出示例请参考**：`example_input_output.md`

---

## 六、最佳实践

### 6.1 需求分解

1. **分层分解**：建议不超过3-4层，避免需求树过深
2. **粒度控制**：子需求应该是可测试、可实现的独立单元
3. **编号规范**：使用统一的前缀和编号规则（如REQ-001.1.1）

### 6.2 需求追溯

1. **全覆盖**：确保每个需求都有对应的测试用例
2. **及时更新**：需求变更时及时更新追溯关系
3. **定期验证**：定期运行验证功能，检查追溯关系完整性

### 6.3 Excel数据管理

1. **数据规范**：使用统一的列名和格式
2. **关系维护**：通过"父需求ID"列维护需求层级关系
3. **版本控制**：对Excel文件进行版本管理

---

## 七、常见问题

### Q1: 如何修改需求编号前缀？

```java
RequirementManager manager = new RequirementManager("REQ");  // 使用REQ作为前缀
// 或者
RequirementManager manager = new RequirementManager("REQ-SYS");  // 使用REQ-SYS作为前缀
```

### Q2: 如何删除需求及其所有子需求？

```java
manager.removeRequirement("REQ-001");  // 会自动删除所有子需求
```

### Q3: 如何导出追溯矩阵到Excel？

目前需要手动实现，可以参考以下思路：
1. 使用Apache POI创建Excel文件
2. 将追溯矩阵数据写入Excel
3. 格式化表格（添加边框、样式等）

### Q4: 如何支持多对多追溯关系？

当前实现已经支持：
- 一个需求可以对应多个测试用例
- 一个测试用例可以对应多个需求

只需多次调用 `establishTraceability()` 即可。

---

## 八、需求树Word文档生成

### 8.1 生成基础需求树Word文档

将需求树以树状结构输出到Word文档：

```java
RequirementManager reqManager = new RequirementManager("REQ");
// ... 创建需求并分解 ...

RequirementTreeWordGenerator generator = new RequirementTreeWordGenerator();
generator.generateRequirementTreeWord(reqManager, "output/requirement_tree.docx");
```

**输出内容**：
- 需求树结构（树状格式，带缩进）
- 需求编号、名称、类型、优先级
- 需求描述（斜体、灰色）
- 统计信息（总需求数、根需求数、叶子需求数等）

### 8.2 生成带追溯信息的需求树Word文档

在需求树中显示每个需求对应的测试用例：

```java
RequirementManager reqManager = new RequirementManager("REQ");
TraceabilityManager traceManager = new TraceabilityManager(reqManager);
// ... 建立追溯关系 ...

RequirementTreeWordGenerator generator = new RequirementTreeWordGenerator();
generator.generateRequirementTreeWordWithTraceability(
    reqManager, traceManager, "output/requirement_tree_with_trace.docx");
```

**输出内容**：
- 基础需求树结构
- 每个需求对应的测试用例（蓝色显示）
- 追溯统计信息（已追溯需求数、追溯覆盖率）

### 8.3 Word文档格式说明

**树状结构**：
```
REQ-001 用户管理系统 [功能需求] (高)
    REQ-001.1 用户登录功能 [功能需求] (高)
        → 测试用例: TC-001
        REQ-001.1.1 用户名密码验证 [功能需求] (高)
            → 测试用例: TC-002
        REQ-001.1.2 登录状态管理 [功能需求] (高)
    REQ-001.2 用户注册功能 [功能需求] (高)
        → 测试用例: TC-003
    REQ-001.3 用户信息管理 [功能需求] (中)
```

**格式特点**：
- 使用缩进表示层级关系
- 使用树状符号（├─、└─）表示节点类型
- 需求描述以斜体、灰色显示
- 追溯信息以蓝色显示
- 根节点和重要节点加粗显示

### 8.4 完整示例

参考 `RequirementTreeWordExample.java` 文件，包含完整的使用示例。

运行示例：
```bash
mvn exec:java -Dexec.mainClass="pub.developers.docautogenbyexcel.example.RequirementTreeWordExample"
```

---

## 九、扩展开发

### 9.1 添加自定义属性

`Requirement` 类支持扩展属性：

```java
requirement.addAttribute("负责人", "张三");
requirement.addAttribute("预计工时", "8小时");
String owner = requirement.getAttribute("负责人");
```

### 9.2 自定义需求类型

在 `Requirement.RequirementType` 枚举中添加新类型。

### 9.3 自定义Word文档格式

可以修改 `RequirementTreeWordGenerator` 类来自定义：
- 树状符号样式
- 字体大小和颜色
- 缩进距离
- 统计信息格式

---

## 九、输入输出说明

### 9.1 输入说明

#### 输入类型1：代码方式

**需求输入**：
```java
// 创建根需求
Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
rootReq.setDescription("实现完整的用户管理功能");
rootReq.setType(Requirement.RequirementType.FUNCTIONAL);
rootReq.setPriority(Requirement.Priority.HIGH);

// 分解需求（输入子需求名称列表）
List<String> childNames = Arrays.asList(
    "用户登录功能",
    "用户注册功能",
    "用户信息管理"
);
```

**测试用例输入**：
```java
TestCase tc1 = new TestCase();
tc1.setId("TC-001");
tc1.setTestName("用户登录功能测试");
tc1.addColumnData("需求ID", "REQ-001.1");  // 关联需求编号
```

#### 输入类型2：Excel文件方式

**需求Excel文件（requirements.xlsx）**：

| 需求ID | 需求编号 | 需求名称 | 需求描述 | 需求类型 | 优先级 | 父需求ID | 状态 |
|--------|---------|---------|---------|---------|--------|---------|------|
| REQ-001 | REQ-001 | 用户管理系统 | 实现完整的用户管理功能 | 功能需求 | 高 | | 已批准 |
| REQ-002 | REQ-001.1 | 用户登录功能 | 实现用户登录功能 | 功能需求 | 高 | REQ-001 | 已批准 |
| REQ-003 | REQ-001.2 | 用户注册功能 | 实现用户注册功能 | 功能需求 | 高 | REQ-001 | 已批准 |

**测试用例Excel文件（test_cases.xlsx）**：

| 模块编号 | 测试项名称 | 标识 | 需求ID | 测试内容 | 测试策略与方法 |
|---------|-----------|------|--------|---------|--------------|
| 5.3 | 用户登录功能测试 | TC-001 | REQ-001.1 | 验证用户登录功能 | 1) 输入正确用户名密码<br>2) 点击登录按钮 |
| 5.3 | 用户注册功能测试 | TC-002 | REQ-001.2 | 验证用户注册功能 | 1) 输入注册信息<br>2) 提交注册 |

### 9.2 输出说明

#### 输出1：需求树结构

```
REQ-001 用户管理系统
  REQ-001.1 用户登录功能
    REQ-001.1.1 用户名密码验证
    REQ-001.1.2 登录状态管理
  REQ-001.2 用户注册功能
  REQ-001.3 用户信息管理
```

**说明**：以文本树形结构展示需求的层级关系，每个需求显示编号和名称。

#### 输出2：追溯矩阵

```
追溯矩阵：
  REQ-001.1 用户登录功能
      -> 测试用例: [TC-001]
  REQ-001.1.1 用户名密码验证
      -> 测试用例: [TC-002]
  REQ-001.2 用户注册功能
      -> 测试用例: [TC-003]
```

**数据结构**：
```java
Map<String, List<String>> matrix = {
    "REQ-001.1" -> ["TC-001"],
    "REQ-001.1.1" -> ["TC-002"],
    "REQ-001.2" -> ["TC-003"]
}
```

#### 输出3：追溯关系验证报告

```
=== 追溯关系验证报告 ===
未追溯的需求数量: 3
  - REQ-006
  - REQ-004
  - REQ-001
未追溯的测试用例数量: 0
无效的追溯关系数量: 0
```

**说明**：
- 列出所有没有关联测试用例的需求
- 列出所有没有关联需求的测试用例
- 列出所有无效的追溯关系（需求或测试用例不存在）

#### 输出4：追溯覆盖率报告

```
=== 追溯覆盖率报告 ===
需求覆盖率: 3/6 (50.00%)
测试用例覆盖率: 3/3 (100.00%)
```

**说明**：
- **需求覆盖率**：有测试用例的需求数量 / 总需求数量
- **测试用例覆盖率**：有需求关联的测试用例数量 / 总测试用例数量

### 9.3 完整输入输出示例

**输入**：
1. 根需求：REQ-001 "用户管理系统"
2. 子需求：["用户登录功能", "用户注册功能", "用户信息管理"]
3. 测试用例：TC-001关联REQ-001.1, TC-002关联REQ-001.2

**处理**：
```java
// 创建需求并分解
RequirementManager reqManager = new RequirementManager("REQ");
Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
reqManager.addRequirement(rootReq);
List<String> childNames = Arrays.asList("用户登录功能", "用户注册功能", "用户信息管理");
reqManager.autoDecomposeRequirement("REQ-001", childNames);

// 建立追溯关系
TraceabilityManager traceManager = new TraceabilityManager(reqManager);
traceManager.establishTraceability("REQ-001.1", "TC-001");
traceManager.establishTraceability("REQ-001.2", "TC-002");
```

**输出**：
```
需求树结构：
REQ-001 用户管理系统
  REQ-001.1 用户登录功能
  REQ-001.2 用户注册功能
  REQ-001.3 用户信息管理

追溯矩阵：
  REQ-001.1 -> [TC-001]
  REQ-001.2 -> [TC-002]

追溯覆盖率：
  需求覆盖率: 2/4 (50.00%)
  测试用例覆盖率: 2/2 (100.00%)
```

**详细输入输出示例请参考**：`example_input_output.md`

---

*文档版本：v1.1*  
*最后更新：2025-12-01*

