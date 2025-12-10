# 需求分解和需求追溯功能 - 输入输出示例

## 一、功能概述

**输入**：需求数据（手动创建或从Excel读取）、测试用例数据  
**输出**：需求树结构、追溯矩阵、追溯验证报告、追溯覆盖率报告

---

## 二、输入示例

### 2.1 方式一：代码方式输入

#### 输入代码

```java
// 1. 创建需求管理器
RequirementManager reqManager = new RequirementManager("REQ");

// 2. 创建根需求（输入）
Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
rootReq.setDescription("实现完整的用户管理功能");
rootReq.setType(Requirement.RequirementType.FUNCTIONAL);
rootReq.setPriority(Requirement.Priority.HIGH);
reqManager.addRequirement(rootReq);

// 3. 分解需求（输入：子需求名称列表）
List<String> childNames = Arrays.asList(
    "用户登录功能",
    "用户注册功能",
    "用户信息管理"
);
List<Requirement> children = reqManager.autoDecomposeRequirement("REQ-001", childNames);

// 4. 创建测试用例（输入）
TestCase tc1 = new TestCase();
tc1.setId("TC-001");
tc1.setTestName("用户登录功能测试");
tc1.addColumnData("需求ID", "REQ-001.1");  // 关联到需求REQ-001.1

TestCase tc2 = new TestCase();
tc2.setId("TC-002");
tc2.setTestName("用户注册功能测试");
tc2.addColumnData("需求ID", "REQ-001.2");  // 关联到需求REQ-001.2

// 5. 建立追溯关系（输入：需求ID和测试用例ID）
TraceabilityManager traceManager = new TraceabilityManager(reqManager);
traceManager.establishTraceability("REQ-001.1", "TC-001");
traceManager.establishTraceability("REQ-001.2", "TC-002");
```

#### 输入数据结构

**需求输入**：
- 需求ID：REQ-001
- 需求名称：用户管理系统
- 需求描述：实现完整的用户管理功能
- 需求类型：功能需求
- 优先级：高
- 子需求名称列表：["用户登录功能", "用户注册功能", "用户信息管理"]

**测试用例输入**：
- 测试用例ID：TC-001, TC-002
- 测试用例名称：用户登录功能测试, 用户注册功能测试
- 关联需求ID：REQ-001.1, REQ-001.2

---

### 2.2 方式二：Excel文件输入

#### Excel文件格式（requirements.xlsx）

| 需求ID | 需求编号 | 需求名称 | 需求描述 | 需求类型 | 优先级 | 父需求ID | 状态 |
|--------|---------|---------|---------|---------|--------|---------|------|
| REQ-001 | REQ-001 | 用户管理系统 | 实现完整的用户管理功能 | 功能需求 | 高 | | 已批准 |
| REQ-002 | REQ-001.1 | 用户登录功能 | 实现用户登录功能 | 功能需求 | 高 | REQ-001 | 已批准 |
| REQ-003 | REQ-001.2 | 用户注册功能 | 实现用户注册功能 | 功能需求 | 高 | REQ-001 | 已批准 |
| REQ-004 | REQ-001.3 | 用户信息管理 | 实现用户信息管理功能 | 功能需求 | 中 | REQ-001 | 已批准 |

#### Excel文件格式（test_cases.xlsx）

| 模块编号 | 测试项名称 | 标识 | 需求ID | 测试内容 | 测试策略与方法 | 判定准则 |
|---------|-----------|------|--------|---------|--------------|---------|
| 5.3 | 用户登录功能测试 | TC-001 | REQ-001.1 | 验证用户登录功能 | 1) 输入正确用户名密码<br>2) 点击登录按钮 | 1) 登录成功<br>2) 跳转到首页 |
| 5.3 | 用户注册功能测试 | TC-002 | REQ-001.2 | 验证用户注册功能 | 1) 输入注册信息<br>2) 提交注册 | 1) 注册成功<br>2) 发送验证邮件 |

#### 读取Excel的代码

```java
// 读取需求Excel
RequirementExcelReader reqReader = new RequirementExcelReader();
List<Requirement> requirements = reqReader.readRequirements("requirements.xlsx");

// 读取测试用例Excel
ExcelReader testReader = new ExcelReader();
Map<String, ModuleData> moduleDataMap = testReader.readExcel("test_cases.xlsx");

// 添加到管理器
RequirementManager reqManager = new RequirementManager("REQ");
for (Requirement req : requirements) {
    reqManager.addRequirement(req);
}

// 建立追溯关系
TraceabilityManager traceManager = new TraceabilityManager(reqManager);
for (ModuleData moduleData : moduleDataMap.values()) {
    for (TestCase testCase : moduleData.getTestCases()) {
        String requirementId = testCase.getColumnValue("需求ID");
        if (requirementId != null && !requirementId.trim().isEmpty()) {
            // 根据需求编号查找需求
            Requirement req = reqManager.findRequirementByNumber(requirementId.trim());
            if (req != null) {
                traceManager.establishTraceability(req.getRequirementId(), testCase.getId());
            }
        }
    }
}
```

---

## 三、输出示例

### 3.1 需求树结构输出

```
REQ-001 用户管理系统
  REQ-001.1 用户登录功能
    REQ-001.1.1 用户名密码验证
    REQ-001.1.2 登录状态管理
  REQ-001.2 用户注册功能
  REQ-001.3 用户信息管理
```

**输出数据结构**：
- 树形结构的需求层级关系
- 每个需求包含：编号、名称、层级深度

---

### 3.2 追溯矩阵输出

```
追溯矩阵：
  REQ-001.1 用户登录功能
      -> 测试用例: [TC-001]
  REQ-001.1.1 用户名密码验证
      -> 测试用例: [TC-002]
  REQ-001.2 用户注册功能
      -> 测试用例: [TC-003]
```

**输出数据结构**：
```java
Map<String, List<String>> matrix = {
    "REQ-001.1" -> ["TC-001"],
    "REQ-001.1.1" -> ["TC-002"],
    "REQ-001.2" -> ["TC-003"]
}
```

---

### 3.3 追溯关系验证报告输出

```
=== 追溯关系验证报告 ===
未追溯的需求数量: 3
  - REQ-006
  - REQ-004
  - REQ-001
未追溯的测试用例数量: 0
无效的追溯关系数量: 0
```

**输出数据结构**：
- 未追溯的需求列表
- 未追溯的测试用例列表
- 无效的追溯关系列表（需求或测试用例不存在）

---

### 3.4 追溯覆盖率报告输出

```
=== 追溯覆盖率报告 ===
需求覆盖率: 3/6 (50.00%)
测试用例覆盖率: 3/3 (100.00%)
```

**输出数据结构**：
- 总需求数量：6
- 已追溯需求数量：3
- 需求覆盖率：50.00%
- 总测试用例数量：3
- 已追溯测试用例数量：3
- 测试用例覆盖率：100.00%

---

## 四、完整示例

### 4.1 输入

**需求数据**：
1. 根需求：REQ-001 "用户管理系统"
2. 子需求列表：["用户登录功能", "用户注册功能", "用户信息管理"]
3. 孙需求列表（针对"用户登录功能"）：["用户名密码验证", "登录状态管理"]

**测试用例数据**：
1. TC-001 "用户登录功能测试" -> 关联 REQ-001.1
2. TC-002 "用户名密码验证测试" -> 关联 REQ-001.1.1
3. TC-003 "用户注册功能测试" -> 关联 REQ-001.2

### 4.2 处理过程

```java
// 1. 创建需求并分解
RequirementManager reqManager = new RequirementManager("REQ");
Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
reqManager.addRequirement(rootReq);

List<String> childNames = Arrays.asList("用户登录功能", "用户注册功能", "用户信息管理");
List<Requirement> children = reqManager.autoDecomposeRequirement("REQ-001", childNames);

Requirement loginReq = children.get(0);
List<String> loginChildren = Arrays.asList("用户名密码验证", "登录状态管理");
reqManager.autoDecomposeRequirement(loginReq.getRequirementId(), loginChildren);

// 2. 创建测试用例
List<TestCase> testCases = Arrays.asList(
    createTestCase("TC-001", "用户登录功能测试", "REQ-001.1"),
    createTestCase("TC-002", "用户名密码验证测试", "REQ-001.1.1"),
    createTestCase("TC-003", "用户注册功能测试", "REQ-001.2")
);

// 3. 建立追溯关系
TraceabilityManager traceManager = new TraceabilityManager(reqManager);
for (TestCase tc : testCases) {
    String reqNumber = tc.getColumnValue("需求ID");
    Requirement req = reqManager.findRequirementByNumber(reqNumber);
    if (req != null) {
        traceManager.establishTraceability(req.getRequirementId(), tc.getId());
    }
}
```

### 4.3 输出

**1. 需求树结构**：
```
REQ-001 用户管理系统
  REQ-001.1 用户登录功能
    REQ-001.1.1 用户名密码验证
    REQ-001.1.2 登录状态管理
  REQ-001.2 用户注册功能
  REQ-001.3 用户信息管理
```

**2. 追溯矩阵**：
```
REQ-001.1 -> [TC-001]
REQ-001.1.1 -> [TC-002]
REQ-001.2 -> [TC-003]
```

**3. 追溯验证报告**：
```
未追溯的需求数量: 3
  - REQ-006 (登录状态管理)
  - REQ-004 (用户信息管理)
  - REQ-001 (用户管理系统)
未追溯的测试用例数量: 0
无效的追溯关系数量: 0
```

**4. 追溯覆盖率报告**：
```
需求覆盖率: 3/6 (50.00%)
测试用例覆盖率: 3/3 (100.00%)
```

---

## 五、输入输出总结

### 5.1 输入类型

| 输入类型 | 格式 | 说明 |
|---------|------|------|
| 需求数据 | Java对象或Excel文件 | 需求的基本信息和层级关系 |
| 测试用例数据 | Java对象或Excel文件 | 测试用例信息和关联的需求ID |
| 追溯关系 | 需求ID + 测试用例ID | 建立需求与测试用例的关联 |

### 5.2 输出类型

| 输出类型 | 格式 | 说明 |
|---------|------|------|
| 需求树 | 文本树形结构 | 展示需求的层级关系 |
| 追溯矩阵 | Map<String, List<String>> | 需求到测试用例的映射关系 |
| 验证报告 | 文本报告 | 未追溯的需求和测试用例列表 |
| 覆盖率报告 | 百分比数据 | 需求覆盖率和测试用例覆盖率 |

---

*文档版本：v1.0*  
*最后更新：2025-12-01*

