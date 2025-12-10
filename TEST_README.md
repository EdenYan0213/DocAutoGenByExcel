# 需求分解和需求追溯功能测试说明

## 一、测试概述

本测试使用SpringBootTest框架，所有测试方法集中在一个测试类中，每个测试方法可以独立运行。

**测试类**：`src/test/java/pub/developers/docautogenbyexcel/RequirementTraceabilityTest.java`

## 快速参考

### 运行所有测试
```bash
mvn test
```

### 运行单个测试方法
```bash
# 测试1：代码方式输入输出
mvn test -Dtest="RequirementTraceabilityTest#testCodeInputOutput"

# 测试2：Excel方式输入输出
mvn test -Dtest="RequirementTraceabilityTest#testExcelInputOutput"

# 测试3：需求分解功能
mvn test -Dtest="RequirementTraceabilityTest#testRequirementDecomposition"

# 测试4：需求追溯功能
mvn test -Dtest="RequirementTraceabilityTest#testRequirementTraceability"

# 测试5：边界情况
mvn test -Dtest="RequirementTraceabilityTest#testEdgeCases"
```

### IDE中运行
- **IntelliJ IDEA**：点击测试方法左侧的绿色运行按钮
- **Eclipse**：右键点击方法名 -> "Run As" -> "JUnit Test"

## 二、测试方法列表

| 测试方法 | 说明 | 测试内容 |
|---------|------|---------|
| `testCodeInputOutput()` | 代码方式输入输出测试 | 测试通过代码创建需求、分解需求、建立追溯关系的完整流程 |
| `testExcelInputOutput()` | Excel方式输入输出测试 | 测试从Excel读取需求和数据，建立追溯关系的完整流程 |
| `testRequirementDecomposition()` | 需求分解功能测试 | 测试手动分解、自动分解、多级分解功能 |
| `testRequirementTraceability()` | 需求追溯功能测试 | 测试追溯关系建立、查询、验证、覆盖率计算 |
| `testEdgeCases()` | 边界情况测试 | 测试空列表、单个需求、搜索、删除等边界情况 |

## 三、运行测试

### 3.1 方式一：运行所有测试

```bash
# 运行所有测试
mvn test

# 或者只运行这个测试类
mvn test -Dtest=RequirementTraceabilityTest
```

### 3.2 方式二：运行单个测试方法

#### 使用Maven运行

```bash
# 运行测试1：代码方式输入输出测试
mvn test -Dtest="RequirementTraceabilityTest#testCodeInputOutput"

# 运行测试2：Excel方式输入输出测试
mvn test -Dtest="RequirementTraceabilityTest#testExcelInputOutput"

# 运行测试3：需求分解功能测试
mvn test -Dtest="RequirementTraceabilityTest#testRequirementDecomposition"

# 运行测试4：需求追溯功能测试
mvn test -Dtest="RequirementTraceabilityTest#testRequirementTraceability"

# 运行测试5：边界情况测试
mvn test -Dtest="RequirementTraceabilityTest#testEdgeCases"
```

#### 使用IDE运行

1. **IntelliJ IDEA**：
   - 打开 `RequirementTraceabilityTest.java`
   - 点击测试方法左侧的运行按钮（绿色三角形）
   - 或者右键点击方法名，选择"Run 'testXXX()'"

2. **Eclipse**：
   - 打开 `RequirementTraceabilityTest.java`
   - 右键点击测试方法，选择"Run As" -> "JUnit Test"

3. **VS Code**：
   - 打开测试文件
   - 点击测试方法上方的"Run Test"链接

### 3.3 方式三：使用JUnit Runner

```bash
# 运行所有测试
java -cp "target/test-classes:target/classes:$(mvn dependency:build-classpath -q -DincludeScope=test)" \
     org.junit.platform.console.ConsoleLauncher \
     --class-path target/test-classes \
     --select-class pub.developers.docautogenbyexcel.RequirementTraceabilityTest

# 运行单个测试方法
java -cp "target/test-classes:target/classes:$(mvn dependency:build-classpath -q -DincludeScope=test)" \
     org.junit.platform.console.ConsoleLauncher \
     --class-path target/test-classes \
     --select-method pub.developers.docautogenbyexcel.RequirementTraceabilityTest#testCodeInputOutput
```

## 四、测试方法详解

### 4.1 testCodeInputOutput() - 代码方式输入输出测试

**测试内容**：
- ✅ 创建需求管理器
- ✅ 创建根需求（输入）
- ✅ 分解需求（输入子需求名称列表）
- ✅ 创建测试用例（输入）
- ✅ 建立追溯关系（输入）
- ✅ 验证输出：
  - 需求树结构
  - 追溯矩阵
  - 追溯验证报告
  - 追溯覆盖率报告

**运行命令**：
```bash
mvn test -Dtest="RequirementTraceabilityTest#testCodeInputOutput"
```

### 4.2 testExcelInputOutput() - Excel方式输入输出测试

**测试内容**：
- ✅ 创建测试Excel文件（输入）
- ✅ 从Excel读取需求数据（输入）
- ✅ 从Excel读取测试用例数据（输入）
- ✅ 添加到管理器并建立追溯关系
- ✅ 验证输出（同测试1）

**测试文件**：
- `test_output/test_requirements.xlsx` - 需求Excel文件（自动生成）
- `test_output/test_cases.xlsx` - 测试用例Excel文件（自动生成）

**运行命令**：
```bash
mvn test -Dtest="RequirementTraceabilityTest#testExcelInputOutput"
```

### 4.3 testRequirementDecomposition() - 需求分解功能测试

**测试内容**：
- ✅ 测试手动分解：使用`decomposeRequirement()`方法
- ✅ 测试自动分解：使用`autoDecomposeRequirement()`方法
- ✅ 测试多级分解：对子需求进一步分解
- ✅ 验证需求树结构：确保层级关系正确

**运行命令**：
```bash
mvn test -Dtest="RequirementTraceabilityTest#testRequirementDecomposition"
```

### 4.4 testRequirementTraceability() - 需求追溯功能测试

**测试内容**：
- ✅ 创建需求和测试用例
- ✅ 建立追溯关系（包括一对多和多对一关系）
- ✅ 验证追溯查询：
  - 根据需求ID查询对应的测试用例
  - 根据测试用例ID查询对应的需求
- ✅ 验证追溯验证功能
- ✅ 验证追溯覆盖率计算

**运行命令**：
```bash
mvn test -Dtest="RequirementTraceabilityTest#testRequirementTraceability"
```

### 4.5 testEdgeCases() - 边界情况测试

**测试内容**：
- ✅ 测试空需求列表
- ✅ 测试单个需求
- ✅ 测试需求搜索
- ✅ 测试叶子需求查询
- ✅ 测试需求删除

**运行命令**：
```bash
mvn test -Dtest="RequirementTraceabilityTest#testEdgeCases"
```

## 五、测试输出示例

### 运行单个测试

```
========================================
【测试1】代码方式输入输出测试
========================================

1.1 创建需求管理器
1.2 创建根需求（输入）
  ✅ 根需求创建成功: REQ-001 用户管理系统
1.3 分解需求（输入子需求名称列表）
  ✅ 需求分解成功，创建了3个子需求
...

✅ 测试1通过
```

### 运行所有测试

```
[INFO] Tests run: 5, Failures: 0, Errors: 0, Skipped: 0
[INFO] 
[INFO] ------------------------------------------------------------------------
[INFO] BUILD SUCCESS
[INFO] ------------------------------------------------------------------------
```

## 六、快速参考

### 常用命令

```bash
# 编译测试代码
mvn test-compile

# 运行所有测试
mvn test

# 运行单个测试方法
mvn test -Dtest="RequirementTraceabilityTest#testCodeInputOutput"

# 运行多个测试方法（用逗号分隔）
mvn test -Dtest="RequirementTraceabilityTest#testCodeInputOutput,RequirementTraceabilityTest#testExcelInputOutput"

# 跳过测试，只编译
mvn package -DskipTests
```

### IDE快捷键

- **IntelliJ IDEA**：
  - `Ctrl+Shift+F10` (Windows/Linux) 或 `Ctrl+R` (Mac) - 运行当前测试方法
  - `Ctrl+Shift+F9` (Windows/Linux) 或 `Ctrl+R` (Mac) - 调试当前测试方法

- **Eclipse**：
  - `Alt+Shift+X, T` - 运行JUnit测试

## 七、测试结果说明

### 7.1 测试通过标准

- ✅ **所有断言通过**：测试方法中的所有断言都成功
- ✅ **无异常抛出**：测试过程中没有抛出未捕获的异常
- ✅ **输出验证通过**：所有输出格式正确，数据完整

### 7.2 测试失败处理

如果测试失败：
1. 查看控制台输出的错误信息
2. 检查测试数据是否正确
3. 确认所有依赖都已正确安装
4. 查看堆栈跟踪信息定位问题

## 八、注意事项

1. **Excel文件路径**：Excel测试会在`test_output`目录下创建测试文件
2. **测试隔离**：每个测试方法都是独立的，不会相互影响
3. **Spring Boot上下文**：使用`@SpringBootTest`会加载Spring Boot上下文，但测试不依赖Spring功能
4. **测试数据**：测试数据在每次运行时都会重新创建

## 九、扩展测试

如果需要添加新的测试方法：

```java
@Test
@DisplayName("新功能测试")
public void testNewFeature() {
    // 测试代码
}
```

然后在IDE中运行或使用Maven命令：
```bash
mvn test -Dtest="RequirementTraceabilityTest#testNewFeature"
```

---

*文档版本：v2.0*  
*最后更新：2025-12-01*
