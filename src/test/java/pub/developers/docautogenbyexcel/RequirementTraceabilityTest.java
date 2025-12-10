package pub.developers.docautogenbyexcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import pub.developers.docautogenbyexcel.manager.RequirementManager;
import pub.developers.docautogenbyexcel.manager.TraceabilityManager;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.reader.ExcelReader;
import pub.developers.docautogenbyexcel.reader.RequirementExcelReader;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

/**
 * 需求分解和需求追溯功能测试
 * 使用SpringBootTest，每个测试方法可以独立运行
 */
@SpringBootTest
public class RequirementTraceabilityTest {
    
    private static final String TEST_DIR = "test_output";
    private static final String REQUIREMENTS_EXCEL = TEST_DIR + "/test_requirements.xlsx";
    private static final String TEST_CASES_EXCEL = TEST_DIR + "/test_cases.xlsx";
    
    /**
     * 测试1：代码方式输入输出测试
     */
    @Test
    @DisplayName("代码方式输入输出测试")
    public void testCodeInputOutput() {
        System.out.println("\n========================================");
        System.out.println("【测试1】代码方式输入输出测试");
        System.out.println("========================================\n");
        
        try {
            System.out.println("1.1 创建需求管理器");
            RequirementManager reqManager = new RequirementManager("REQ");
            
            System.out.println("1.2 创建根需求（输入）");
            Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
            rootReq.setDescription("实现完整的用户管理功能");
            rootReq.setType(Requirement.RequirementType.FUNCTIONAL);
            rootReq.setPriority(Requirement.Priority.HIGH);
            reqManager.addRequirement(rootReq);
            
            // 验证输入
            Requirement retrieved = reqManager.getRequirement("REQ-001");
            if (retrieved == null || !retrieved.getRequirementName().equals("用户管理系统")) {
                System.out.println("  ❌ 需求创建失败");
                throw new AssertionError("需求创建失败");
            }
            System.out.println("  ✅ 根需求创建成功: " + retrieved.getRequirementNumber() + " " + retrieved.getRequirementName());
            
            System.out.println("1.3 分解需求（输入子需求名称列表）");
            List<String> childNames = Arrays.asList(
                "用户登录功能",
                "用户注册功能",
                "用户信息管理"
            );
            List<Requirement> children = reqManager.autoDecomposeRequirement("REQ-001", childNames);
            
            if (children.size() != 3) {
                System.out.println("  ❌ 需求分解失败，期望3个子需求，实际" + children.size());
                throw new AssertionError("需求分解失败");
            }
            System.out.println("  ✅ 需求分解成功，创建了" + children.size() + "个子需求");
            
            System.out.println("1.4 创建测试用例（输入）");
            List<TestCase> testCases = new ArrayList<>();
            TestCase tc1 = new TestCase();
            tc1.setId("TC-001");
            tc1.setTestName("用户登录功能测试");
            tc1.addColumnData("需求ID", "REQ-001.1");
            testCases.add(tc1);
            
            TestCase tc2 = new TestCase();
            tc2.setId("TC-002");
            tc2.setTestName("用户注册功能测试");
            tc2.addColumnData("需求ID", "REQ-001.2");
            testCases.add(tc2);
            
            System.out.println("  ✅ 创建了" + testCases.size() + "个测试用例");
            
            System.out.println("1.5 建立追溯关系（输入）");
            TraceabilityManager traceManager = new TraceabilityManager(reqManager);
            for (TestCase tc : testCases) {
                String reqNumber = tc.getColumnValue("需求ID");
                Requirement req = reqManager.findRequirementByNumber(reqNumber);
                if (req != null) {
                    traceManager.establishTraceability(req.getRequirementId(), tc.getId());
                }
            }
            
            System.out.println("  ✅ 建立了" + traceManager.generateTraceMatrix().size() + "个追溯关系");
            
            System.out.println("1.6 验证输出");
            // 输出1：需求树结构
            System.out.println("\n  输出1：需求树结构");
            RequirementManager.RequirementTree tree = reqManager.getRequirementTree();
            List<Requirement> roots = tree.getRoots();
            if (roots.isEmpty()) {
                System.out.println("    ❌ 需求树为空");
                throw new AssertionError("需求树为空");
            }
            System.out.println("    ✅ 需求树结构：");
            tree.printTree();
            
            // 输出2：追溯矩阵
            System.out.println("\n  输出2：追溯矩阵");
            Map<String, List<String>> matrix = traceManager.generateTraceMatrix();
            if (matrix.isEmpty()) {
                System.out.println("    ❌ 追溯矩阵为空");
                throw new AssertionError("追溯矩阵为空");
            }
            System.out.println("    ✅ 追溯矩阵：");
            for (Map.Entry<String, List<String>> entry : matrix.entrySet()) {
                Requirement req = reqManager.getRequirement(entry.getKey());
                System.out.println("      " + req.getRequirementNumber() + " -> " + entry.getValue());
            }
            
            // 输出3：追溯验证报告
            System.out.println("\n  输出3：追溯验证报告");
            TraceabilityManager.TraceabilityValidationResult validation = 
                traceManager.validateTraceability(reqManager.getAllRequirements(), testCases);
            System.out.println("    ✅ 验证完成");
            validation.printReport();
            
            // 输出4：追溯覆盖率报告
            System.out.println("\n  输出4：追溯覆盖率报告");
            TraceabilityManager.TraceabilityCoverage coverage = 
                traceManager.calculateCoverage(reqManager.getAllRequirements(), testCases);
            System.out.println("    ✅ 覆盖率计算完成");
            coverage.printReport();
            
            System.out.println("\n✅ 测试1通过");
        } catch (Exception e) {
            System.out.println("  ❌ 测试异常: " + e.getMessage());
            e.printStackTrace();
            throw new AssertionError("测试失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 测试2：Excel方式输入输出测试
     */
    @Test
    @DisplayName("Excel方式输入输出测试")
    public void testExcelInputOutput() {
        System.out.println("\n========================================");
        System.out.println("【测试2】Excel方式输入输出测试");
        System.out.println("========================================\n");
        
        // 创建测试目录
        createTestDirectory();
        
        try {
            System.out.println("2.1 创建测试Excel文件（输入）");
            createRequirementsExcel();
            createTestCasesExcel();
            System.out.println("  ✅ Excel文件创建成功");
            
            System.out.println("2.2 从Excel读取需求数据（输入）");
            RequirementExcelReader reqReader = new RequirementExcelReader();
            List<Requirement> requirements = reqReader.readRequirements(REQUIREMENTS_EXCEL);
            
            if (requirements.isEmpty()) {
                System.out.println("  ❌ 未读取到需求数据");
                throw new AssertionError("未读取到需求数据");
            }
            System.out.println("  ✅ 从Excel读取了" + requirements.size() + "个需求");
            
            System.out.println("2.3 从Excel读取测试用例数据（输入）");
            ExcelReader testReader = new ExcelReader();
            Map<String, ModuleData> moduleDataMap = testReader.readExcel(TEST_CASES_EXCEL);
            
            if (moduleDataMap.isEmpty()) {
                System.out.println("  ❌ 未读取到测试用例数据");
                throw new AssertionError("未读取到测试用例数据");
            }
            
            int totalTestCases = 0;
            for (ModuleData moduleData : moduleDataMap.values()) {
                totalTestCases += moduleData.getTestCases().size();
            }
            System.out.println("  ✅ 从Excel读取了" + totalTestCases + "个测试用例");
            
            System.out.println("2.4 添加到管理器并建立追溯关系");
            RequirementManager reqManager = new RequirementManager("REQ");
            for (Requirement req : requirements) {
                reqManager.addRequirement(req);
            }
            
            // 建立需求树关系
            for (Requirement req : requirements) {
                if (req.getParentRequirementId() != null && !req.getParentRequirementId().isEmpty()) {
                    Requirement parent = reqManager.getRequirement(req.getParentRequirementId());
                    if (parent != null) {
                        parent.addChild(req);
                    }
                }
            }
            
            TraceabilityManager traceManager = new TraceabilityManager(reqManager);
            List<TestCase> allTestCases = new ArrayList<>();
            for (ModuleData moduleData : moduleDataMap.values()) {
                for (TestCase testCase : moduleData.getTestCases()) {
                    allTestCases.add(testCase);
                    String requirementId = testCase.getColumnValue("需求ID");
                    if (requirementId != null && !requirementId.trim().isEmpty()) {
                        Requirement req = reqManager.findRequirementByNumber(requirementId.trim());
                        if (req != null) {
                            traceManager.establishTraceability(req.getRequirementId(), testCase.getId());
                        }
                    }
                }
            }
            
            System.out.println("  ✅ 建立了" + traceManager.generateTraceMatrix().size() + "个追溯关系");
            
            System.out.println("2.5 验证输出");
            // 输出1：需求树结构
            System.out.println("\n  输出1：需求树结构");
            RequirementManager.RequirementTree tree = reqManager.getRequirementTree();
            tree.printTree();
            
            // 输出2：追溯矩阵
            System.out.println("\n  输出2：追溯矩阵");
            Map<String, List<String>> matrix = traceManager.generateTraceMatrix();
            for (Map.Entry<String, List<String>> entry : matrix.entrySet()) {
                Requirement req = reqManager.getRequirement(entry.getKey());
                System.out.println("    " + req.getRequirementNumber() + " -> " + entry.getValue());
            }
            
            // 输出3：追溯验证报告
            System.out.println("\n  输出3：追溯验证报告");
            TraceabilityManager.TraceabilityValidationResult validation = 
                traceManager.validateTraceability(reqManager.getAllRequirements(), allTestCases);
            validation.printReport();
            
            // 输出4：追溯覆盖率报告
            System.out.println("\n  输出4：追溯覆盖率报告");
            TraceabilityManager.TraceabilityCoverage coverage = 
                traceManager.calculateCoverage(reqManager.getAllRequirements(), allTestCases);
            coverage.printReport();
            
            System.out.println("\n✅ 测试2通过");
        } catch (Exception e) {
            System.out.println("  ❌ 测试异常: " + e.getMessage());
            e.printStackTrace();
            throw new AssertionError("测试失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 测试3：需求分解功能测试
     */
    @Test
    @DisplayName("需求分解功能测试")
    public void testRequirementDecomposition() {
        System.out.println("\n========================================");
        System.out.println("【测试3】需求分解功能测试");
        System.out.println("========================================\n");
        
        try {
            RequirementManager manager = new RequirementManager("REQ");
            
            System.out.println("3.1 测试手动分解");
            Requirement root = new Requirement("REQ-001", "系统功能");
            manager.addRequirement(root);
            
            List<Requirement> children = new ArrayList<>();
            Requirement child1 = new Requirement();
            child1.setRequirementName("功能1");
            children.add(child1);
            
            Requirement child2 = new Requirement();
            child2.setRequirementName("功能2");
            children.add(child2);
            
            List<Requirement> decomposed = manager.decomposeRequirement("REQ-001", children);
            if (decomposed.size() != 2) {
                System.out.println("  ❌ 手动分解失败");
                throw new AssertionError("手动分解失败");
            }
            System.out.println("  ✅ 手动分解成功，创建了" + decomposed.size() + "个子需求");
            
            System.out.println("3.2 测试自动分解");
            List<String> childNames = Arrays.asList("功能3", "功能4");
            List<Requirement> autoDecomposed = manager.autoDecomposeRequirement("REQ-001", childNames);
            if (autoDecomposed.size() != 2) {
                System.out.println("  ❌ 自动分解失败");
                throw new AssertionError("自动分解失败");
            }
            System.out.println("  ✅ 自动分解成功，创建了" + autoDecomposed.size() + "个子需求");
            
            System.out.println("3.3 测试多级分解");
            Requirement child = decomposed.get(0);
            List<String> grandChildNames = Arrays.asList("子功能1", "子功能2");
            List<Requirement> grandChildren = manager.autoDecomposeRequirement(child.getRequirementId(), grandChildNames);
            if (grandChildren.size() != 2) {
                System.out.println("  ❌ 多级分解失败");
                throw new AssertionError("多级分解失败");
            }
            System.out.println("  ✅ 多级分解成功，创建了" + grandChildren.size() + "个孙需求");
            
            System.out.println("3.4 验证需求树结构");
            RequirementManager.RequirementTree tree = manager.getRequirementTree();
            List<Requirement> roots = tree.getRoots();
            if (roots.isEmpty()) {
                System.out.println("  ❌ 需求树为空");
                throw new AssertionError("需求树为空");
            }
            System.out.println("  ✅ 需求树结构正确");
            tree.printTree();
            
            System.out.println("\n✅ 测试3通过");
        } catch (Exception e) {
            System.out.println("  ❌ 测试异常: " + e.getMessage());
            e.printStackTrace();
            throw new AssertionError("测试失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 测试4：需求追溯功能测试
     */
    @Test
    @DisplayName("需求追溯功能测试")
    public void testRequirementTraceability() {
        System.out.println("\n========================================");
        System.out.println("【测试4】需求追溯功能测试");
        System.out.println("========================================\n");
        
        try {
            RequirementManager reqManager = new RequirementManager("REQ");
            TraceabilityManager traceManager = new TraceabilityManager(reqManager);
            
            System.out.println("4.1 创建需求和测试用例");
            Requirement req1 = new Requirement("REQ-001", "需求1");
            req1.setRequirementNumber("REQ-001");
            reqManager.addRequirement(req1);
            
            Requirement req2 = new Requirement("REQ-002", "需求2");
            req2.setRequirementNumber("REQ-002");
            reqManager.addRequirement(req2);
            
            TestCase tc1 = new TestCase();
            tc1.setId("TC-001");
            tc1.setTestName("测试用例1");
            
            TestCase tc2 = new TestCase();
            tc2.setId("TC-002");
            tc2.setTestName("测试用例2");
            
            System.out.println("  ✅ 创建了2个需求和2个测试用例");
            
            System.out.println("4.2 建立追溯关系");
            traceManager.establishTraceability("REQ-001", "TC-001");
            traceManager.establishTraceability("REQ-001", "TC-002");  // 一个需求对应多个测试用例
            traceManager.establishTraceability("REQ-002", "TC-002");  // 一个测试用例对应多个需求
            
            Map<String, List<String>> matrix = traceManager.generateTraceMatrix();
            if (matrix.size() != 2) {
                System.out.println("  ❌ 追溯关系建立失败");
                throw new AssertionError("追溯关系建立失败");
            }
            System.out.println("  ✅ 建立了" + matrix.size() + "个追溯关系");
            
            System.out.println("4.3 验证追溯查询");
            List<String> testCasesForReq1 = traceManager.getTestCaseIdsByRequirement("REQ-001");
            if (testCasesForReq1.size() != 2) {
                System.out.println("  ❌ 查询需求对应的测试用例失败");
                throw new AssertionError("查询需求对应的测试用例失败");
            }
            System.out.println("  ✅ REQ-001对应的测试用例: " + testCasesForReq1);
            
            List<String> requirementsForTc2 = traceManager.getRequirementIdsByTestCase("TC-002");
            if (requirementsForTc2.size() != 2) {
                System.out.println("  ❌ 查询测试用例对应的需求失败");
                throw new AssertionError("查询测试用例对应的需求失败");
            }
            System.out.println("  ✅ TC-002对应的需求: " + requirementsForTc2);
            
            System.out.println("4.4 验证追溯验证功能");
            List<Requirement> requirements = Arrays.asList(req1, req2);
            List<TestCase> testCases = Arrays.asList(tc1, tc2);
            TraceabilityManager.TraceabilityValidationResult validation = 
                traceManager.validateTraceability(requirements, testCases);
            
            if (!validation.isValid()) {
                System.out.println("  ⚠️ 存在未追溯的需求或测试用例（这是正常的）");
            }
            System.out.println("  ✅ 追溯验证完成");
            
            System.out.println("4.5 验证追溯覆盖率计算");
            TraceabilityManager.TraceabilityCoverage coverage = 
                traceManager.calculateCoverage(requirements, testCases);
            
            if (coverage.getRequirementCoverage() < 0 || coverage.getRequirementCoverage() > 100) {
                System.out.println("  ❌ 覆盖率计算错误");
                throw new AssertionError("覆盖率计算错误");
            }
            System.out.println("  ✅ 覆盖率计算正确");
            coverage.printReport();
            
            System.out.println("\n✅ 测试4通过");
        } catch (Exception e) {
            System.out.println("  ❌ 测试异常: " + e.getMessage());
            e.printStackTrace();
            throw new AssertionError("测试失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 测试5：边界情况测试
     */
    @Test
    @DisplayName("边界情况测试")
    public void testEdgeCases() {
        System.out.println("\n========================================");
        System.out.println("【测试5】边界情况测试");
        System.out.println("========================================\n");
        
        try {
            RequirementManager manager = new RequirementManager("REQ");
            
            System.out.println("5.1 测试空需求列表");
            List<Requirement> allReqs = manager.getAllRequirements();
            if (allReqs.size() != 0) {
                System.out.println("  ❌ 空需求列表测试失败");
                throw new AssertionError("空需求列表测试失败");
            }
            System.out.println("  ✅ 空需求列表处理正确");
            
            System.out.println("5.2 测试单个需求");
            Requirement req = new Requirement("REQ-001", "单个需求");
            manager.addRequirement(req);
            if (manager.getAllRequirements().size() != 1) {
                System.out.println("  ❌ 单个需求测试失败");
                throw new AssertionError("单个需求测试失败");
            }
            System.out.println("  ✅ 单个需求处理正确");
            
            System.out.println("5.3 测试需求搜索");
            List<Requirement> searchResults = manager.searchRequirements("单个");
            if (searchResults.size() != 1) {
                System.out.println("  ❌ 需求搜索失败");
                throw new AssertionError("需求搜索失败");
            }
            System.out.println("  ✅ 需求搜索功能正常");
            
            System.out.println("5.4 测试叶子需求查询");
            List<Requirement> leafReqs = manager.getLeafRequirements();
            if (leafReqs.size() != 1) {
                System.out.println("  ❌ 叶子需求查询失败");
                throw new AssertionError("叶子需求查询失败");
            }
            System.out.println("  ✅ 叶子需求查询功能正常");
            
            System.out.println("5.5 测试需求删除");
            manager.removeRequirement("REQ-001");
            if (manager.getAllRequirements().size() != 0) {
                System.out.println("  ❌ 需求删除失败");
                throw new AssertionError("需求删除失败");
            }
            System.out.println("  ✅ 需求删除功能正常");
            
            System.out.println("\n✅ 测试5通过");
        } catch (Exception e) {
            System.out.println("  ❌ 测试异常: " + e.getMessage());
            e.printStackTrace();
            throw new AssertionError("测试失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 创建测试目录
     */
    private static void createTestDirectory() {
        File dir = new File(TEST_DIR);
        if (!dir.exists()) {
            dir.mkdirs();
        }
    }
    
    /**
     * 创建需求Excel文件
     */
    private static void createRequirementsExcel() throws Exception {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(REQUIREMENTS_EXCEL)) {
            
            Sheet sheet = workbook.createSheet("需求");
            
            // 创建表头
            Row headerRow = sheet.createRow(0);
            String[] headers = {"需求ID", "需求编号", "需求名称", "需求描述", "需求类型", "优先级", "父需求ID", "状态"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }
            
            // 创建数据行
            Object[][] data = {
                {"REQ-001", "REQ-001", "用户管理系统", "实现完整的用户管理功能", "功能需求", "高", "", "已批准"},
                {"REQ-002", "REQ-001.1", "用户登录功能", "实现用户登录功能", "功能需求", "高", "REQ-001", "已批准"},
                {"REQ-003", "REQ-001.2", "用户注册功能", "实现用户注册功能", "功能需求", "高", "REQ-001", "已批准"},
                {"REQ-004", "REQ-001.3", "用户信息管理", "实现用户信息管理功能", "功能需求", "中", "REQ-001", "已批准"}
            };
            
            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < data[i].length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(data[i][j].toString());
                }
            }
            
            workbook.write(out);
        }
    }
    
    /**
     * 创建测试用例Excel文件
     */
    private static void createTestCasesExcel() throws Exception {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(TEST_CASES_EXCEL)) {
            
            Sheet sheet = workbook.createSheet("测试用例");
            
            // 创建表头
            Row headerRow = sheet.createRow(0);
            String[] headers = {"模块编号", "测试项名称", "标识", "需求ID", "测试内容", "测试策略与方法", "判定准则"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }
            
            // 创建数据行
            Object[][] data = {
                {"5.3", "用户登录功能测试", "TC-001", "REQ-001.1", "验证用户登录功能", "1) 输入正确用户名密码\n2) 点击登录按钮", "1) 登录成功\n2) 跳转到首页"},
                {"5.3", "用户注册功能测试", "TC-002", "REQ-001.2", "验证用户注册功能", "1) 输入注册信息\n2) 提交注册", "1) 注册成功\n2) 发送验证邮件"},
                {"5.3", "用户信息管理测试", "TC-003", "REQ-001.3", "验证用户信息管理功能", "1) 查看用户信息\n2) 修改用户信息", "1) 信息显示正确\n2) 修改成功"}
            };
            
            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < data[i].length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(data[i][j].toString());
                }
            }
            
            workbook.write(out);
        }
    }
}
