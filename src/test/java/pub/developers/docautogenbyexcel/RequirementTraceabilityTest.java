package pub.developers.docautogenbyexcel;

import org.apache.poi.ss.usermodel.*;
//package pub.developers.docautogenbyexcel;
//
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.junit.jupiter.api.DisplayName;
//import org.junit.jupiter.api.Test;
//import org.springframework.boot.test.context.SpringBootTest;
//import pub.developers.docautogenbyexcel.manager.RequirementManager;
//import pub.developers.docautogenbyexcel.manager.TraceabilityManager;
//import pub.developers.docautogenbyexcel.model.ModuleData;
//import pub.developers.docautogenbyexcel.model.Requirement;
//import pub.developers.docautogenbyexcel.model.TestCase;
//import pub.developers.docautogenbyexcel.reader.ExcelReader;
//import pub.developers.docautogenbyexcel.reader.RequirementExcelReader;
//
//import java.io.File;
//import java.io.FileOutputStream;
//import java.util.*;
//
///**
// * 需求分解和需求追溯功能测试
// * 使用SpringBootTest，每个测试方法可以独立运行
// */
////@SpringBootTest
//public class RequirementTraceabilityTest {
//
//    private static final String TEST_DIR = "test_output";
//    private static final String REQUIREMENTS_EXCEL = TEST_DIR + "/test_requirements.xlsx";
//    private static final String TEST_CASES_EXCEL = TEST_DIR + "/test_cases.xlsx";
//
//    /**
//     * 测试1：代码方式输入输出测试
//     */
//    @Test
//    @DisplayName("代码方式输入输出测试")
//    public void testCodeInputOutput() {
//        System.out.println("\n========================================");
//        System.out.println("【测试1】代码方式输入输出测试");
//        System.out.println("========================================\n");
//
//        try {
//            System.out.println("1.1 创建需求管理器");
//            RequirementManager reqManager = new RequirementManager("REQ");
//
//            System.out.println("1.2 创建根需求（输入）");
//            Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
//            rootReq.setDescription("实现完整的用户管理功能");
//            rootReq.setType(Requirement.RequirementType.FUNCTIONAL);
//            rootReq.setPriority(Requirement.Priority.HIGH);
//            reqManager.addRequirement(rootReq);
//
//            // 验证输入
//            Requirement retrieved = reqManager.getRequirement("REQ-001");
//            if (retrieved == null || !retrieved.getRequirementName().equals("用户管理系统")) {
//                System.out.println("  ❌ 需求创建失败");
//                throw new AssertionError("需求创建失败");
//            }
//            System.out.println("  ✅ 根需求创建成功: " + retrieved.getRequirementNumber() + " " + retrieved.getRequirementName());
//
//            System.out.println("1.3 分解需求（输入子需求名称列表）");
//            List<String> childNames = Arrays.asList(
//                "用户登录功能",
//                "用户注册功能",
//                "用户信息管理"
//            );
//            List<Requirement> children = reqManager.autoDecomposeRequirement("REQ-001", childNames);
//
//            if (children.size() != 3) {
//                System.out.println("  ❌ 需求分解失败，期望3个子需求，实际" + children.size());
//                throw new AssertionError("需求分解失败");
//            }
//            System.out.println("  ✅ 需求分解成功，创建了" + children.size() + "个子需求");
//
//            System.out.println("1.4 创建测试用例（输入）");
//            List<TestCase> testCases = new ArrayList<>();
//            TestCase tc1 = new TestCase();
//            tc1.setId("TC-001");
//            tc1.setTestName("用户登录功能测试");
//            tc1.addColumnData("需求ID", "REQ-001.1");
//            testCases.add(tc1);
//
//            TestCase tc2 = new TestCase();
//            tc2.setId("TC-002");
//            tc2.setTestName("用户注册功能测试");
//            tc2.addColumnData("需求ID", "REQ-001.2");
//            testCases.add(tc2);
//
//            System.out.println("  ✅ 创建了" + testCases.size() + "个测试用例");
//
//            System.out.println("1.5 建立追溯关系（输入）");
//            TraceabilityManager traceManager = new TraceabilityManager(reqManager);
//            for (TestCase tc : testCases) {
//                String reqNumber = tc.getColumnValue("需求ID");
//                Requirement req = reqManager.findRequirementByNumber(reqNumber);
//                if (req != null) {
//                    traceManager.establishTraceability(req.getRequirementId(), tc.getId());
//                }
//            }
//
//            System.out.println("  ✅ 建立了" + traceManager.generateTraceMatrix().size() + "个追溯关系");
//
//            System.out.println("1.6 验证输出");
//            // 输出1：需求树结构
//            System.out.println("\n  输出1：需求树结构");
//            RequirementManager.RequirementTree tree = reqManager.getRequirementTree();
//            List<Requirement> roots = tree.getRoots();
//            if (roots.isEmpty()) {
//                System.out.println("    ❌ 需求树为空");
//                throw new AssertionError("需求树为空");
//            }
//            System.out.println("    ✅ 需求树结构：");
//            tree.printTree();
//
//            // 输出2：追溯矩阵
//            System.out.println("\n  输出2：追溯矩阵");
//            Map<String, List<String>> matrix = traceManager.generateTraceMatrix();
//            if (matrix.isEmpty()) {
//                System.out.println("    ❌ 追溯矩阵为空");
//                throw new AssertionError("追溯矩阵为空");
//            }
//            System.out.println("    ✅ 追溯矩阵：");
//            for (Map.Entry<String, List<String>> entry : matrix.entrySet()) {
//                Requirement req = reqManager.getRequirement(entry.getKey());
//                System.out.println("      " + req.getRequirementNumber() + " -> " + entry.getValue());
//            }
//
//            // 输出3：追溯验证报告
//            System.out.println("\n  输出3：追溯验证报告");
//            TraceabilityManager.TraceabilityValidationResult validation =
//                traceManager.validateTraceability(reqManager.getAllRequirements(), testCases);
//            System.out.println("    ✅ 验证完成");
//            validation.printReport();
//
//            // 输出4：追溯覆盖率报告
//            System.out.println("\n  输出4：追溯覆盖率报告");
//            TraceabilityManager.TraceabilityCoverage coverage =
//                traceManager.calculateCoverage(reqManager.getAllRequirements(), testCases);
//            System.out.println("    ✅ 覆盖率计算完成");
//            coverage.printReport();
//
//            System.out.println("\n✅ 测试1通过");
//        } catch (Exception e) {
//            System.out.println("  ❌ 测试异常: " + e.getMessage());
//            e.printStackTrace();
//            throw new AssertionError("测试失败: " + e.getMessage(), e);
//        }
//    }
//
//    /**
//     * 测试2：Excel方式输入输出测试
//     */
//    @Test
//    @DisplayName("Excel方式输入输出测试")
//    public void testExcelInputOutput() {
//        System.out.println("\n========================================");
//        System.out.println("【测试2】Excel方式输入输出测试");
//        System.out.println("========================================\n");
//
//        // 创建测试目录
//        createTestDirectory();
//
//        try {
//            System.out.println("2.1 创建测试Excel文件（输入）");
//            createRequirementsExcel();
//            createTestCasesExcel();
//            System.out.println("  ✅ Excel文件创建成功");
//
//            System.out.println("2.2 从Excel读取需求数据（输入）");
//            RequirementExcelReader reqReader = new RequirementExcelReader();
//            List<Requirement> requirements = reqReader.readRequirements(REQUIREMENTS_EXCEL);
//
//            if (requirements.isEmpty()) {
//                System.out.println("  ❌ 未读取到需求数据");
//                throw new AssertionError("未读取到需求数据");
//            }
//            System.out.println("  ✅ 从Excel读取了" + requirements.size() + "个需求");
//
//            System.out.println("2.3 从Excel读取测试用例数据（输入）");
//            ExcelReader testReader = new ExcelReader();
//            Map<String, ModuleData> moduleDataMap = testReader.readExcel(TEST_CASES_EXCEL);
//
//            if (moduleDataMap.isEmpty()) {
//                System.out.println("  ❌ 未读取到测试用例数据");
//                throw new AssertionError("未读取到测试用例数据");
//            }
//
//            int totalTestCases = 0;
//            for (ModuleData moduleData : moduleDataMap.values()) {
//                totalTestCases += moduleData.getTestCases().size();
//            }
//            System.out.println("  ✅ 从Excel读取了" + totalTestCases + "个测试用例");
//
//            System.out.println("2.4 添加到管理器并建立追溯关系");
//            RequirementManager reqManager = new RequirementManager("REQ");
//            for (Requirement req : requirements) {
//                reqManager.addRequirement(req);
//            }
//
//            // 建立需求树关系
//            for (Requirement req : requirements) {
//                if (req.getParentRequirementId() != null && !req.getParentRequirementId().isEmpty()) {
//                    Requirement parent = reqManager.getRequirement(req.getParentRequirementId());
//                    if (parent != null) {
//                        parent.addChild(req);
//                    }
//                }
//            }
//
//            TraceabilityManager traceManager = new TraceabilityManager(reqManager);
//            List<TestCase> allTestCases = new ArrayList<>();
//            for (ModuleData moduleData : moduleDataMap.values()) {
//                for (TestCase testCase : moduleData.getTestCases()) {
//                    allTestCases.add(testCase);
//                    String requirementId = testCase.getColumnValue("需求ID");
//                    if (requirementId != null && !requirementId.trim().isEmpty()) {
//                        Requirement req = reqManager.findRequirementByNumber(requirementId.trim());
//                        if (req != null) {
//                            traceManager.establishTraceability(req.getRequirementId(), testCase.getId());
//                        }
//                    }
//                }
//            }
//
//            System.out.println("  ✅ 建立了" + traceManager.generateTraceMatrix().size() + "个追溯关系");
//
//            System.out.println("2.5 验证输出");
//            // 输出1：需求树结构
//            System.out.println("\n  输出1：需求树结构");
//            RequirementManager.RequirementTree tree = reqManager.getRequirementTree();
//            tree.printTree();
//
//            // 输出2：追溯矩阵
//            System.out.println("\n  输出2：追溯矩阵");
//            Map<String, List<String>> matrix = traceManager.generateTraceMatrix();
//            for (Map.Entry<String, List<String>> entry : matrix.entrySet()) {
//                Requirement req = reqManager.getRequirement(entry.getKey());
//                System.out.println("    " + req.getRequirementNumber() + " -> " + entry.getValue());
//            }
//
//            // 输出3：追溯验证报告
//            System.out.println("\n  输出3：追溯验证报告");
//            TraceabilityManager.TraceabilityValidationResult validation =
//                traceManager.validateTraceability(reqManager.getAllRequirements(), allTestCases);
//            validation.printReport();
//
//            // 输出4：追溯覆盖率报告
//            System.out.println("\n  输出4：追溯覆盖率报告");
//            TraceabilityManager.TraceabilityCoverage coverage =
//                traceManager.calculateCoverage(reqManager.getAllRequirements(), allTestCases);
//            coverage.printReport();
//
//            System.out.println("\n✅ 测试2通过");
//        } catch (Exception e) {
//            System.out.println("  ❌ 测试异常: " + e.getMessage());
//            e.printStackTrace();
//            throw new AssertionError("测试失败: " + e.getMessage(), e);
//        }
//    }
//}
//    @Test
