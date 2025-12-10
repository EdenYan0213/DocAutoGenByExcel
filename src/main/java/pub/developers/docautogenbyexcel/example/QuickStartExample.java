package pub.developers.docautogenbyexcel.example;

import pub.developers.docautogenbyexcel.manager.RequirementManager;
import pub.developers.docautogenbyexcel.manager.TraceabilityManager;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;

import java.util.*;

/**
 * 快速开始示例
 * 演示如何快速使用需求分解和需求追溯功能
 */
public class QuickStartExample {
    
    public static void main(String[] args) {
        System.out.println("========== 快速开始：需求分解和追溯 ==========\n");
        
        // 步骤1：创建需求管理器
        RequirementManager reqManager = new RequirementManager("REQ");
        
        // 步骤2：创建需求并分解
        createAndDecomposeRequirements(reqManager);
        
        // 步骤3：创建测试用例
        List<TestCase> testCases = createTestCases();
        
        // 步骤4：建立追溯关系
        TraceabilityManager traceManager = new TraceabilityManager(reqManager);
        establishTraceability(reqManager, testCases, traceManager);
        
        // 步骤5：生成报告
        generateReports(reqManager, testCases, traceManager);
    }
    
    /**
     * 创建需求并分解
     */
    private static void createAndDecomposeRequirements(RequirementManager manager) {
        System.out.println("【步骤1】创建需求并分解");
        
        // 创建根需求
        Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
        rootReq.setDescription("实现完整的用户管理功能");
        rootReq.setType(Requirement.RequirementType.FUNCTIONAL);
        rootReq.setPriority(Requirement.Priority.HIGH);
        manager.addRequirement(rootReq);
        System.out.println("  ✓ 创建根需求: " + rootReq.getRequirementNumber() + " " + rootReq.getRequirementName());
        
        // 自动分解为子需求
        List<String> childNames = Arrays.asList(
            "用户登录功能",
            "用户注册功能",
            "用户信息管理"
        );
        List<Requirement> children = manager.autoDecomposeRequirement("REQ-001", childNames);
        System.out.println("  ✓ 分解为 " + children.size() + " 个子需求");
        
        // 对"用户登录功能"进一步分解
        Requirement loginReq = children.get(0);
        List<String> loginChildren = Arrays.asList(
            "用户名密码验证",
            "登录状态管理"
        );
        manager.autoDecomposeRequirement(loginReq.getRequirementId(), loginChildren);
        System.out.println("  ✓ 进一步分解'用户登录功能'为 2 个子需求");
        
        // 打印需求树
        System.out.println("\n需求树结构：");
        RequirementManager.RequirementTree tree = manager.getRequirementTree();
        tree.printTree();
        System.out.println();
    }
    
    /**
     * 创建测试用例
     */
    private static List<TestCase> createTestCases() {
        System.out.println("【步骤2】创建测试用例");
        
        List<TestCase> testCases = new ArrayList<>();
        
        TestCase tc1 = new TestCase();
        tc1.setId("TC-001");
        tc1.setTestName("用户登录功能测试");
        tc1.addColumnData("需求ID", "REQ-001.1");  // 关联到"用户登录功能"
        testCases.add(tc1);
        System.out.println("  ✓ 创建测试用例: " + tc1.getId() + " " + tc1.getTestName());
        
        TestCase tc2 = new TestCase();
        tc2.setId("TC-002");
        tc2.setTestName("用户名密码验证测试");
        tc2.addColumnData("需求ID", "REQ-001.1.1");  // 关联到"用户名密码验证"
        testCases.add(tc2);
        System.out.println("  ✓ 创建测试用例: " + tc2.getId() + " " + tc2.getTestName());
        
        TestCase tc3 = new TestCase();
        tc3.setId("TC-003");
        tc3.setTestName("用户注册功能测试");
        tc3.addColumnData("需求ID", "REQ-001.2");  // 关联到"用户注册功能"
        testCases.add(tc3);
        System.out.println("  ✓ 创建测试用例: " + tc3.getId() + " " + tc3.getTestName());
        
        System.out.println();
        return testCases;
    }
    
    /**
     * 建立追溯关系
     */
    private static void establishTraceability(
            RequirementManager reqManager,
            List<TestCase> testCases,
            TraceabilityManager traceManager) {
        
        System.out.println("【步骤3】建立追溯关系");
        
        // 从测试用例的"需求ID"字段自动建立追溯关系
        for (TestCase testCase : testCases) {
            String requirementId = testCase.getColumnValue("需求ID");
            if (requirementId != null && !requirementId.trim().isEmpty()) {
                // 根据需求编号查找需求ID
                Requirement req = reqManager.findRequirementByNumber(requirementId.trim());
                if (req != null) {
                    traceManager.establishTraceability(req.getRequirementId(), testCase.getId());
                    System.out.println("  ✓ 建立追溯: " + req.getRequirementNumber() + " <-> " + testCase.getId());
                } else {
                    System.out.println("  ✗ 未找到需求: " + requirementId);
                }
            }
        }
        System.out.println();
    }
    
    /**
     * 生成报告
     */
    private static void generateReports(
            RequirementManager reqManager,
            List<TestCase> testCases,
            TraceabilityManager traceManager) {
        
        System.out.println("【步骤4】生成报告");
        
        // 1. 追溯矩阵
        System.out.println("\n1. 追溯矩阵：");
        Map<String, List<String>> matrix = traceManager.generateTraceMatrix();
        for (Map.Entry<String, List<String>> entry : matrix.entrySet()) {
            Requirement req = reqManager.getRequirement(entry.getKey());
            System.out.println("   " + req.getRequirementNumber() + " " + req.getRequirementName());
            System.out.println("      -> 测试用例: " + entry.getValue());
        }
        
        // 2. 追溯验证
        System.out.println("\n2. 追溯关系验证：");
        List<Requirement> allReqs = reqManager.getAllRequirements();
        TraceabilityManager.TraceabilityValidationResult validation = 
                traceManager.validateTraceability(allReqs, testCases);
        validation.printReport();
        
        // 3. 追溯覆盖率
        System.out.println("\n3. 追溯覆盖率：");
        TraceabilityManager.TraceabilityCoverage coverage = 
                traceManager.calculateCoverage(allReqs, testCases);
        coverage.printReport();
        
        System.out.println("\n========== 完成 ==========");
    }
}

