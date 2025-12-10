package pub.developers.docautogenbyexcel.example;

import pub.developers.docautogenbyexcel.manager.RequirementManager;
import pub.developers.docautogenbyexcel.manager.TraceabilityManager;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.reader.RequirementExcelReader;

import java.util.*;

/**
 * 需求分解和需求追溯功能使用示例
 */
public class RequirementTraceabilityExample {
    
    public static void main(String[] args) {
        try {
            // 示例1：需求分解
            demonstrateRequirementDecomposition();
            
            // 示例2：需求追溯
            demonstrateRequirementTraceability();
            
            // 示例3：从Excel读取需求并建立追溯关系
            demonstrateExcelIntegration();
            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 示例1：演示需求分解功能
     */
    private static void demonstrateRequirementDecomposition() {
        System.out.println("\n========== 示例1：需求分解 ==========");
        
        RequirementManager manager = new RequirementManager("REQ");
        
        // 创建根需求
        Requirement rootReq = new Requirement("REQ-001", "用户登录功能");
        rootReq.setDescription("实现用户登录功能，包括用户名密码登录");
        rootReq.setType(Requirement.RequirementType.FUNCTIONAL);
        rootReq.setPriority(Requirement.Priority.HIGH);
        manager.addRequirement(rootReq);
        
        // 手动分解：创建子需求
        List<Requirement> children = new ArrayList<>();
        
        Requirement child1 = new Requirement();
        child1.setRequirementName("用户名密码验证");
        child1.setDescription("验证用户输入的用户名和密码是否正确");
        child1.setType(Requirement.RequirementType.FUNCTIONAL);
        children.add(child1);
        
        Requirement child2 = new Requirement();
        child2.setRequirementName("登录状态管理");
        child2.setDescription("管理用户的登录状态，包括登录、登出");
        child2.setType(Requirement.RequirementType.FUNCTIONAL);
        children.add(child2);
        
        Requirement child3 = new Requirement();
        child3.setRequirementName("错误处理");
        child3.setDescription("处理登录失败、账户锁定等错误情况");
        child3.setType(Requirement.RequirementType.FUNCTIONAL);
        children.add(child3);
        
        // 执行分解
        List<Requirement> decomposed = manager.decomposeRequirement("REQ-001", children);
        System.out.println("需求分解完成，共创建 " + decomposed.size() + " 个子需求");
        
        // 自动分解：根据名称列表自动创建子需求
        List<String> childNames = Arrays.asList("记住密码功能", "自动登录功能");
        List<Requirement> autoDecomposed = manager.autoDecomposeRequirement("REQ-001", childNames);
        System.out.println("自动分解完成，共创建 " + autoDecomposed.size() + " 个子需求");
        
        // 打印需求树
        System.out.println("\n需求树结构：");
        RequirementManager.RequirementTree tree = manager.getRequirementTree();
        tree.printTree();
        
        // 查询功能
        System.out.println("\n查询功能：");
        List<Requirement> searchResults = manager.searchRequirements("登录");
        System.out.println("搜索'登录'，找到 " + searchResults.size() + " 个需求");
        
        List<Requirement> leafReqs = manager.getLeafRequirements();
        System.out.println("叶子需求数量: " + leafReqs.size());
    }
    
    /**
     * 示例2：演示需求追溯功能
     */
    private static void demonstrateRequirementTraceability() {
        System.out.println("\n========== 示例2：需求追溯 ==========");
        
        RequirementManager reqManager = new RequirementManager("REQ");
        TraceabilityManager traceManager = new TraceabilityManager(reqManager);
        
        // 创建需求
        Requirement req1 = new Requirement("REQ-001", "用户登录功能");
        req1.setRequirementNumber("REQ-001");
        reqManager.addRequirement(req1);
        
        Requirement req2 = new Requirement("REQ-002", "用户注册功能");
        req2.setRequirementNumber("REQ-002");
        reqManager.addRequirement(req2);
        
        // 创建测试用例
        TestCase test1 = new TestCase();
        test1.setId("TC-001");
        test1.setTestName("登录功能测试");
        
        TestCase test2 = new TestCase();
        test2.setId("TC-002");
        test2.setTestName("注册功能测试");
        
        TestCase test3 = new TestCase();
        test3.setId("TC-003");
        test3.setTestName("登录边界测试");
        
        // 建立追溯关系
        traceManager.establishTraceability("REQ-001", "TC-001");
        traceManager.establishTraceability("REQ-001", "TC-003");
        traceManager.establishTraceability("REQ-002", "TC-002");
        
        System.out.println("已建立3个追溯关系");
        
        // 查询追溯关系
        System.out.println("\n查询功能：");
        List<String> testCasesForReq1 = traceManager.getTestCaseIdsByRequirement("REQ-001");
        System.out.println("REQ-001对应的测试用例: " + testCasesForReq1);
        
        List<String> requirementsForTest1 = traceManager.getRequirementIdsByTestCase("TC-001");
        System.out.println("TC-001对应的需求: " + requirementsForTest1);
        
        // 生成追溯矩阵
        System.out.println("\n追溯矩阵：");
        Map<String, List<String>> matrix = traceManager.generateTraceMatrix();
        for (Map.Entry<String, List<String>> entry : matrix.entrySet()) {
            System.out.println("需求 " + entry.getKey() + " -> 测试用例 " + entry.getValue());
        }
        
        // 验证追溯关系
        System.out.println("\n追溯关系验证：");
        List<Requirement> requirements = Arrays.asList(req1, req2);
        List<TestCase> testCases = Arrays.asList(test1, test2, test3);
        
        TraceabilityManager.TraceabilityValidationResult validation = 
                traceManager.validateTraceability(requirements, testCases);
        validation.printReport();
        
        // 计算覆盖率
        System.out.println("\n追溯覆盖率：");
        TraceabilityManager.TraceabilityCoverage coverage = 
                traceManager.calculateCoverage(requirements, testCases);
        coverage.printReport();
    }
    
    /**
     * 示例3：从Excel读取需求并建立追溯关系
     */
    private static void demonstrateExcelIntegration() {
        System.out.println("\n========== 示例3：Excel集成 ==========");
        
        try {
            // 读取需求Excel（需要提供实际的Excel文件路径）
            String requirementExcelPath = "requirements.xlsx";
            
            // 如果文件不存在，跳过此示例
            java.io.File file = new java.io.File(requirementExcelPath);
            if (!file.exists()) {
                System.out.println("需求Excel文件不存在，跳过此示例");
                System.out.println("Excel格式要求：");
                System.out.println("  必填列：需求ID, 需求名称");
                System.out.println("  可选列：需求编号, 需求描述, 需求类型, 优先级, 父需求ID, 状态");
                return;
            }
            
            RequirementExcelReader reader = new RequirementExcelReader();
            List<Requirement> requirements = reader.readRequirements(requirementExcelPath);
            
            // 添加到需求管理器
            RequirementManager reqManager = new RequirementManager("REQ");
            for (Requirement req : requirements) {
                reqManager.addRequirement(req);
            }
            
            System.out.println("从Excel读取了 " + requirements.size() + " 个需求");
            
            // 建立需求树关系
            for (Requirement req : requirements) {
                if (req.getParentRequirementId() != null && !req.getParentRequirementId().isEmpty()) {
                    Requirement parent = reqManager.getRequirement(req.getParentRequirementId());
                    if (parent != null) {
                        parent.addChild(req);
                    }
                }
            }
            
            // 打印需求树
            System.out.println("\n需求树结构：");
            RequirementManager.RequirementTree tree = reqManager.getRequirementTree();
            tree.printTree();
            
        } catch (Exception e) {
            System.out.println("Excel读取失败: " + e.getMessage());
        }
    }
}

