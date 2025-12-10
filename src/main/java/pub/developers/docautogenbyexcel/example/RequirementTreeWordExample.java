package pub.developers.docautogenbyexcel.example;

import pub.developers.docautogenbyexcel.generator.RequirementTreeWordGenerator;
import pub.developers.docautogenbyexcel.manager.RequirementManager;
import pub.developers.docautogenbyexcel.manager.TraceabilityManager;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;

import java.util.Arrays;
import java.util.List;

/**
 * 需求树Word文档生成示例
 */
public class RequirementTreeWordExample {
    
    public static void main(String[] args) {
        try {
            System.out.println("========================================");
            System.out.println("需求树Word文档生成示例");
            System.out.println("========================================\n");
            
            // 1. 创建需求管理器并构建需求树
            RequirementManager reqManager = new RequirementManager("REQ");
            buildRequirementTree(reqManager);
            
            // 2. 生成基础需求树Word文档
            System.out.println("【示例1】生成基础需求树Word文档");
            RequirementTreeWordGenerator generator = new RequirementTreeWordGenerator();
            generator.generateRequirementTreeWord(reqManager, "output/requirement_tree.docx");
            System.out.println("✅ 基础需求树Word文档生成成功\n");
            
            // 3. 建立追溯关系
            System.out.println("【示例2】建立追溯关系");
            TraceabilityManager traceManager = new TraceabilityManager(reqManager);
            List<TestCase> testCases = createTestCases();
            establishTraceability(reqManager, traceManager, testCases);
            System.out.println("✅ 追溯关系建立完成\n");
            
            // 4. 生成带追溯信息的需求树Word文档
            System.out.println("【示例3】生成带追溯信息的需求树Word文档");
            generator.generateRequirementTreeWordWithTraceability(
                    reqManager, traceManager, "output/requirement_tree_with_trace.docx");
            System.out.println("✅ 带追溯信息的需求树Word文档生成成功\n");
            
            System.out.println("========================================");
            System.out.println("所有Word文档生成完成！");
            System.out.println("========================================");
            System.out.println("输出文件：");
            System.out.println("  - output/requirement_tree.docx (基础需求树)");
            System.out.println("  - output/requirement_tree_with_trace.docx (带追溯信息的需求树)");
            System.out.println("========================================");
            
        } catch (Exception e) {
            System.err.println("生成失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 构建需求树
     */
    private static void buildRequirementTree(RequirementManager manager) {
        // 创建根需求
        Requirement rootReq = new Requirement("REQ-001", "用户管理系统");
        rootReq.setDescription("实现完整的用户管理功能，包括登录、注册、信息管理");
        rootReq.setType(Requirement.RequirementType.FUNCTIONAL);
        rootReq.setPriority(Requirement.Priority.HIGH);
        manager.addRequirement(rootReq);
        
        // 分解为子需求
        List<String> childNames = Arrays.asList(
            "用户登录功能",
            "用户注册功能",
            "用户信息管理"
        );
        List<Requirement> children = manager.autoDecomposeRequirement("REQ-001", childNames);
        
        // 设置子需求描述
        children.get(0).setDescription("实现用户登录功能，支持用户名密码登录");
        children.get(1).setDescription("实现用户注册功能，支持邮箱注册");
        children.get(2).setDescription("实现用户信息管理功能，包括查看和修改");
        
        // 进一步分解"用户登录功能"
        Requirement loginReq = children.get(0);
        List<String> loginChildren = Arrays.asList(
            "用户名密码验证",
            "登录状态管理"
        );
        List<Requirement> loginSubChildren = manager.autoDecomposeRequirement(loginReq.getRequirementId(), loginChildren);
        
        loginSubChildren.get(0).setDescription("验证用户输入的用户名和密码是否正确");
        loginSubChildren.get(1).setDescription("管理用户的登录状态，包括登录、登出");
        
        System.out.println("✅ 需求树构建完成");
        System.out.println("需求树结构：");
        RequirementManager.RequirementTree tree = manager.getRequirementTree();
        tree.printTree();
        System.out.println();
    }
    
    /**
     * 创建测试用例
     */
    private static List<TestCase> createTestCases() {
        TestCase tc1 = new TestCase();
        tc1.setId("TC-001");
        tc1.setTestName("用户登录功能测试");
        tc1.addColumnData("需求ID", "REQ-001.1");
        
        TestCase tc2 = new TestCase();
        tc2.setId("TC-002");
        tc2.setTestName("用户名密码验证测试");
        tc2.addColumnData("需求ID", "REQ-001.1.1");
        
        TestCase tc3 = new TestCase();
        tc3.setId("TC-003");
        tc3.setTestName("用户注册功能测试");
        tc3.addColumnData("需求ID", "REQ-001.2");
        
        return Arrays.asList(tc1, tc2, tc3);
    }
    
    /**
     * 建立追溯关系
     */
    private static void establishTraceability(
            RequirementManager reqManager,
            TraceabilityManager traceManager,
            List<TestCase> testCases) {
        
        for (TestCase tc : testCases) {
            String reqNumber = tc.getColumnValue("需求ID");
            Requirement req = reqManager.findRequirementByNumber(reqNumber);
            if (req != null) {
                traceManager.establishTraceability(req.getRequirementId(), tc.getId());
                System.out.println("  ✓ 建立追溯: " + req.getRequirementNumber() + " <-> " + tc.getId());
            }
        }
    }
}

