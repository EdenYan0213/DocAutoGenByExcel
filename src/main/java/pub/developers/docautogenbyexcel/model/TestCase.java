package pub.developers.docautogenbyexcel.model;

/**
 * 测试用例数据模型
 */
public class TestCase {
    private String moduleNumber;  // 模块编号，如"5.3"
    private String testName;      // 测试名称
    private String id;            // 测试用例标识
    private String content;       // 测试内容
    private String strategy;      // 测试策略与方法
    private String criteria;      // 判定准则
    private String stopCondition; // 测试终止条件
    private String trace;         // 追踪关系

    public TestCase() {
    }

    public TestCase(String moduleNumber, String testName, String id, String content,
                    String strategy, String criteria, String stopCondition, String trace) {
        this.moduleNumber = moduleNumber;
        this.testName = testName;
        this.id = id;
        this.content = content;
        this.strategy = strategy;
        this.criteria = criteria;
        this.stopCondition = stopCondition;
        this.trace = trace;
    }

    // Getters and Setters
    public String getModuleNumber() {
        return moduleNumber;
    }

    public void setModuleNumber(String moduleNumber) {
        this.moduleNumber = moduleNumber;
    }

    public String getTestName() {
        return testName;
    }

    public void setTestName(String testName) {
        this.testName = testName;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public String getStrategy() {
        return strategy;
    }

    public void setStrategy(String strategy) {
        this.strategy = strategy;
    }

    public String getCriteria() {
        return criteria;
    }

    public void setCriteria(String criteria) {
        this.criteria = criteria;
    }

    public String getStopCondition() {
        return stopCondition;
    }

    public void setStopCondition(String stopCondition) {
        this.stopCondition = stopCondition;
    }

    public String getTrace() {
        return trace;
    }

    public void setTrace(String trace) {
        this.trace = trace;
    }

    @Override
    public String toString() {
        return "TestCase{" +
                "moduleNumber='" + moduleNumber + '\'' +
                ", testName='" + testName + '\'' +
                ", id='" + id + '\'' +
                '}';
    }
}

