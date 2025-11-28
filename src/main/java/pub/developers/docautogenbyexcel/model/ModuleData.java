package pub.developers.docautogenbyexcel.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 模块数据模型，包含模块编号和该模块下的所有测试用例
 */
public class ModuleData {
    private String moduleNumber;  // 模块编号，如"5.3"
    private List<TestCase> testCases;  // 该模块下的测试用例列表

    public ModuleData(String moduleNumber) {
        this.moduleNumber = moduleNumber;
        this.testCases = new ArrayList<>();
    }

    public void addTestCase(TestCase testCase) {
        this.testCases.add(testCase);
    }

    public String getModuleNumber() {
        return moduleNumber;
    }

    public void setModuleNumber(String moduleNumber) {
        this.moduleNumber = moduleNumber;
    }

    public List<TestCase> getTestCases() {
        return testCases;
    }

    public void setTestCases(List<TestCase> testCases) {
        this.testCases = testCases;
    }

    public int getTestCaseCount() {
        return testCases != null ? testCases.size() : 0;
    }
}

