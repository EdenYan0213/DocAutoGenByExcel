package pub.developers.docautogenbyexcel.model;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 测试用例数据模型
 */
public class TestCase {
    private String moduleNumber;  // 模块编号，如"5.3"
    private Map<String, String> columnData; // 动态列数据，key为列名，value为列值

    public TestCase() {
        this.columnData = new LinkedHashMap<>();
    }

    public TestCase(String moduleNumber) {
        this.moduleNumber = moduleNumber;
        this.columnData = new LinkedHashMap<>();
    }

    // Getters and Setters
    public String getModuleNumber() {
        return moduleNumber;
    }

    public void setModuleNumber(String moduleNumber) {
        this.moduleNumber = moduleNumber;
    }

    /**
     * 获取所有列数据
     */
    public Map<String, String> getColumnData() {
        return columnData;
    }

    /**
     * 设置列数据
     */
    public void setColumnData(Map<String, String> columnData) {
        this.columnData = columnData;
    }

    /**
     * 添加列数据
     */
    public void addColumnData(String columnName, String value) {
        if (columnData == null) {
            columnData = new LinkedHashMap<>();
        }
        columnData.put(columnName, value != null ? value : "");
    }

    /**
     * 获取指定列的值
     */
    public String getColumnValue(String columnName) {
        return columnData != null ? columnData.getOrDefault(columnName, "") : "";
    }

    // 兼容旧的方法（为了向后兼容）
    public String getTestName() {
        // 优先查找中文列名"测试项名称"，如果找不到再查找英文列名"testName"
        String value = getColumnValue("测试项名称");
        if (value == null || value.trim().isEmpty()) {
            value = getColumnValue("testName");
        }
        return value != null ? value : "";
    }

    public void setTestName(String testName) {
        addColumnData("testName", testName);
    }

    public String getId() {
        // 优先查找中文列名"标识"，如果找不到再查找英文列名"id"
        String value = getColumnValue("标识");
        if (value == null || value.trim().isEmpty()) {
            value = getColumnValue("id");
        }
        return value != null ? value : "";
    }

    public void setId(String id) {
        addColumnData("id", id);
    }

    public String getContent() {
        return getColumnValue("content");
    }

    public void setContent(String content) {
        addColumnData("content", content);
    }

    public String getStrategy() {
        return getColumnValue("strategy");
    }

    public void setStrategy(String strategy) {
        addColumnData("strategy", strategy);
    }

    public String getCriteria() {
        return getColumnValue("criteria");
    }

    public void setCriteria(String criteria) {
        addColumnData("criteria", criteria);
    }

    public String getStopCondition() {
        return getColumnValue("stopCondition");
    }

    public void setStopCondition(String stopCondition) {
        addColumnData("stopCondition", stopCondition);
    }

    public String getTrace() {
        return getColumnValue("trace");
    }

    public void setTrace(String trace) {
        addColumnData("trace", trace);
    }

    @Override
    public String toString() {
        return "TestCase{" +
                "moduleNumber='" + moduleNumber + '\'' +
                ", columnData=" + columnData +
                '}';
    }
}

