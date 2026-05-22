package pub.developers.docautogenbyexcel.model;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 测试结果数据模型。
 */
public class TestResult {
  private String testCaseId;
  private String executionDate;
  private String verdict;
  private String defectId;
  private Map<String, String> attributes;

  public TestResult() {
    this.attributes = new LinkedHashMap<>();
  }

  public TestResult(String testCaseId, String executionDate, String verdict, String defectId) {
    this();
    this.testCaseId = testCaseId;
    this.executionDate = executionDate;
    this.verdict = verdict;
    this.defectId = defectId;
  }

  public void addAttribute(String key, String value) {
    attributes.put(key, value != null ? value : "");
  }

  public String getAttribute(String key) {
    return attributes.getOrDefault(key, "");
  }

  public String getTestCaseId() {
    return testCaseId;
  }

  public void setTestCaseId(String testCaseId) {
    this.testCaseId = testCaseId;
  }

  public String getExecutionDate() {
    return executionDate;
  }

  public void setExecutionDate(String executionDate) {
    this.executionDate = executionDate;
  }

  public String getVerdict() {
    return verdict;
  }

  public void setVerdict(String verdict) {
    this.verdict = verdict;
  }

  public String getDefectId() {
    return defectId;
  }

  public void setDefectId(String defectId) {
    this.defectId = defectId;
  }

  public Map<String, String> getAttributes() {
    return attributes;
  }

  public void setAttributes(Map<String, String> attributes) {
    this.attributes = attributes;
  }
}
