package pub.developers.docautogenbyexcel.dto;

/**
 * 测试结果导入DTO。
 */
public class TestResultImportItem {
  private String tcid;
  private String executionDate;
  private String result;
  private String defectId;

  public String getTcid() {
    return tcid;
  }

  public void setTcid(String tcid) {
    this.tcid = tcid;
  }

  public String getExecutionDate() {
    return executionDate;
  }

  public void setExecutionDate(String executionDate) {
    this.executionDate = executionDate;
  }

  public String getResult() {
    return result;
  }

  public void setResult(String result) {
    this.result = result;
  }

  public String getDefectId() {
    return defectId;
  }

  public void setDefectId(String defectId) {
    this.defectId = defectId;
  }
}
