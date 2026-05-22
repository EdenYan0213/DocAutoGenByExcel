package pub.developers.docautogenbyexcel.dto;

import java.util.ArrayList;
import java.util.List;

/**
 * 测试用例导入DTO。
 */
public class TestCaseImportItem {
  private String tcid;
  private String title;
  private String preconditions;
  private String procedure;
  private String expectedResult;
  private List<String> reqIds = new ArrayList<>();

  public String getTcid() {
    return tcid;
  }

  public void setTcid(String tcid) {
    this.tcid = tcid;
  }

  public String getTitle() {
    return title;
  }

  public void setTitle(String title) {
    this.title = title;
  }

  public String getPreconditions() {
    return preconditions;
  }

  public void setPreconditions(String preconditions) {
    this.preconditions = preconditions;
  }

  public String getProcedure() {
    return procedure;
  }

  public void setProcedure(String procedure) {
    this.procedure = procedure;
  }

  public String getExpectedResult() {
    return expectedResult;
  }

  public void setExpectedResult(String expectedResult) {
    this.expectedResult = expectedResult;
  }

  public List<String> getReqIds() {
    return reqIds;
  }

  public void setReqIds(List<String> reqIds) {
    this.reqIds = reqIds;
  }
}
