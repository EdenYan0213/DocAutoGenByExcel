package pub.developers.docautogenbyexcel.hub;

import java.util.ArrayList;
import java.util.List;

/**
 * 写入导入摘要。
 */
public class ImportSummary {
  private int successCount;
  private int failedCount;
  private final List<String> warnings;
  private final List<String> errors;

  public ImportSummary() {
    this.warnings = new ArrayList<>();
    this.errors = new ArrayList<>();
  }

  public int getSuccessCount() {
    return successCount;
  }

  public void setSuccessCount(int successCount) {
    this.successCount = successCount;
  }

  public int getFailedCount() {
    return failedCount;
  }

  public void setFailedCount(int failedCount) {
    this.failedCount = failedCount;
  }

  public List<String> getWarnings() {
    return warnings;
  }

  public List<String> getErrors() {
    return errors;
  }

  public void addWarning(String warning) {
    this.warnings.add(warning);
  }

  public void addError(String error) {
    this.errors.add(error);
  }
}
