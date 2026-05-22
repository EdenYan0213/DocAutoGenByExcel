package pub.developers.docautogenbyexcel.importer;

import org.springframework.stereotype.Service;
import pub.developers.docautogenbyexcel.dto.TestResultImportItem;
import pub.developers.docautogenbyexcel.hub.DataHub;
import pub.developers.docautogenbyexcel.hub.ImportSummary;
import pub.developers.docautogenbyexcel.model.TestResult;

import java.util.ArrayList;
import java.util.List;

/**
 * 测试结果导入组件。
 */
@Service
public class TestResultImporter {

  private final DataHub dataHub;

  public TestResultImporter(DataHub dataHub) {
    this.dataHub = dataHub;
  }

  public ImportSummary importResults(String excelPath, List<TestResultImportItem> items) throws Exception {
    List<TestResult> results = new ArrayList<>();
    for (TestResultImportItem item : items) {
      TestResult tr = new TestResult();
      tr.setTestCaseId(nonNull(item.getTcid()));
      tr.setExecutionDate(nonNull(item.getExecutionDate()));
      tr.setVerdict(nonNull(item.getResult()));
      tr.setDefectId(nonNull(item.getDefectId()));
      results.add(tr);
    }
    return dataHub.appendTestResults(excelPath, results);
  }

  private String nonNull(String text) {
    return text == null ? "" : text;
  }
}
