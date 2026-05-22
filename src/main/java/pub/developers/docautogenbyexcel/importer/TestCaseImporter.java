package pub.developers.docautogenbyexcel.importer;

import org.springframework.stereotype.Service;
import pub.developers.docautogenbyexcel.dto.TestCaseImportItem;
import pub.developers.docautogenbyexcel.hub.DataHub;
import pub.developers.docautogenbyexcel.hub.ImportSummary;
import pub.developers.docautogenbyexcel.model.TestCase;

import java.util.ArrayList;
import java.util.List;

/**
 * 测试用例导入组件。
 */
@Service
public class TestCaseImporter {

  private final DataHub dataHub;

  public TestCaseImporter(DataHub dataHub) {
    this.dataHub = dataHub;
  }

  public ImportSummary importCases(String excelPath, List<TestCaseImportItem> items) throws Exception {
    List<TestCase> testCases = new ArrayList<>();
    for (TestCaseImportItem item : items) {
      TestCase tc = new TestCase("0");
      tc.addColumnData("TCID", nonNull(item.getTcid()));
      tc.addColumnData("测试用例标识", nonNull(item.getTcid()));
      tc.addColumnData("Title", nonNull(item.getTitle()));
      tc.addColumnData("Preconditions", nonNull(item.getPreconditions()));
      tc.addColumnData("Procedure", nonNull(item.getProcedure()));
      tc.addColumnData("ExpectedResult", nonNull(item.getExpectedResult()));
      tc.addColumnData("ReqID", String.join(",", item.getReqIds() == null ? List.of() : item.getReqIds()));
      testCases.add(tc);
    }
    return dataHub.writeTestCases(excelPath, testCases);
  }

  private String nonNull(String text) {
    return text == null ? "" : text;
  }
}
