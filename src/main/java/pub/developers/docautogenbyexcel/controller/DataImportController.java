package pub.developers.docautogenbyexcel.controller;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import pub.developers.docautogenbyexcel.dto.TestCaseImportItem;
import pub.developers.docautogenbyexcel.dto.TestResultImportItem;
import pub.developers.docautogenbyexcel.hub.ImportSummary;
import pub.developers.docautogenbyexcel.importer.TestCaseImporter;
import pub.developers.docautogenbyexcel.importer.TestResultImporter;

import java.util.List;
import java.util.Map;

/**
 * 数据接口层：测试用例与测试结果导入API。
 */
@RestController
@RequestMapping("/api")
@CrossOrigin(origins = "*")
public class DataImportController {

  private final TestCaseImporter testCaseImporter;
  private final TestResultImporter testResultImporter;

  public DataImportController(TestCaseImporter testCaseImporter, TestResultImporter testResultImporter) {
    this.testCaseImporter = testCaseImporter;
    this.testResultImporter = testResultImporter;
  }

  @PostMapping("/testcase/import")
  public ResponseEntity<?> importTestCases(
      @RequestParam("excelPath") String excelPath,
      @RequestBody List<TestCaseImportItem> testCases) {
    try {
      ImportSummary summary = testCaseImporter.importCases(excelPath, testCases);
      return ResponseEntity.ok(Map.of(
          "success", true,
          "successCount", summary.getSuccessCount(),
          "failedCount", summary.getFailedCount(),
          "warnings", summary.getWarnings(),
          "errors", summary.getErrors()));
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body(Map.of("success", false, "error", e.getMessage()));
    }
  }

  @PostMapping("/testresult/import")
  public ResponseEntity<?> importTestResults(
      @RequestParam("excelPath") String excelPath,
      @RequestBody List<TestResultImportItem> testResults) {
    try {
      ImportSummary summary = testResultImporter.importResults(excelPath, testResults);
      return ResponseEntity.ok(Map.of(
          "success", true,
          "successCount", summary.getSuccessCount(),
          "failedCount", summary.getFailedCount(),
          "warnings", summary.getWarnings(),
          "errors", summary.getErrors()));
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body(Map.of("success", false, "error", e.getMessage()));
    }
  }
}
