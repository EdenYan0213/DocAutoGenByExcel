package pub.developers.docautogenbyexcel.controller;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import pub.developers.docautogenbyexcel.baseline.RequirementBaselineBuilder;

import java.util.Map;

/**
 * 需求基线建立接口。
 */
@RestController
@RequestMapping("/api/baseline")
@CrossOrigin(origins = "*")
public class BaselineController {

  private final RequirementBaselineBuilder requirementBaselineBuilder;

  public BaselineController(RequirementBaselineBuilder requirementBaselineBuilder) {
    this.requirementBaselineBuilder = requirementBaselineBuilder;
  }

  @PostMapping("/build")
  public ResponseEntity<?> buildBaseline(
      @RequestParam("srsPath") String srsPath,
      @RequestParam("excelPath") String excelPath) {
    try {
      int count = requirementBaselineBuilder.buildBaseline(srsPath, excelPath);
      return ResponseEntity.ok(Map.of(
          "success", true,
          "requirementCount", count,
          "message", "需求基线建立完成"));
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body(Map.of(
          "success", false,
          "error", e.getMessage()));
    }
  }
}
