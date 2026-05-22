package pub.developers.docautogenbyexcel.controller;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import pub.developers.docautogenbyexcel.service.RequirementService;

import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/requirements")
@CrossOrigin(origins = "*")
public class RequirementController {

    private final RequirementService requirementService;
    private final ObjectMapper objectMapper;

    public RequirementController(RequirementService requirementService, ObjectMapper objectMapper) {
        this.requirementService = requirementService;
        this.objectMapper = objectMapper;
    }

    @PostMapping("/parse")
    public ResponseEntity<?> parseRequirements(
            @RequestParam("srs") MultipartFile srsFile,
            @RequestParam(value = "projectName", required = false) String projectName) {
        try {
            if (srsFile == null || srsFile.isEmpty()) {
                return ResponseEntity.badRequest().body(Map.of("success", false, "error", "请上传SRS文档"));
            }
            RequirementService.ParseResult result = requirementService.parseSrs(srsFile, projectName);
            return ResponseEntity.ok(Map.of("success", true, "data", result));
        } catch (Exception e) {
            return ResponseEntity.internalServerError().body(Map.of("success", false, "error", e.getMessage()));
        }
    }

    @PostMapping("/confirm")
    public ResponseEntity<?> confirmRequirements(
            @RequestParam("requirements") String requirementsJson,
            @RequestParam(value = "excel", required = false) MultipartFile excelFile) {
        try {
            List<RequirementService.ParsedRequirement> requirements = objectMapper.readValue(
                    requirementsJson, new TypeReference<>() {
                    });
            if (requirements == null || requirements.isEmpty()) {
                return ResponseEntity.badRequest().body(Map.of("success", false, "error", "需求列表为空"));
            }
            RequirementService.ConfirmResult result = requirementService.writeRequirements(requirements, excelFile);
            return ResponseEntity.ok(Map.of(
                    "success", true,
                    "outputFileName", result.outputFileName(),
                    "downloadUrl", result.downloadUrl()));
        } catch (Exception e) {
            return ResponseEntity.internalServerError().body(Map.of("success", false, "error", e.getMessage()));
        }
    }

    @GetMapping("/download/{fileName}")
    public ResponseEntity<?> downloadRequirementExcel(
            @org.springframework.web.bind.annotation.PathVariable String fileName) {
        try {
            byte[] content = requirementService.loadRequirementExcel(fileName);
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDisposition(ContentDisposition.attachment()
                    .filename(fileName, StandardCharsets.UTF_8)
                    .build());
            headers.setContentLength(content.length);
            return ResponseEntity.ok().headers(headers).body(content);
        } catch (Exception e) {
            return ResponseEntity.notFound().build();
        }
    }
}
