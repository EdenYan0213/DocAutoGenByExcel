package pub.developers.docautogenbyexcel.service;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import pub.developers.docautogenbyexcel.baseline.LLMRequirementParser;
import pub.developers.docautogenbyexcel.hub.DataHub;
import pub.developers.docautogenbyexcel.model.Requirement;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.UUID;

@Service
public class RequirementService {

    private static final String STORAGE_DIR = "storage";
    private static final String UPLOAD_DIR = STORAGE_DIR + "/uploads";
    private static final String OUTPUT_DIR = STORAGE_DIR + "/outputs";
    private static final DateTimeFormatter TIMESTAMP_FORMAT = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

    private final DataHub dataHub;
    private final LLMRequirementParser parser;

    public RequirementService(DataHub dataHub, LLMRequirementParser parser) {
        this.dataHub = dataHub;
        this.parser = parser;
        ensureDirectories();
    }

    public record ParsedRequirement(String reqId, String title, String description, String priority) {
    }

    public record ValidationIssue(String severity, String requirementId, String field, String message,
            String suggestion) {
    }

    public record ValidationReport(List<ValidationIssue> warnings,
            List<ValidationIssue> errors,
            List<String> passes,
            Map<String, Object> stats) {
    }

    public record ParseResult(String projectName,
            String parsedAt,
            List<ParsedRequirement> requirements,
            ValidationReport report) {
    }

    public record ConfirmResult(String outputFileName, String downloadUrl) {
    }

    public ParseResult parseSrs(MultipartFile srsFile, String projectName) throws Exception {
        String srsText = readWordText(srsFile);
        List<Requirement> requirements = parser.parseRequirements(srsText);
        List<ParsedRequirement> payload = new ArrayList<>();
        for (Requirement req : requirements) {
            payload.add(new ParsedRequirement(
                    req.getRequirementId(),
                    req.getRequirementName(),
                    req.getDescription(),
                    mapPriority(req.getPriority())));
        }

        ValidationReport report = validate(payload);
        String resolvedProject = (projectName == null || projectName.isBlank())
                ? stripExtension(srsFile.getOriginalFilename())
                : projectName.trim();

        return new ParseResult(
                resolvedProject == null ? "未命名项目" : resolvedProject,
                LocalDateTime.now().toString(),
                payload,
                report);
    }

    public ConfirmResult writeRequirements(List<ParsedRequirement> requirements, MultipartFile excelFile)
            throws Exception {
        Path basePath;
        String originalName = excelFile != null ? excelFile.getOriginalFilename() : null;
        String outputFileName = buildOutputName(originalName);
        Path outputPath = Paths.get(OUTPUT_DIR, outputFileName);

        if (excelFile != null && !excelFile.isEmpty()) {
            Path uploadPath = saveUploadedFile(excelFile);
            Files.copy(uploadPath, outputPath);
            basePath = outputPath;
        } else {
            basePath = outputPath;
        }

        List<Requirement> model = new ArrayList<>();
        for (ParsedRequirement payload : requirements) {
            Requirement req = new Requirement();
            req.setRequirementId(payload.reqId());
            req.setRequirementNumber(payload.reqId());
            req.setRequirementName(payload.title());
            req.setDescription(payload.description());
            req.setPriority(parsePriority(payload.priority()));
            model.add(req);
        }
        dataHub.writeRequirements(basePath.toString(), model);

        return new ConfirmResult(outputFileName, "/api/requirements/download/" + outputFileName);
    }

    public byte[] loadRequirementExcel(String fileName) throws Exception {
        Path path = Paths.get(OUTPUT_DIR, fileName);
        return Files.readAllBytes(path);
    }

    private ValidationReport validate(List<ParsedRequirement> requirements) {
        List<ValidationIssue> warnings = new ArrayList<>();
        List<ValidationIssue> errors = new ArrayList<>();
        List<String> passes = new ArrayList<>();

        Map<String, Integer> idCounts = new HashMap<>();
        for (ParsedRequirement req : requirements) {
            String id = safe(req.reqId());
            if (!id.isBlank()) {
                idCounts.put(id, idCounts.getOrDefault(id, 0) + 1);
            }
        }

        for (ParsedRequirement req : requirements) {
            String id = safe(req.reqId());
            String title = safe(req.title());
            String desc = safe(req.description());
            String priority = safe(req.priority());

            if (id.isBlank()) {
                errors.add(new ValidationIssue("error", "(未分配)", "reqId", "需求ID缺失", "请补齐REQ-XXX"));
            }
            if (idCounts.getOrDefault(id, 0) > 1) {
                errors.add(new ValidationIssue("error", id, "reqId", "ReqID重复", "请修改为唯一编号"));
            }
            if (title.isBlank()) {
                errors.add(new ValidationIssue("error", id, "title", "标题为空", "请补充标题"));
            } else if (title.length() > 20) {
                warnings.add(new ValidationIssue("warning", id, "title", "标题超过20字", "建议精简"));
            }
            if (desc.isBlank()) {
                errors.add(new ValidationIssue("error", id, "description", "描述为空", "请补充描述"));
            }
            if (!priority.isBlank() && !isPriorityValid(priority)) {
                warnings.add(new ValidationIssue("warning", id, "priority", "优先级取值异常", "已建议改为“中”"));
            }
        }

        if (errors.isEmpty()) {
            passes.add("字段完整性校验通过");
        }
        if (warnings.isEmpty()) {
            passes.add("无需要关注的警告项");
        }
        passes.add("JSON格式合法");

        int attention = errors.size() + warnings.size();
        Map<String, Object> stats = new LinkedHashMap<>();
        stats.put("total", requirements.size());
        stats.put("warningCount", warnings.size());
        stats.put("errorCount", errors.size());
        stats.put("attention", attention);

        return new ValidationReport(warnings, errors, passes, stats);
    }

    private String mapPriority(Requirement.Priority priority) {
        if (priority == null) {
            return "中";
        }
        return priority.getDescription();
    }

    private Requirement.Priority parsePriority(String priority) {
        if (priority == null) {
            return Requirement.Priority.MEDIUM;
        }
        String normalized = priority.trim();
        if ("高".equals(normalized) || "HIGH".equalsIgnoreCase(normalized)) {
            return Requirement.Priority.HIGH;
        }
        if ("低".equals(normalized) || "LOW".equalsIgnoreCase(normalized)) {
            return Requirement.Priority.LOW;
        }
        return Requirement.Priority.MEDIUM;
    }

    private boolean isPriorityValid(String priority) {
        String value = priority.trim().toUpperCase(Locale.ROOT);
        return "高".equals(priority) || "中".equals(priority) || "低".equals(priority)
                || "HIGH".equals(value) || "MEDIUM".equals(value) || "LOW".equals(value);
    }

    private String readWordText(MultipartFile file) throws Exception {
        StringBuilder sb = new StringBuilder();
        try (XWPFDocument document = new XWPFDocument(file.getInputStream())) {
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText();
                if (text != null && !text.isBlank()) {
                    sb.append(text).append(System.lineSeparator());
                }
            }
        }
        return sb.toString();
    }

    private String stripExtension(String name) {
        if (name == null) {
            return null;
        }
        int idx = name.lastIndexOf('.');
        return idx > 0 ? name.substring(0, idx) : name;
    }

    private Path saveUploadedFile(MultipartFile file) throws Exception {
        String sessionId = UUID.randomUUID().toString().substring(0, 8);
        String safeName = sessionId + "_"
                + file.getOriginalFilename().replaceAll("[^a-zA-Z0-9._\\u4e00-\\u9fa5-]", "_");
        Path path = Paths.get(UPLOAD_DIR, safeName);
        try (var input = file.getInputStream(); var output = new FileOutputStream(path.toFile())) {
            input.transferTo(output);
        }
        return path;
    }

    private String buildOutputName(String originalName) {
        String timestamp = LocalDateTime.now().format(TIMESTAMP_FORMAT);
        String base = originalName == null ? "requirements" : stripExtension(originalName);
        return base + "_" + timestamp + ".xlsx";
    }

    private String safe(String value) {
        return value == null ? "" : value.trim();
    }

    private void ensureDirectories() {
        try {
            Files.createDirectories(Paths.get(UPLOAD_DIR));
            Files.createDirectories(Paths.get(OUTPUT_DIR));
        } catch (Exception e) {
            throw new IllegalStateException("Failed to create storage directories", e);
        }
    }
}
