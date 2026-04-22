package pub.developers.docautogenbyexcel.service;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.processor.TableFillProcessor;
import pub.developers.docautogenbyexcel.processor.WordProcessor;
import pub.developers.docautogenbyexcel.reader.ExcelReader;
import pub.developers.docautogenbyexcel.reader.TableDataReader;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 文档处理服务
 * 支持本地存储和S3云存储
 */
@Service
public class DocumentService {

    private static final String STORAGE_DIR = "storage";
    private static final String UPLOAD_DIR = STORAGE_DIR + "/uploads";
    private static final String OUTPUT_DIR = STORAGE_DIR + "/outputs";

    private static final DateTimeFormatter TIMESTAMP_FORMAT = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

    @Value("${storage.type:local}")
    private String storageType;

    @Autowired(required = false)
    private S3StorageService s3StorageService;

    public DocumentService() {
        // 确保存储目录存在（仅本地存储需要）
        try {
            Files.createDirectories(Paths.get(UPLOAD_DIR));
            Files.createDirectories(Paths.get(OUTPUT_DIR));
        } catch (IOException e) {
            throw new RuntimeException("无法创建存储目录", e);
        }
    }

    /**
     * 文档处理结果
     */
    public record ProcessResult(
            String outputId,
            String outputFileName,
            String outputPath,
            int moduleCount,
            String message) {
    }

    /**
     * 已处理的文档信息
     */
    public record DocumentInfo(
            String id,
            String fileName,
            String originalExcel,
            String originalWord,
            long fileSize,
            String createdAt) {
    }

    /**
     * 处理上传的Excel和Word文件
     */
    public ProcessResult processDocuments(InputStream excelStream, String excelFileName,
            InputStream wordStream, String wordFileName) throws Exception {
        String timestamp = LocalDateTime.now().format(TIMESTAMP_FORMAT);
        String sessionId = UUID.randomUUID().toString().substring(0, 8);

        // 保存上传的文件
        String excelPath = saveUploadedFile(excelStream, excelFileName, sessionId);
        String wordPath = saveUploadedFile(wordStream, wordFileName, sessionId);

        // 生成输出文件名
        String baseName = wordFileName.replaceAll("\\.docx?$", "");
        String outputFileName = baseName + "_" + timestamp + ".docx";
        String outputPath = OUTPUT_DIR + "/" + outputFileName;

        // 读取Excel数据
        ExcelReader excelReader = new ExcelReader();
        Map<String, ModuleData> moduleDataMap = excelReader.readExcel(excelPath);

        // 处理Word模板
        WordProcessor wordProcessor = new WordProcessor();
        int successCount = wordProcessor.processWord(wordPath, outputPath, moduleDataMap);

        // 处理其他表格
        processAdditionalTables(excelPath, outputPath);

        // 如果使用S3存储，上传到S3
        if ("s3".equals(storageType) && s3StorageService != null) {
            String s3Key = s3StorageService.generateS3Key(outputFileName, "outputs");
            try (FileInputStream fis = new FileInputStream(outputPath)) {
                s3StorageService.uploadFile(s3Key, fis,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            }
            // 可选：上传后删除本地文件
            // Files.deleteIfExists(Paths.get(outputPath));
        }

        // 生成文档ID
        String outputId = sessionId + "_" + timestamp;

        return new ProcessResult(
                outputId,
                outputFileName,
                outputPath,
                successCount,
                "成功处理 " + successCount + " 个模块");
    }

    /**
     * 保存上传的文件
     */
    private String saveUploadedFile(InputStream inputStream, String fileName, String sessionId)
            throws IOException {
        String safeName = sessionId + "_" + fileName.replaceAll("[^a-zA-Z0-9._\\u4e00-\\u9fa5-]", "_");
        Path filePath = Paths.get(UPLOAD_DIR, safeName);
        Files.copy(inputStream, filePath, StandardCopyOption.REPLACE_EXISTING);
        return filePath.toString();
    }

    /**
     * 处理其他表格
     */
    private void processAdditionalTables(String excelPath, String outputPath) {
        try {
            TableDataReader tableReader = new TableDataReader();
            TableFillProcessor tableFillProcessor = new TableFillProcessor();

            Map<String, BasicInfoData> basicInfoMap = tableReader.readBasicInfo(excelPath);
            Map<String, ListTableData> allListData = tableReader.readAllListTableData(excelPath);

            if (!basicInfoMap.isEmpty() || !allListData.isEmpty()) {
                try (FileInputStream fis = new FileInputStream(outputPath);
                        XWPFDocument document = new XWPFDocument(fis)) {

                    tableFillProcessor.fillBasicInfoTables(document, basicInfoMap);
                    tableFillProcessor.fillListTables(document, allListData);

                    try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                        document.write(fos);
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("处理其他表格时出现警告: " + e.getMessage());
        }
    }

    /**
     * 获取处理后的文档
     */
    public byte[] getOutputDocument(String fileName) throws IOException {
        if ("s3".equals(storageType) && s3StorageService != null) {
            // 从S3获取文件
            String s3Key = s3StorageService.generateS3Key(fileName, "outputs");
            return s3StorageService.downloadFile(s3Key);
        } else {
            // 从本地文件系统获取
            Path filePath = Paths.get(OUTPUT_DIR, fileName);
            if (!Files.exists(filePath)) {
                throw new FileNotFoundException("文件不存在: " + fileName);
            }
            return Files.readAllBytes(filePath);
        }
    }

    /**
     * 根据输出ID获取文档
     */
    public byte[] getOutputDocumentById(String outputId) throws IOException {
        Path localPath = resolveLocalFileByOutputId(outputId);
        if (localPath != null && Files.exists(localPath)) {
            return Files.readAllBytes(localPath);
        }
        throw new FileNotFoundException("文档不存在: " + outputId);
    }

    /**
     * 获取所有已处理的文档列表
     */
    public List<DocumentInfo> listOutputDocuments() {
        // 从本地输出目录读取
        try (var stream = Files.list(Paths.get(OUTPUT_DIR))) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(path -> path.getFileName().toString().toLowerCase().endsWith(".docx"))
                    .sorted((a, b) -> {
                        try {
                            return Long.compare(
                                    Files.getLastModifiedTime(b).toMillis(),
                                    Files.getLastModifiedTime(a).toMillis());
                        } catch (IOException e) {
                            return 0;
                        }
                    })
                    .map(path -> {
                        String fileName = path.getFileName().toString();
                        String timestamp = extractTimestampFromFileName(fileName);
                        String id = timestamp != null ? "local_" + timestamp : fileName;
                        long size = 0L;
                        String createdAt = "";
                        try {
                            size = Files.size(path);
                            createdAt = LocalDateTime.ofInstant(
                                    Files.getLastModifiedTime(path).toInstant(),
                                    ZoneId.systemDefault()).toString();
                        } catch (IOException ignored) {
                        }
                        return new DocumentInfo(id, fileName, "", "", size, createdAt);
                    })
                    .collect(Collectors.toList());
        } catch (IOException e) {
            return List.of();
        }
    }

    /**
     * 删除文档
     */
    public boolean deleteDocument(String fileName) {
        if ("s3".equals(storageType) && s3StorageService != null) {
            // 从S3删除文件
            String s3Key = s3StorageService.generateS3Key(fileName, "outputs");
            return s3StorageService.deleteFile(s3Key);
        } else {
            // 从本地文件系统删除
            try {
                Path filePath = Paths.get(OUTPUT_DIR, fileName);
                return Files.deleteIfExists(filePath);
            } catch (IOException e) {
                return false;
            }
        }
    }

    /**
     * 根据输出ID删除文档
     */
    public boolean deleteDocumentById(String outputId) {
        // 根据 outputId 删除本地文件
        try {
            Path localPath = resolveLocalFileByOutputId(outputId);
            if (localPath != null) {
                return Files.deleteIfExists(localPath);
            }
        } catch (IOException e) {
            return false;
        }
        return false;
    }

    /**
     * 清理旧文件（保留最近N天）
     */
    public int cleanupOldFiles(int daysToKeep) {
        int deletedCount = 0;

        // 清理输出目录（本地文件）
        try {
            long cutoffTimeMillis = System.currentTimeMillis() - (daysToKeep * 24L * 60 * 60 * 1000);
            deletedCount += cleanupDirectory(Paths.get(OUTPUT_DIR), cutoffTimeMillis);
        } catch (IOException e) {
            System.err.println("清理输出目录时出错: " + e.getMessage());
        }

        // 清理上传目录（本地文件）
        try {
            long cutoffTimeMillis = System.currentTimeMillis() - (daysToKeep * 24L * 60 * 60 * 1000);
            deletedCount += cleanupDirectory(Paths.get(UPLOAD_DIR), cutoffTimeMillis);
        } catch (IOException e) {
            System.err.println("清理上传目录时出错: " + e.getMessage());
        }

        return deletedCount;
    }

    private Path resolveLocalFileByOutputId(String outputId) throws IOException {
        if (outputId == null || outputId.isBlank()) {
            return null;
        }

        String[] parts = outputId.split("_", 2);
        if (parts.length < 2) {
            return null;
        }

        String timestamp = parts[1];
        Path outputDir = Paths.get(OUTPUT_DIR);
        if (!Files.exists(outputDir)) {
            return null;
        }

        try (var stream = Files.list(outputDir)) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(path -> {
                        String name = path.getFileName().toString();
                        return name.endsWith(".docx") && name.contains("_" + timestamp + ".docx");
                    })
                    .findFirst()
                    .orElse(null);
        }
    }

    private String extractTimestampFromFileName(String fileName) {
        if (fileName == null || !fileName.endsWith(".docx")) {
            return null;
        }

        int lastUnderscore = fileName.lastIndexOf('_');
        int dot = fileName.lastIndexOf('.');
        if (lastUnderscore < 0 || dot < 0 || lastUnderscore >= dot) {
            return null;
        }

        String ts = fileName.substring(lastUnderscore + 1, dot);
        if (ts.matches("\\d{14}")) {
            return ts;
        }
        return null;
    }

    private int cleanupDirectory(Path directory, long cutoffTime) throws IOException {
        int deletedCount = 0;
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (Files.getLastModifiedTime(path).toMillis() < cutoffTime) {
                    Files.delete(path);
                    deletedCount++;
                }
            }
        }
        return deletedCount;
    }
}
