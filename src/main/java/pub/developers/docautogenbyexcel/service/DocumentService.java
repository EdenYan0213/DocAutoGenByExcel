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
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 文档处理服务
 * 支持本地存储和S3云存储（可选）
 */
@Service
public class DocumentService {

    private static final String STORAGE_DIR = "storage";
    private static final String UPLOAD_DIR = STORAGE_DIR + "/uploads";
    private static final String OUTPUT_DIR = STORAGE_DIR + "/outputs";
    
    private static final DateTimeFormatter TIMESTAMP_FORMAT = 
        DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

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
        String message
    ) {}

    /**
     * 已处理的文档信息
     */
    public record DocumentInfo(
        String id,
        String fileName,
        String originalExcel,
        String originalWord,
        long fileSize,
        String createdAt
    ) {}

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
            "成功处理 " + successCount + " 个模块"
        );
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
     * 获取所有已处理的文档列表
     */
    public List<DocumentInfo> listOutputDocuments() throws IOException {
        List<DocumentInfo> documents = new ArrayList<>();
        
        if ("s3".equals(storageType) && s3StorageService != null) {
            // 从S3列出文件
            // TODO: 实现S3列表功能
            // List<String> s3Keys = s3StorageService.listFiles("outputs/");
            // for (String key : s3Keys) {
            //     String fileName = key.substring(key.lastIndexOf('/') + 1);
            //     // 获取文件元数据（大小、创建时间等）
            //     // ...
            // }
            throw new UnsupportedOperationException("S3列表功能待实现");
        } else {
            // 从本地文件系统列出
            try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(OUTPUT_DIR), "*.docx")) {
                for (Path path : stream) {
                    String fileName = path.getFileName().toString();
                    long fileSize = Files.size(path);
                    String createdAt = Files.getLastModifiedTime(path).toString();
                    
                    // 从文件名解析ID
                    String id = fileName.replaceAll("\\.docx$", "");
                    
                    documents.add(new DocumentInfo(
                        id,
                        fileName,
                        "", // 原始Excel名称（简化版不保存）
                        "", // 原始Word名称
                        fileSize,
                        createdAt
                    ));
                }
            }
        }
        
        // 按创建时间倒序排列
        documents.sort((a, b) -> b.createdAt().compareTo(a.createdAt()));
        
        return documents;
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
     * 清理旧文件（保留最近7天）
     */
    public int cleanupOldFiles(int daysToKeep) {
        int deletedCount = 0;
        long cutoffTime = System.currentTimeMillis() - (daysToKeep * 24L * 60 * 60 * 1000);
        
        try {
            // 清理上传目录
            deletedCount += cleanupDirectory(Paths.get(UPLOAD_DIR), cutoffTime);
            // 清理输出目录
            deletedCount += cleanupDirectory(Paths.get(OUTPUT_DIR), cutoffTime);
        } catch (IOException e) {
            System.err.println("清理文件时出错: " + e.getMessage());
        }
        
        return deletedCount;
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

