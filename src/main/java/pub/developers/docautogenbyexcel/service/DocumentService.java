package pub.developers.docautogenbyexcel.service;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import pub.developers.docautogenbyexcel.entity.DocumentEntity;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.processor.TableFillProcessor;
import pub.developers.docautogenbyexcel.processor.WordProcessor;
import pub.developers.docautogenbyexcel.reader.ExcelReader;
import pub.developers.docautogenbyexcel.reader.TableDataReader;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;
import pub.developers.docautogenbyexcel.repository.DocumentRepository;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 文档处理服务
 * 支持本地存储、S3云存储和数据库存储
 */
@Service
public class DocumentService {

    private static final String STORAGE_DIR = "storage";
    private static final String UPLOAD_DIR = STORAGE_DIR + "/uploads";
    private static final String OUTPUT_DIR = STORAGE_DIR + "/outputs";
    
    private static final DateTimeFormatter TIMESTAMP_FORMAT = 
        DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

    @Value("${storage.type:database}")
    private String storageType;

    @Value("${storage.save-to-database:true}")
    private boolean saveToDatabase;

    @Autowired(required = false)
    private S3StorageService s3StorageService;

    @Autowired
    private DocumentRepository documentRepository;

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
        
        // 读取生成的文档内容
        byte[] documentContent = Files.readAllBytes(Paths.get(outputPath));
        
        // 如果使用S3存储，上传到S3
        String s3Key = null;
        if ("s3".equals(storageType) && s3StorageService != null) {
            s3Key = s3StorageService.generateS3Key(outputFileName, "outputs");
            try (FileInputStream fis = new FileInputStream(outputPath)) {
                s3StorageService.uploadFile(s3Key, fis, 
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            }
            // 可选：上传后删除本地文件
            // Files.deleteIfExists(Paths.get(outputPath));
        }
        
        // 生成文档ID
        String outputId = sessionId + "_" + timestamp;
        
        // 保存到数据库
        if (saveToDatabase) {
            saveDocumentToDatabase(outputId, outputFileName, excelFileName, wordFileName,
                documentContent, outputPath, s3Key, successCount);
        }
        
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
        // 优先从数据库获取
        Optional<DocumentEntity> docEntity = documentRepository.findByOutputFileName(fileName);
        if (docEntity.isPresent()) {
            DocumentEntity entity = docEntity.get();
            // 如果存储类型是database，直接从数据库读取
            if ("database".equals(entity.getStorageType())) {
                return entity.getContent();
            }
            // 如果存储类型是s3，从S3获取
            if ("s3".equals(entity.getStorageType()) && s3StorageService != null && entity.getS3Key() != null) {
                return s3StorageService.downloadFile(entity.getS3Key());
            }
            // 如果存储类型是local，从本地文件系统获取
            if ("local".equals(entity.getStorageType()) && entity.getLocalFilePath() != null) {
                Path filePath = Paths.get(entity.getLocalFilePath());
                if (Files.exists(filePath)) {
                    return Files.readAllBytes(filePath);
                }
            }
        }
        
        // 如果数据库中没有，尝试从文件系统获取（向后兼容）
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
        Optional<DocumentEntity> docEntity = documentRepository.findByOutputId(outputId);
        if (docEntity.isPresent()) {
            DocumentEntity entity = docEntity.get();
            if ("database".equals(entity.getStorageType())) {
                return entity.getContent();
            }
            if ("s3".equals(entity.getStorageType()) && s3StorageService != null && entity.getS3Key() != null) {
                return s3StorageService.downloadFile(entity.getS3Key());
            }
            if ("local".equals(entity.getStorageType()) && entity.getLocalFilePath() != null) {
                Path filePath = Paths.get(entity.getLocalFilePath());
                if (Files.exists(filePath)) {
                    return Files.readAllBytes(filePath);
                }
            }
        }
        throw new FileNotFoundException("文档不存在: " + outputId);
    }

    /**
     * 获取所有已处理的文档列表
     */
    public List<DocumentInfo> listOutputDocuments() {
        // 优先从数据库查询
        List<DocumentEntity> entities = documentRepository.findAllByOrderByCreatedAtDesc();
        
        return entities.stream()
            .map(entity -> new DocumentInfo(
                entity.getOutputId(),
                entity.getOutputFileName(),
                entity.getOriginalExcelName() != null ? entity.getOriginalExcelName() : "",
                entity.getOriginalWordName() != null ? entity.getOriginalWordName() : "",
                entity.getFileSize(),
                entity.getCreatedAt().toString()
            ))
            .collect(Collectors.toList());
    }

    /**
     * 保存文档到数据库
     */
    @Transactional
    private void saveDocumentToDatabase(String outputId, String outputFileName,
                                        String originalExcelName, String originalWordName,
                                        byte[] content, String localFilePath, String s3Key,
                                        int moduleCount) {
        DocumentEntity entity = new DocumentEntity();
        entity.setOutputId(outputId);
        entity.setOutputFileName(outputFileName);
        entity.setOriginalExcelName(originalExcelName);
        entity.setOriginalWordName(originalWordName);
        entity.setContent(content);
        entity.setFileSize((long) content.length);
        entity.setModuleCount(moduleCount);
        entity.setMessage("成功处理 " + moduleCount + " 个模块");
        entity.setLocalFilePath(localFilePath);
        entity.setS3Key(s3Key);
        entity.setStorageType(storageType);
        
        documentRepository.save(entity);
    }

    /**
     * 删除文档
     */
    @Transactional
    public boolean deleteDocument(String fileName) {
        // 从数据库查找并删除
        Optional<DocumentEntity> docEntity = documentRepository.findByOutputFileName(fileName);
        if (docEntity.isPresent()) {
            DocumentEntity entity = docEntity.get();
            
            // 如果存储类型是s3，同时从S3删除
            if ("s3".equals(entity.getStorageType()) && s3StorageService != null && entity.getS3Key() != null) {
                s3StorageService.deleteFile(entity.getS3Key());
            }
            
            // 如果存储类型是local，同时从本地文件系统删除
            if ("local".equals(entity.getStorageType()) && entity.getLocalFilePath() != null) {
                try {
                    Files.deleteIfExists(Paths.get(entity.getLocalFilePath()));
                } catch (IOException e) {
                    // 忽略文件删除错误，继续删除数据库记录
                }
            }
            
            // 从数据库删除
            documentRepository.delete(entity);
            return true;
        }
        
        // 如果数据库中没有，尝试从文件系统删除（向后兼容）
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
    @Transactional
    public boolean deleteDocumentById(String outputId) {
        Optional<DocumentEntity> docEntity = documentRepository.findByOutputId(outputId);
        if (docEntity.isPresent()) {
            DocumentEntity entity = docEntity.get();
            
            // 如果存储类型是s3，同时从S3删除
            if ("s3".equals(entity.getStorageType()) && s3StorageService != null && entity.getS3Key() != null) {
                s3StorageService.deleteFile(entity.getS3Key());
            }
            
            // 如果存储类型是local，同时从本地文件系统删除
            if ("local".equals(entity.getStorageType()) && entity.getLocalFilePath() != null) {
                try {
                    Files.deleteIfExists(Paths.get(entity.getLocalFilePath()));
                } catch (IOException e) {
                    // 忽略文件删除错误
                }
            }
            
            // 从数据库删除
            documentRepository.delete(entity);
            return true;
        }
        return false;
    }

    /**
     * 清理旧文件（保留最近N天）
     */
    @Transactional
    public int cleanupOldFiles(int daysToKeep) {
        int deletedCount = 0;
        LocalDateTime cutoffTime = LocalDateTime.now().minusDays(daysToKeep);
        
        // 从数据库查找旧文档
        List<DocumentEntity> oldDocuments = documentRepository.findByCreatedAtBetween(
            LocalDateTime.of(2000, 1, 1, 0, 0), cutoffTime);
        
        for (DocumentEntity entity : oldDocuments) {
            // 如果存储类型是s3，同时从S3删除
            if ("s3".equals(entity.getStorageType()) && s3StorageService != null && entity.getS3Key() != null) {
                s3StorageService.deleteFile(entity.getS3Key());
            }
            
            // 如果存储类型是local，同时从本地文件系统删除
            if ("local".equals(entity.getStorageType()) && entity.getLocalFilePath() != null) {
                try {
                    Files.deleteIfExists(Paths.get(entity.getLocalFilePath()));
                } catch (IOException e) {
                    // 忽略文件删除错误
                }
            }
            
            // 从数据库删除
            documentRepository.delete(entity);
            deletedCount++;
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

