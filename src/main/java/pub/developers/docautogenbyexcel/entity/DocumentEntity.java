package pub.developers.docautogenbyexcel.entity;

import jakarta.persistence.*;
import java.time.LocalDateTime;

/**
 * 文档实体类
 * 用于存储处理后的文档信息
 */
@Entity
@Table(name = "documents", indexes = {
    @Index(name = "idx_created_at", columnList = "createdAt"),
    @Index(name = "idx_output_id", columnList = "outputId")
})
public class DocumentEntity {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    /**
     * 输出文档ID（业务标识）
     */
    @Column(nullable = false, unique = true, length = 100)
    private String outputId;

    /**
     * 输出文件名
     */
    @Column(nullable = false, length = 500)
    private String outputFileName;

    /**
     * 原始Excel文件名
     */
    @Column(length = 500)
    private String originalExcelName;

    /**
     * 原始Word文件名
     */
    @Column(length = 500)
    private String originalWordName;

    /**
     * 文档内容（BLOB）
     */
    @Lob
    @Column(nullable = false, columnDefinition = "BLOB")
    private byte[] content;

    /**
     * 文件大小（字节）
     */
    @Column(nullable = false)
    private Long fileSize;

    /**
     * 处理的模块数量
     */
    @Column(nullable = false)
    private Integer moduleCount;

    /**
     * 处理消息
     */
    @Column(length = 1000)
    private String message;

    /**
     * 创建时间
     */
    @Column(nullable = false, updatable = false)
    private LocalDateTime createdAt;

    /**
     * 更新时间
     */
    @Column(nullable = false)
    private LocalDateTime updatedAt;

    /**
     * 本地文件路径（如果使用本地存储）
     */
    @Column(length = 1000)
    private String localFilePath;

    /**
     * S3存储路径（如果使用S3存储）
     */
    @Column(length = 1000)
    private String s3Key;

    /**
     * 存储类型：local, s3, database
     */
    @Column(length = 20)
    private String storageType;

    @PrePersist
    protected void onCreate() {
        createdAt = LocalDateTime.now();
        updatedAt = LocalDateTime.now();
        if (storageType == null) {
            storageType = "database";
        }
    }

    @PreUpdate
    protected void onUpdate() {
        updatedAt = LocalDateTime.now();
    }

    // Getters and Setters

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getOutputId() {
        return outputId;
    }

    public void setOutputId(String outputId) {
        this.outputId = outputId;
    }

    public String getOutputFileName() {
        return outputFileName;
    }

    public void setOutputFileName(String outputFileName) {
        this.outputFileName = outputFileName;
    }

    public String getOriginalExcelName() {
        return originalExcelName;
    }

    public void setOriginalExcelName(String originalExcelName) {
        this.originalExcelName = originalExcelName;
    }

    public String getOriginalWordName() {
        return originalWordName;
    }

    public void setOriginalWordName(String originalWordName) {
        this.originalWordName = originalWordName;
    }

    public byte[] getContent() {
        return content;
    }

    public void setContent(byte[] content) {
        this.content = content;
    }

    public Long getFileSize() {
        return fileSize;
    }

    public void setFileSize(Long fileSize) {
        this.fileSize = fileSize;
    }

    public Integer getModuleCount() {
        return moduleCount;
    }

    public void setModuleCount(Integer moduleCount) {
        this.moduleCount = moduleCount;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public LocalDateTime getCreatedAt() {
        return createdAt;
    }

    public void setCreatedAt(LocalDateTime createdAt) {
        this.createdAt = createdAt;
    }

    public LocalDateTime getUpdatedAt() {
        return updatedAt;
    }

    public void setUpdatedAt(LocalDateTime updatedAt) {
        this.updatedAt = updatedAt;
    }

    public String getLocalFilePath() {
        return localFilePath;
    }

    public void setLocalFilePath(String localFilePath) {
        this.localFilePath = localFilePath;
    }

    public String getS3Key() {
        return s3Key;
    }

    public void setS3Key(String s3Key) {
        this.s3Key = s3Key;
    }

    public String getStorageType() {
        return storageType;
    }

    public void setStorageType(String storageType) {
        this.storageType = storageType;
    }
}

