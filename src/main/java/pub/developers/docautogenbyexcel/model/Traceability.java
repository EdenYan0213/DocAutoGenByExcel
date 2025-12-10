package pub.developers.docautogenbyexcel.model;

import java.util.Date;

/**
 * 需求追溯关系模型
 * 建立需求与测试用例之间的追溯关系
 */
public class Traceability {
    private String traceabilityId;        // 追溯关系唯一标识
    private String requirementId;         // 需求ID
    private String requirementNumber;      // 需求编号（冗余，便于显示）
    private String testCaseId;             // 测试用例ID
    private String testCaseName;          // 测试用例名称（冗余，便于显示）
    private TraceType traceType;          // 追溯类型
    private TraceStatus status;           // 追溯状态
    private String description;            // 追溯关系描述
    private Date createTime;              // 创建时间
    private Date updateTime;              // 更新时间
    
    // 追溯类型枚举
    public enum TraceType {
        FORWARD("正向追溯"),      // 从需求到测试用例
        BACKWARD("反向追溯"),     // 从测试用例到需求
        BIDIRECTIONAL("双向追溯"); // 双向追溯
        
        private final String description;
        
        TraceType(String description) {
            this.description = description;
        }
        
        public String getDescription() {
            return description;
        }
    }
    
    // 追溯状态枚举
    public enum TraceStatus {
        ESTABLISHED("已建立"),    // 追溯关系已建立
        VERIFIED("已验证"),      // 追溯关系已验证
        INVALID("无效"),         // 追溯关系无效
        PENDING("待验证");       // 待验证
        
        private final String description;
        
        TraceStatus(String description) {
            this.description = description;
        }
        
        public String getDescription() {
            return description;
        }
    }
    
    public Traceability() {
        this.traceType = TraceType.FORWARD;
        this.status = TraceStatus.ESTABLISHED;
        this.createTime = new Date();
        this.updateTime = new Date();
    }
    
    public Traceability(String requirementId, String testCaseId) {
        this();
        this.requirementId = requirementId;
        this.testCaseId = testCaseId;
        this.traceabilityId = generateTraceabilityId(requirementId, testCaseId);
    }
    
    /**
     * 生成追溯关系ID
     */
    private String generateTraceabilityId(String reqId, String testId) {
        return "TRACE-" + reqId + "-" + testId;
    }
    
    // Getters and Setters
    public String getTraceabilityId() {
        return traceabilityId;
    }
    
    public void setTraceabilityId(String traceabilityId) {
        this.traceabilityId = traceabilityId;
    }
    
    public String getRequirementId() {
        return requirementId;
    }
    
    public void setRequirementId(String requirementId) {
        this.requirementId = requirementId;
    }
    
    public String getRequirementNumber() {
        return requirementNumber;
    }
    
    public void setRequirementNumber(String requirementNumber) {
        this.requirementNumber = requirementNumber;
    }
    
    public String getTestCaseId() {
        return testCaseId;
    }
    
    public void setTestCaseId(String testCaseId) {
        this.testCaseId = testCaseId;
    }
    
    public String getTestCaseName() {
        return testCaseName;
    }
    
    public void setTestCaseName(String testCaseName) {
        this.testCaseName = testCaseName;
    }
    
    public TraceType getTraceType() {
        return traceType;
    }
    
    public void setTraceType(TraceType traceType) {
        this.traceType = traceType;
    }
    
    public TraceStatus getStatus() {
        return status;
    }
    
    public void setStatus(TraceStatus status) {
        this.status = status;
    }
    
    public String getDescription() {
        return description;
    }
    
    public void setDescription(String description) {
        this.description = description;
    }
    
    public Date getCreateTime() {
        return createTime;
    }
    
    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }
    
    public Date getUpdateTime() {
        return updateTime;
    }
    
    public void setUpdateTime(Date updateTime) {
        this.updateTime = updateTime;
    }
    
    @Override
    public String toString() {
        return "Traceability{" +
                "traceabilityId='" + traceabilityId + '\'' +
                ", requirementId='" + requirementId + '\'' +
                ", testCaseId='" + testCaseId + '\'' +
                ", traceType=" + traceType +
                ", status=" + status +
                '}';
    }
    
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Traceability that = (Traceability) o;
        return traceabilityId != null && traceabilityId.equals(that.traceabilityId);
    }
    
    @Override
    public int hashCode() {
        return traceabilityId != null ? traceabilityId.hashCode() : 0;
    }
}

