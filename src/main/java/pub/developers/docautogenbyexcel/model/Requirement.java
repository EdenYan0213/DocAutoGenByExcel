package pub.developers.docautogenbyexcel.model;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 需求数据模型
 * 支持需求分解和层级管理
 */
public class Requirement {
    private String requirementId;          // 需求唯一标识（如REQ-001）
    private String requirementNumber;      // 需求编号（自动生成，如REQ-001.1）
    private String requirementName;        // 需求名称
    private String description;            // 需求描述
    private RequirementType type;          // 需求类型
    private Priority priority;            // 优先级
    private String parentRequirementId;   // 父需求ID（用于需求分解）
    private List<Requirement> children;   // 子需求列表
    private RequirementStatus status;     // 需求状态
    private Map<String, String> attributes; // 扩展属性（动态字段）
    
    // 需求类型枚举
    public enum RequirementType {
        FUNCTIONAL("功能需求"),
        PERFORMANCE("性能需求"),
        INTERFACE("接口需求"),
        SECURITY("安全需求"),
        RELIABILITY("可靠性需求"),
        MAINTAINABILITY("可维护性需求"),
        OTHER("其他需求");
        
        private final String description;
        
        RequirementType(String description) {
            this.description = description;
        }
        
        public String getDescription() {
            return description;
        }
    }
    
    // 优先级枚举
    public enum Priority {
        HIGH("高"),
        MEDIUM("中"),
        LOW("低");
        
        private final String description;
        
        Priority(String description) {
            this.description = description;
        }
        
        public String getDescription() {
            return description;
        }
    }
    
    // 需求状态枚举
    public enum RequirementStatus {
        DRAFT("草稿"),
        REVIEWING("评审中"),
        APPROVED("已批准"),
        IMPLEMENTING("实现中"),
        IMPLEMENTED("已实现"),
        VERIFIED("已验证"),
        REJECTED("已拒绝");
        
        private final String description;
        
        RequirementStatus(String description) {
            this.description = description;
        }
        
        public String getDescription() {
            return description;
        }
    }
    
    public Requirement() {
        this.children = new ArrayList<>();
        this.attributes = new LinkedHashMap<>();
        this.status = RequirementStatus.DRAFT;
        this.priority = Priority.MEDIUM;
        this.type = RequirementType.FUNCTIONAL;
    }
    
    public Requirement(String requirementId, String requirementName) {
        this();
        this.requirementId = requirementId;
        this.requirementNumber = requirementId;
        this.requirementName = requirementName;
    }
    
    /**
     * 添加子需求
     */
    public void addChild(Requirement child) {
        if (child != null && !children.contains(child)) {
            // 防止循环引用：检查子需求是否是当前需求的祖先
            if (isAncestor(child)) {
                throw new IllegalArgumentException("不能添加循环引用：子需求是当前需求的祖先");
            }
            children.add(child);
            child.setParentRequirementId(this.requirementId);
        }
    }
    
    /**
     * 检查给定的需求是否是当前需求的祖先
     */
    private boolean isAncestor(Requirement potentialAncestor) {
        if (potentialAncestor == null) {
            return false;
        }
        Requirement current = this;
        while (current.getParentRequirementId() != null && !current.getParentRequirementId().isEmpty()) {
            // 这里需要外部管理器来获取父需求，所以暂时只检查直接父需求
            // 完整的检查需要在RequirementManager中实现
            if (current.getParentRequirementId().equals(potentialAncestor.getRequirementId())) {
                return true;
            }
            // 由于没有父需求的引用，这里只能检查直接父需求
            break;
        }
        return false;
    }
    
    /**
     * 移除子需求
     */
    public void removeChild(Requirement child) {
        if (child != null) {
            children.remove(child);
            child.setParentRequirementId(null);
        }
    }
    
    /**
     * 检查是否有子需求
     */
    public boolean hasChildren() {
        return children != null && !children.isEmpty();
    }
    
    /**
     * 获取需求层级深度（根需求为0）
     */
    public int getLevel() {
        if (parentRequirementId == null || parentRequirementId.isEmpty()) {
            return 0;
        }
        // 这里需要外部管理器来计算，因为需要访问父需求
        return -1; // 表示需要外部计算
    }
    
    /**
     * 获取所有后代需求（包括子需求、孙需求等）
     */
    public List<Requirement> getAllDescendants() {
        List<Requirement> descendants = new ArrayList<>();
        for (Requirement child : children) {
            descendants.add(child);
            descendants.addAll(child.getAllDescendants());
        }
        return descendants;
    }
    
    /**
     * 检查是否为叶子需求（没有子需求）
     */
    public boolean isLeaf() {
        return !hasChildren();
    }
    
    /**
     * 添加扩展属性
     */
    public void addAttribute(String key, String value) {
        if (attributes == null) {
            attributes = new LinkedHashMap<>();
        }
        attributes.put(key, value != null ? value : "");
    }
    
    /**
     * 获取扩展属性
     */
    public String getAttribute(String key) {
        return attributes != null ? attributes.getOrDefault(key, "") : "";
    }
    
    // Getters and Setters
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
    
    public String getRequirementName() {
        return requirementName;
    }
    
    public void setRequirementName(String requirementName) {
        this.requirementName = requirementName;
    }
    
    public String getDescription() {
        return description;
    }
    
    public void setDescription(String description) {
        this.description = description;
    }
    
    public RequirementType getType() {
        return type;
    }
    
    public void setType(RequirementType type) {
        this.type = type;
    }
    
    public Priority getPriority() {
        return priority;
    }
    
    public void setPriority(Priority priority) {
        this.priority = priority;
    }
    
    public String getParentRequirementId() {
        return parentRequirementId;
    }
    
    public void setParentRequirementId(String parentRequirementId) {
        this.parentRequirementId = parentRequirementId;
    }
    
    public List<Requirement> getChildren() {
        return children;
    }
    
    public void setChildren(List<Requirement> children) {
        this.children = children;
    }
    
    public RequirementStatus getStatus() {
        return status;
    }
    
    public void setStatus(RequirementStatus status) {
        this.status = status;
    }
    
    public Map<String, String> getAttributes() {
        return attributes;
    }
    
    public void setAttributes(Map<String, String> attributes) {
        this.attributes = attributes;
    }
    
    @Override
    public String toString() {
        return "Requirement{" +
                "requirementId='" + requirementId + '\'' +
                ", requirementNumber='" + requirementNumber + '\'' +
                ", requirementName='" + requirementName + '\'' +
                ", type=" + type +
                ", priority=" + priority +
                ", status=" + status +
                ", childrenCount=" + (children != null ? children.size() : 0) +
                '}';
    }
    
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Requirement that = (Requirement) o;
        return requirementId != null && requirementId.equals(that.requirementId);
    }
    
    @Override
    public int hashCode() {
        return requirementId != null ? requirementId.hashCode() : 0;
    }
}

