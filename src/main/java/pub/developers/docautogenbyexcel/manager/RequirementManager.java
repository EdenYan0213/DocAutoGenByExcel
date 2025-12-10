package pub.developers.docautogenbyexcel.manager;

import pub.developers.docautogenbyexcel.model.Requirement;

import java.util.*;
import java.util.stream.Collectors;

/**
 * 需求管理器
 * 负责需求的分解、查询、管理等功能
 */
public class RequirementManager {
    private Map<String, Requirement> requirementMap;  // 需求ID -> 需求对象
    private List<Requirement> rootRequirements;        // 根需求列表（没有父需求的需求）
    private String requirementPrefix;                   // 需求编号前缀（如"REQ"）
    private int nextSequenceNumber;                     // 下一个序号
    
    public RequirementManager() {
        this.requirementMap = new LinkedHashMap<>();
        this.rootRequirements = new ArrayList<>();
        this.requirementPrefix = "REQ";
        this.nextSequenceNumber = 1;
    }
    
    public RequirementManager(String prefix) {
        this();
        this.requirementPrefix = prefix != null ? prefix : "REQ";
    }
    
    /**
     * 添加需求
     */
    public void addRequirement(Requirement requirement) {
        if (requirement == null || requirement.getRequirementId() == null) {
            return;
        }
        
        // 如果需求已存在，跳过
        if (requirementMap.containsKey(requirement.getRequirementId())) {
            return;
        }
        
        requirementMap.put(requirement.getRequirementId(), requirement);
        
        // 如果是根需求，添加到根需求列表
        if (requirement.getParentRequirementId() == null || requirement.getParentRequirementId().isEmpty()) {
            if (!rootRequirements.contains(requirement)) {
                rootRequirements.add(requirement);
            }
        } else {
            // 如果有父需求，建立父子关系
            Requirement parent = requirementMap.get(requirement.getParentRequirementId());
            if (parent != null) {
                if (!parent.getChildren().contains(requirement)) {
                    parent.addChild(requirement);
                }
            }
        }
    }
    
    /**
     * 根据ID获取需求
     */
    public Requirement getRequirement(String requirementId) {
        return requirementMap.get(requirementId);
    }
    
    /**
     * 获取所有需求
     */
    public List<Requirement> getAllRequirements() {
        return new ArrayList<>(requirementMap.values());
    }
    
    /**
     * 获取所有根需求
     */
    public List<Requirement> getRootRequirements() {
        return new ArrayList<>(rootRequirements);
    }
    
    /**
     * 获取需求的所有子需求
     */
    public List<Requirement> getChildren(String requirementId) {
        Requirement req = requirementMap.get(requirementId);
        return req != null ? new ArrayList<>(req.getChildren()) : new ArrayList<>();
    }
    
    /**
     * 获取需求的所有后代需求（包括子需求、孙需求等）
     */
    public List<Requirement> getDescendants(String requirementId) {
        Requirement req = requirementMap.get(requirementId);
        return req != null ? req.getAllDescendants() : new ArrayList<>();
    }
    
    /**
     * 分解需求：将一个需求分解为多个子需求
     * 
     * @param parentRequirementId 父需求ID
     * @param childRequirements 子需求列表
     * @return 分解后的子需求列表
     */
    public List<Requirement> decomposeRequirement(String parentRequirementId, List<Requirement> childRequirements) {
        Requirement parent = requirementMap.get(parentRequirementId);
        if (parent == null) {
            throw new IllegalArgumentException("父需求不存在: " + parentRequirementId);
        }
        
        List<Requirement> addedChildren = new ArrayList<>();
        
        // 先获取当前已有的子需求数量，用于计算起始索引
        List<Requirement> existingChildren = getChildren(parentRequirementId);
        int startIndex = existingChildren.size() + 1;
        
        for (int i = 0; i < childRequirements.size(); i++) {
            Requirement child = childRequirements.get(i);
            
            // 如果子需求没有ID，自动生成
            if (child.getRequirementId() == null || child.getRequirementId().isEmpty()) {
                child.setRequirementId(generateRequirementId());
            }
            
            // 设置父需求关系
            child.setParentRequirementId(parentRequirementId);
            
            // 自动生成需求编号（基于父需求编号和索引）
            if (child.getRequirementNumber() == null || child.getRequirementNumber().isEmpty()) {
                child.setRequirementNumber(generateChildRequirementNumber(parent.getRequirementNumber(), startIndex + i));
            }
            
            // 添加到管理器（先添加到map，但不建立父子关系，避免重复）
            requirementMap.put(child.getRequirementId(), child);
            
            // 添加到父需求的子需求列表（这会建立父子关系）
            if (!parent.getChildren().contains(child)) {
                parent.addChild(child);
            }
            
            // 如果是根需求，添加到根需求列表
            if (child.getParentRequirementId() == null || child.getParentRequirementId().isEmpty()) {
                if (!rootRequirements.contains(child)) {
                    rootRequirements.add(child);
                }
            } else {
                // 从根需求列表中移除（如果存在）
                rootRequirements.remove(child);
            }
            
            addedChildren.add(child);
        }
        
        return addedChildren;
    }
    
    /**
     * 自动分解需求：根据分解规则自动创建子需求
     * 
     * @param parentRequirementId 父需求ID
     * @param childNames 子需求名称列表
     * @return 分解后的子需求列表
     */
    public List<Requirement> autoDecomposeRequirement(String parentRequirementId, List<String> childNames) {
        Requirement parent = requirementMap.get(parentRequirementId);
        if (parent == null) {
            throw new IllegalArgumentException("父需求不存在: " + parentRequirementId);
        }
        
        List<Requirement> children = new ArrayList<>();
        // 先获取当前已有的子需求数量，用于计算起始索引
        List<Requirement> existingChildren = getChildren(parentRequirementId);
        int startIndex = existingChildren.size() + 1;
        
        for (int i = 0; i < childNames.size(); i++) {
            Requirement child = new Requirement();
            child.setRequirementId(generateRequirementId());
            child.setRequirementName(childNames.get(i));
            child.setRequirementNumber(generateChildRequirementNumber(parent.getRequirementNumber(), startIndex + i));
            child.setParentRequirementId(parentRequirementId);
            child.setType(parent.getType());
            child.setPriority(parent.getPriority());
            children.add(child);
        }
        
        return decomposeRequirement(parentRequirementId, children);
    }
    
    /**
     * 生成需求ID（自动编号）
     */
    public String generateRequirementId() {
        // 确保生成的ID不会与现有需求冲突
        String id;
        do {
            id = requirementPrefix + "-" + String.format("%03d", nextSequenceNumber);
            nextSequenceNumber++;
        } while (requirementMap.containsKey(id));
        return id;
    }
    
    /**
     * 生成子需求编号（基于父需求编号）
     * 例如：父需求REQ-001，子需求为REQ-001.1, REQ-001.2等
     */
    private String generateChildRequirementNumber(String parentNumber, int currentIndex) {
        return parentNumber + "." + currentIndex;
    }
    
    /**
     * 生成子需求编号（基于父需求编号，自动计算下一个索引）
     */
    private String generateChildRequirementNumber(String parentNumber) {
        Requirement parent = findRequirementByNumber(parentNumber);
        if (parent == null) {
            return parentNumber + ".1";
        }
        
        // 获取所有已存在的子需求编号，找出最大的索引
        List<Requirement> siblings = getChildren(parent.getRequirementId());
        int maxIndex = 0;
        for (Requirement sibling : siblings) {
            String siblingNumber = sibling.getRequirementNumber();
            if (siblingNumber != null && siblingNumber.startsWith(parentNumber + ".")) {
                String indexPart = siblingNumber.substring(parentNumber.length() + 1);
                try {
                    // 只取第一级索引（如"1.1"中的"1"）
                    String firstLevel = indexPart.split("\\.")[0];
                    int index = Integer.parseInt(firstLevel);
                    maxIndex = Math.max(maxIndex, index);
                } catch (NumberFormatException e) {
                    // 忽略无法解析的编号
                }
            }
        }
        
        int nextIndex = maxIndex + 1;
        return parentNumber + "." + nextIndex;
    }
    
    /**
     * 根据需求编号查找需求
     */
    public Requirement findRequirementByNumber(String requirementNumber) {
        return requirementMap.values().stream()
                .filter(req -> requirementNumber.equals(req.getRequirementNumber()))
                .findFirst()
                .orElse(null);
    }
    
    /**
     * 获取需求的层级深度
     */
    public int getRequirementLevel(String requirementId) {
        Requirement req = requirementMap.get(requirementId);
        if (req == null) {
            return -1;
        }
        
        int level = 0;
        String parentId = req.getParentRequirementId();
        while (parentId != null && !parentId.isEmpty()) {
            Requirement parent = requirementMap.get(parentId);
            if (parent == null) {
                break;
            }
            level++;
            parentId = parent.getParentRequirementId();
        }
        
        return level;
    }
    
    /**
     * 获取需求树（以树形结构组织）
     */
    public RequirementTree getRequirementTree() {
        // 确保根需求列表是最新的（只包含没有父需求的需求）
        List<Requirement> actualRoots = new ArrayList<>();
        for (Requirement req : requirementMap.values()) {
            if (req.getParentRequirementId() == null || req.getParentRequirementId().isEmpty()) {
                if (!actualRoots.contains(req)) {
                    actualRoots.add(req);
                }
            }
        }
        return new RequirementTree(actualRoots);
    }
    
    /**
     * 搜索需求（根据名称、编号、描述等）
     */
    public List<Requirement> searchRequirements(String keyword) {
        if (keyword == null || keyword.trim().isEmpty()) {
            return getAllRequirements();
        }
        
        String lowerKeyword = keyword.toLowerCase();
        return requirementMap.values().stream()
                .filter(req -> 
                    (req.getRequirementName() != null && req.getRequirementName().toLowerCase().contains(lowerKeyword)) ||
                    (req.getRequirementNumber() != null && req.getRequirementNumber().toLowerCase().contains(lowerKeyword)) ||
                    (req.getDescription() != null && req.getDescription().toLowerCase().contains(lowerKeyword)) ||
                    (req.getRequirementId() != null && req.getRequirementId().toLowerCase().contains(lowerKeyword))
                )
                .collect(Collectors.toList());
    }
    
    /**
     * 获取叶子需求列表（没有子需求的需求）
     */
    public List<Requirement> getLeafRequirements() {
        return requirementMap.values().stream()
                .filter(Requirement::isLeaf)
                .collect(Collectors.toList());
    }
    
    /**
     * 删除需求（同时删除所有子需求）
     */
    public void removeRequirement(String requirementId) {
        Requirement req = requirementMap.get(requirementId);
        if (req == null) {
            return;
        }
        
        // 递归删除所有子需求
        List<Requirement> descendants = req.getAllDescendants();
        for (Requirement descendant : descendants) {
            requirementMap.remove(descendant.getRequirementId());
        }
        
        // 删除当前需求
        requirementMap.remove(requirementId);
        
        // 从父需求的子需求列表中移除
        if (req.getParentRequirementId() != null && !req.getParentRequirementId().isEmpty()) {
            Requirement parent = requirementMap.get(req.getParentRequirementId());
            if (parent != null) {
                parent.removeChild(req);
            }
        } else {
            // 从根需求列表中移除
            rootRequirements.remove(req);
        }
    }
    
    /**
     * 需求树结构（用于可视化）
     */
    public static class RequirementTree {
        private List<Requirement> roots;
        
        public RequirementTree(List<Requirement> roots) {
            this.roots = roots;
        }
        
        public List<Requirement> getRoots() {
            return roots;
        }
        
        /**
         * 打印需求树（用于调试）
         */
        public void printTree() {
            for (Requirement root : roots) {
                printNode(root, 0);
            }
        }
        
        private void printNode(Requirement req, int indent) {
            printNode(req, indent, new HashSet<>());
        }
        
        private void printNode(Requirement req, int indent, Set<String> visited) {
            // 防止循环引用
            if (visited.contains(req.getRequirementId())) {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < indent; i++) {
                    sb.append("  ");
                }
                sb.append(req.getRequirementNumber()).append(" ").append(req.getRequirementName())
                  .append(" [循环引用，已跳过]");
                System.out.println(sb.toString());
                return;
            }
            
            visited.add(req.getRequirementId());
            
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < indent; i++) {
                sb.append("  ");
            }
            sb.append(req.getRequirementNumber()).append(" ").append(req.getRequirementName());
            System.out.println(sb.toString());
            
            for (Requirement child : req.getChildren()) {
                printNode(child, indent + 1, visited);
            }
            
            visited.remove(req.getRequirementId());
        }
    }
}

