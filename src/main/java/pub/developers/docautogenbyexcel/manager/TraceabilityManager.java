package pub.developers.docautogenbyexcel.manager;

import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.model.Traceability;

import java.util.*;
import java.util.stream.Collectors;

/**
 * 需求追溯管理器
 * 负责建立和管理需求与测试用例之间的追溯关系
 */
public class TraceabilityManager {
    private Map<String, Traceability> traceabilityMap;  // 追溯关系ID -> 追溯关系对象
    private Map<String, List<Traceability>> requirementTraceMap;  // 需求ID -> 追溯关系列表
    private Map<String, List<Traceability>> testCaseTraceMap;    // 测试用例ID -> 追溯关系列表
    private RequirementManager requirementManager;      // 需求管理器（用于获取需求信息）
    
    public TraceabilityManager() {
        this.traceabilityMap = new LinkedHashMap<>();
        this.requirementTraceMap = new HashMap<>();
        this.testCaseTraceMap = new HashMap<>();
    }
    
    public TraceabilityManager(RequirementManager requirementManager) {
        this();
        this.requirementManager = requirementManager;
    }
    
    /**
     * 建立追溯关系
     * 
     * @param requirementId 需求ID
     * @param testCaseId 测试用例ID
     * @return 创建的追溯关系对象
     */
    public Traceability establishTraceability(String requirementId, String testCaseId) {
        if (requirementId == null || testCaseId == null) {
            throw new IllegalArgumentException("需求ID和测试用例ID不能为空");
        }
        
        // 检查是否已存在
        Traceability existing = findTraceability(requirementId, testCaseId);
        if (existing != null) {
            return existing;
        }
        
        // 创建新的追溯关系
        Traceability traceability = new Traceability(requirementId, testCaseId);
        
        // 填充冗余信息（便于显示）
        if (requirementManager != null) {
            Requirement req = requirementManager.getRequirement(requirementId);
            if (req != null) {
                traceability.setRequirementNumber(req.getRequirementNumber());
            }
        }
        
        // 添加到映射
        addTraceability(traceability);
        
        return traceability;
    }
    
    /**
     * 批量建立追溯关系
     */
    public List<Traceability> establishTraceabilities(Map<String, List<String>> requirementToTestCases) {
        List<Traceability> result = new ArrayList<>();
        
        for (Map.Entry<String, List<String>> entry : requirementToTestCases.entrySet()) {
            String requirementId = entry.getKey();
            for (String testCaseId : entry.getValue()) {
                Traceability trace = establishTraceability(requirementId, testCaseId);
                result.add(trace);
            }
        }
        
        return result;
    }
    
    /**
     * 添加追溯关系
     */
    public void addTraceability(Traceability traceability) {
        if (traceability == null || traceability.getTraceabilityId() == null) {
            return;
        }
        
        traceabilityMap.put(traceability.getTraceabilityId(), traceability);
        
        // 添加到需求追溯映射
        String reqId = traceability.getRequirementId();
        requirementTraceMap.computeIfAbsent(reqId, k -> new ArrayList<>()).add(traceability);
        
        // 添加到测试用例追溯映射
        String testId = traceability.getTestCaseId();
        testCaseTraceMap.computeIfAbsent(testId, k -> new ArrayList<>()).add(traceability);
    }
    
    /**
     * 查找追溯关系
     */
    public Traceability findTraceability(String requirementId, String testCaseId) {
        String traceId = "TRACE-" + requirementId + "-" + testCaseId;
        return traceabilityMap.get(traceId);
    }
    
    /**
     * 根据需求ID获取所有追溯关系
     */
    public List<Traceability> getTraceabilitiesByRequirement(String requirementId) {
        return new ArrayList<>(requirementTraceMap.getOrDefault(requirementId, new ArrayList<>()));
    }
    
    /**
     * 根据测试用例ID获取所有追溯关系
     */
    public List<Traceability> getTraceabilitiesByTestCase(String testCaseId) {
        return new ArrayList<>(testCaseTraceMap.getOrDefault(testCaseId, new ArrayList<>()));
    }
    
    /**
     * 获取需求对应的所有测试用例ID
     */
    public List<String> getTestCaseIdsByRequirement(String requirementId) {
        return getTraceabilitiesByRequirement(requirementId).stream()
                .map(Traceability::getTestCaseId)
                .distinct()
                .collect(Collectors.toList());
    }
    
    /**
     * 获取测试用例对应的所有需求ID
     */
    public List<String> getRequirementIdsByTestCase(String testCaseId) {
        return getTraceabilitiesByTestCase(testCaseId).stream()
                .map(Traceability::getRequirementId)
                .distinct()
                .collect(Collectors.toList());
    }
    
    /**
     * 删除追溯关系
     */
    public void removeTraceability(String requirementId, String testCaseId) {
        Traceability trace = findTraceability(requirementId, testCaseId);
        if (trace == null) {
            return;
        }
        
        traceabilityMap.remove(trace.getTraceabilityId());
        requirementTraceMap.getOrDefault(requirementId, new ArrayList<>()).remove(trace);
        testCaseTraceMap.getOrDefault(testCaseId, new ArrayList<>()).remove(trace);
    }
    
    /**
     * 生成追溯矩阵
     * 返回需求ID -> 测试用例ID列表的映射
     */
    public Map<String, List<String>> generateTraceMatrix() {
        Map<String, List<String>> matrix = new LinkedHashMap<>();
        
        for (Traceability trace : traceabilityMap.values()) {
            String reqId = trace.getRequirementId();
            String testId = trace.getTestCaseId();
            
            matrix.computeIfAbsent(reqId, k -> new ArrayList<>()).add(testId);
        }
        
        return matrix;
    }
    
    /**
     * 生成追溯矩阵（带详细信息）
     * 返回需求编号 -> 测试用例名称列表的映射
     */
    public Map<String, List<String>> generateTraceMatrixWithDetails() {
        Map<String, List<String>> matrix = new LinkedHashMap<>();
        
        for (Traceability trace : traceabilityMap.values()) {
            String reqNumber = trace.getRequirementNumber();
            String testName = trace.getTestCaseName();
            
            if (reqNumber != null && testName != null) {
                matrix.computeIfAbsent(reqNumber, k -> new ArrayList<>()).add(testName);
            }
        }
        
        return matrix;
    }
    
    /**
     * 验证追溯关系完整性
     * 返回验证结果报告
     */
    public TraceabilityValidationResult validateTraceability(
            Collection<Requirement> requirements, 
            Collection<TestCase> testCases) {
        
        TraceabilityValidationResult result = new TraceabilityValidationResult();
        
        // 收集所有需求ID和测试用例ID
        Set<String> requirementIds = requirements.stream()
                .map(Requirement::getRequirementId)
                .filter(Objects::nonNull)
                .collect(Collectors.toSet());
        
        Set<String> testCaseIds = testCases.stream()
                .map(TestCase::getId)
                .filter(id -> id != null && !id.trim().isEmpty())
                .collect(Collectors.toSet());
        
        // 检查没有测试用例的需求
        for (String reqId : requirementIds) {
            List<Traceability> traces = getTraceabilitiesByRequirement(reqId);
            if (traces == null || traces.isEmpty()) {
                result.addUntracedRequirement(reqId);
            }
        }
        
        // 检查没有需求关联的测试用例
        for (String testId : testCaseIds) {
            List<Traceability> traces = getTraceabilitiesByTestCase(testId);
            if (traces == null || traces.isEmpty()) {
                result.addUntracedTestCase(testId);
            }
        }
        
        // 检查无效的追溯关系（需求或测试用例不存在）
        for (Traceability trace : traceabilityMap.values()) {
            if (!requirementIds.contains(trace.getRequirementId())) {
                result.addInvalidTraceability(trace, "需求不存在");
            }
            if (!testCaseIds.contains(trace.getTestCaseId())) {
                result.addInvalidTraceability(trace, "测试用例不存在");
            }
        }
        
        return result;
    }
    
    /**
     * 计算追溯覆盖率
     */
    public TraceabilityCoverage calculateCoverage(
            Collection<Requirement> requirements, 
            Collection<TestCase> testCases) {
        
        int totalRequirements = requirements.size();
        int totalTestCases = testCases.size();
        
        long tracedRequirements = requirements.stream()
                .map(Requirement::getRequirementId)
                .filter(reqId -> {
                    List<Traceability> traces = getTraceabilitiesByRequirement(reqId);
                    return traces != null && !traces.isEmpty();
                })
                .count();
        
        long tracedTestCases = testCases.stream()
                .map(TestCase::getId)
                .filter(testId -> {
                    List<Traceability> traces = getTraceabilitiesByTestCase(testId);
                    return traces != null && !traces.isEmpty();
                })
                .count();
        
        double requirementCoverage = totalRequirements > 0 ? 
                (double) tracedRequirements / totalRequirements * 100 : 0;
        double testCaseCoverage = totalTestCases > 0 ? 
                (double) tracedTestCases / totalTestCases * 100 : 0;
        
        return new TraceabilityCoverage(
                totalRequirements, tracedRequirements, requirementCoverage,
                totalTestCases, tracedTestCases, testCaseCoverage
        );
    }
    
    /**
     * 追溯关系验证结果
     */
    public static class TraceabilityValidationResult {
        private List<String> untracedRequirements = new ArrayList<>();
        private List<String> untracedTestCases = new ArrayList<>();
        private List<Map.Entry<Traceability, String>> invalidTraceabilities = new ArrayList<>();
        
        public void addUntracedRequirement(String requirementId) {
            untracedRequirements.add(requirementId);
        }
        
        public void addUntracedTestCase(String testCaseId) {
            untracedTestCases.add(testCaseId);
        }
        
        public void addInvalidTraceability(Traceability trace, String reason) {
            invalidTraceabilities.add(new AbstractMap.SimpleEntry<>(trace, reason));
        }
        
        public boolean isValid() {
            return untracedRequirements.isEmpty() && 
                   untracedTestCases.isEmpty() && 
                   invalidTraceabilities.isEmpty();
        }
        
        public List<String> getUntracedRequirements() {
            return untracedRequirements;
        }
        
        public List<String> getUntracedTestCases() {
            return untracedTestCases;
        }
        
        public List<Map.Entry<Traceability, String>> getInvalidTraceabilities() {
            return invalidTraceabilities;
        }
        
        public void printReport() {
            System.out.println("=== 追溯关系验证报告 ===");
            System.out.println("未追溯的需求数量: " + untracedRequirements.size());
            if (!untracedRequirements.isEmpty()) {
                untracedRequirements.forEach(id -> System.out.println("  - " + id));
            }
            
            System.out.println("未追溯的测试用例数量: " + untracedTestCases.size());
            if (!untracedTestCases.isEmpty()) {
                untracedTestCases.forEach(id -> System.out.println("  - " + id));
            }
            
            System.out.println("无效的追溯关系数量: " + invalidTraceabilities.size());
            if (!invalidTraceabilities.isEmpty()) {
                invalidTraceabilities.forEach(entry -> 
                    System.out.println("  - " + entry.getKey() + ": " + entry.getValue())
                );
            }
        }
    }
    
    /**
     * 追溯覆盖率
     */
    public static class TraceabilityCoverage {
        private int totalRequirements;
        private long tracedRequirements;
        private double requirementCoverage;
        private int totalTestCases;
        private long tracedTestCases;
        private double testCaseCoverage;
        
        public TraceabilityCoverage(int totalRequirements, long tracedRequirements, double requirementCoverage,
                                   int totalTestCases, long tracedTestCases, double testCaseCoverage) {
            this.totalRequirements = totalRequirements;
            this.tracedRequirements = tracedRequirements;
            this.requirementCoverage = requirementCoverage;
            this.totalTestCases = totalTestCases;
            this.tracedTestCases = tracedTestCases;
            this.testCaseCoverage = testCaseCoverage;
        }
        
        public void printReport() {
            System.out.println("=== 追溯覆盖率报告 ===");
            System.out.println(String.format("需求覆盖率: %d/%d (%.2f%%)", 
                    tracedRequirements, totalRequirements, requirementCoverage));
            System.out.println(String.format("测试用例覆盖率: %d/%d (%.2f%%)", 
                    tracedTestCases, totalTestCases, testCaseCoverage));
        }
        
        // Getters
        public int getTotalRequirements() { return totalRequirements; }
        public long getTracedRequirements() { return tracedRequirements; }
        public double getRequirementCoverage() { return requirementCoverage; }
        public int getTotalTestCases() { return totalTestCases; }
        public long getTracedTestCases() { return tracedTestCases; }
        public double getTestCaseCoverage() { return testCaseCoverage; }
    }
}

