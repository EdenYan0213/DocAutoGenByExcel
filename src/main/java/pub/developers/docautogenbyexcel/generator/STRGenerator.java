package pub.developers.docautogenbyexcel.generator;

import pub.developers.docautogenbyexcel.hub.DataHub;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.model.TestResult;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;

import java.util.*;

/**
 * 文档生成层：软件测试报告(STR)生成器
 * 继承AbstractDocumentGenerator，遵循extractData() → generateContent() → save()模板方法
 *
 * 生成《软件测试报告》文档及统计分析
 */
public class STRGenerator extends AbstractDocumentGenerator {

  public STRGenerator(DataHub dataHub) {
    super(dataHub);
  }

  /**
   * STR生成器的数据容器
   * 封装从Excel提取的所有数据
   */
  public static class STRExtractedData implements ExtractedData {
    private final Map<String, ModuleData> moduleDataMap;
    private final List<TestResult> testResults;
    private final Map<String, BasicInfoData> basicInfoMap;
    private final Map<String, ListTableData> listTableMap;
    private final List<Requirement> requirements;
    private final List<TestCase> testCases;

    public STRExtractedData(
        Map<String, ModuleData> moduleDataMap,
        List<TestResult> testResults,
        Map<String, BasicInfoData> basicInfoMap,
        Map<String, ListTableData> listTableMap,
        List<Requirement> requirements,
        List<TestCase> testCases) {
      this.moduleDataMap = moduleDataMap;
      this.testResults = testResults;
      this.basicInfoMap = basicInfoMap;
      this.listTableMap = listTableMap;
      this.requirements = requirements;
      this.testCases = testCases;
    }

    public Map<String, ModuleData> getModuleDataMap() {
      return moduleDataMap;
    }

    public Map<String, BasicInfoData> getBasicInfoMap() {
      return basicInfoMap;
    }

    public List<TestResult> getTestResults() {
      return testResults;
    }

    public Map<String, ListTableData> getListTableMap() {
      return listTableMap;
    }

    public List<Requirement> getRequirements() {
      return requirements;
    }

    public List<TestCase> getTestCases() {
      return testCases;
    }
  }

  @Override
  protected STRExtractedData extractData(String excelPath) throws Exception {
    Map<String, ModuleData> moduleDataMap = dataHub.loadModuleData(excelPath);
    List<TestResult> testResults = dataHub.loadTestResults(excelPath);
    Map<String, BasicInfoData> basicInfoMap = dataHub.loadBasicInfo(excelPath);
    Map<String, ListTableData> listTableMap = dataHub.loadListTables(excelPath);
    List<Requirement> requirements = dataHub.loadRequirements(excelPath);
    List<TestCase> testCases = dataHub.loadTestCases(excelPath);
    return new STRExtractedData(moduleDataMap, testResults, basicInfoMap, listTableMap, requirements, testCases);
  }

  @Override
  protected int generateContent(String templatePath, String outputPath, ExtractedData extractedData)
      throws Exception {
    STRExtractedData strData = (STRExtractedData) extractedData;
    StrStatistics stats = analyzeStatistics(strData.getTestResults());
    System.out.println("STR统计：总结果=" + stats.total() + "，通过=" + stats.passed()
        + "，失败=" + stats.failed() + "，其他=" + stats.others());
    int moduleCount = wordDocumentBuilder.buildModuleSections(templatePath, outputPath, strData.getModuleDataMap());
    return moduleCount;
  }

  @Override
  protected GenerateResult save(String outputPath, ExtractedData extractedData, int contentResult) throws Exception {
    STRExtractedData strData = (STRExtractedData) extractedData;
    wordDocumentBuilder.fillAdditionalTables(outputPath, strData.getBasicInfoMap(), strData.getListTableMap(), strData.getModuleDataMap());
    
    // 生成STR统计分析：测试统计、需求通过率、缺陷汇总
    generateSTRStatistics(outputPath, strData.getTestResults(), strData.getRequirements(), strData.getTestCases());
    
    return new GenerateResult(contentResult);
  }

  private StrStatistics analyzeStatistics(List<TestResult> testResults) {
    int total = testResults == null ? 0 : testResults.size();
    int passed = 0;
    int failed = 0;
    int others = 0;
    Map<String, Integer> verdictCounter = new HashMap<>();

    if (testResults != null) {
      for (TestResult tr : testResults) {
        String verdict = tr.getVerdict() == null ? "" : tr.getVerdict().trim();
        verdictCounter.put(verdict, verdictCounter.getOrDefault(verdict, 0) + 1);
        if ("通过".equals(verdict) || "PASS".equalsIgnoreCase(verdict)) {
          passed++;
        } else if ("失败".equals(verdict) || "FAIL".equalsIgnoreCase(verdict)) {
          failed++;
        } else {
          others++;
        }
      }
    }
    return new StrStatistics(total, passed, failed, others, verdictCounter);
  }

  /**
   * 生成STR统计分析数据
   * 包括：测试统计概览、需求通过率、缺陷汇总
   */
  private void generateSTRStatistics(String outputPath,
      List<TestResult> testResults,
      List<Requirement> requirements,
      List<TestCase> testCases) throws Exception {
    
    // 1. 测试统计概览
    int totalCases = testCases.size();
    int executedCount = 0;
    int passedCount = 0;
    int failedCount = 0;
    int blockedCount = 0;
    Map<String, Integer> verdictCounter = new HashMap<>();
    
    if (testResults != null) {
      executedCount = testResults.size();
      for (TestResult tr : testResults) {
        String verdict = tr.getVerdict() == null ? "" : tr.getVerdict().trim();
        verdictCounter.put(verdict, verdictCounter.getOrDefault(verdict, 0) + 1);
        if ("通过".equals(verdict) || "PASS".equalsIgnoreCase(verdict)) {
          passedCount++;
        } else if ("失败".equals(verdict) || "FAIL".equalsIgnoreCase(verdict)) {
          failedCount++;
        } else if ("阻塞".equals(verdict) || "BLOCKED".equalsIgnoreCase(verdict)) {
          blockedCount++;
        }
      }
    }
    
    double passRate = totalCases > 0 ? (passedCount * 100.0 / totalCases) : 0;
    
    System.out.println("===== STR测试统计概览 =====");
    System.out.println("总用例数: " + totalCases);
    System.out.println("已执行数: " + executedCount);
    System.out.println("通过数: " + passedCount);
    System.out.println("失败数: " + failedCount);
    System.out.println("阻塞数: " + blockedCount);
    System.out.println("通过率: " + String.format("%.2f%%", passRate));
    
    // 2. 需求通过率计算
    // 构建需求->用例映射和用例->需求映射
    Map<String, List<String>> reqToTcMap = new LinkedHashMap<>();
    Map<String, List<String>> tcToReqMap = new LinkedHashMap<>();
    Map<String, String> reqIdToTitleMap = new LinkedHashMap<>();
    
    for (Requirement req : requirements) {
      String reqId = req.getRequirementId();
      if (reqId != null && !reqId.isBlank()) {
        reqIdToTitleMap.put(reqId, req.getRequirementName());
        reqToTcMap.put(reqId, new ArrayList<>());
      }
    }
    
    for (TestCase tc : testCases) {
      String tcid = tc.getColumnValue("TCID");
      if (tcid == null || tcid.isBlank()) {
        tcid = tc.getColumnValue("标识");
      }
      if (tcid != null && !tcid.isBlank()) {
        tcToReqMap.put(tcid, new ArrayList<>());
        
        // 从用例的追踪关系获取覆盖的需求
        String trace = tc.getColumnValue("追踪关系");
        if (trace == null || trace.isBlank()) {
          trace = tc.getColumnValue("需求标识");
        }
        if (trace != null && !trace.isBlank()) {
          String[] reqIds = trace.split("[,，;；\\s]+");
          for (String reqId : reqIds) {
            reqId = reqId.trim();
            if (!reqId.isEmpty()) {
              tcToReqMap.get(tcid).add(reqId);
              if (reqToTcMap.containsKey(reqId)) {
                reqToTcMap.get(reqId).add(tcid);
              }
            }
          }
        }
      }
    }
    
    // 为每个需求计算通过率
    Map<String, Double> reqPassRateMap = new LinkedHashMap<>();
    for (String reqId : reqToTcMap.keySet()) {
      List<String> tcList = reqToTcMap.get(reqId);
      if (tcList.isEmpty()) {
        reqPassRateMap.put(reqId, 0.0); // 无用例覆盖
        continue;
      }
      
      int reqPassed = 0;
      for (String tcid : tcList) {
        // 查找该用例的最新测试结果
        TestResult latestResult = findLatestResult(testResults, tcid);
        if (latestResult != null) {
          String verdict = latestResult.getVerdict();
          if ("通过".equals(verdict) || "PASS".equalsIgnoreCase(verdict)) {
            reqPassed++;
          }
        }
      }
      double rate = tcList.size() > 0 ? (reqPassed * 100.0 / tcList.size()) : 0;
      reqPassRateMap.put(reqId, rate);
    }
    
    // 统计需求通过率
    int reqPassCount = 0;
    int reqFailCount = 0;
    int reqNoCoverCount = 0;
    for (Double rate : reqPassRateMap.values()) {
      if (rate == 0.0) {
        reqNoCoverCount++;
      } else if (rate >= 100.0) {
        reqPassCount++;
      } else {
        reqFailCount++;
      }
    }
    
    System.out.println("===== 需求通过率统计 =====");
    System.out.println("总需求数: " + reqIdToTitleMap.size());
    System.out.println("通过需求数: " + reqPassCount);
    System.out.println("未通过需求数: " + reqFailCount);
    System.out.println("无用例覆盖需求数: " + reqNoCoverCount);
    
    // 3. 缺陷汇总
    Map<String, List<String>> defectToTcMap = new LinkedHashMap<>();
    Map<String, List<String>> defectToReqMap = new LinkedHashMap<>();
    
    if (testResults != null) {
      for (TestResult tr : testResults) {
        String defectId = tr.getDefectId();
        if (defectId != null && !defectId.isBlank()) {
          String tcid = tr.getTestCaseId();
          
          if (!defectToTcMap.containsKey(defectId)) {
            defectToTcMap.put(defectId, new ArrayList<>());
            defectToReqMap.put(defectId, new ArrayList<>());
          }
          if (!defectToTcMap.get(defectId).contains(tcid)) {
            defectToTcMap.get(defectId).add(tcid);
            
            // 关联受影响的需求
            List<String> reqs = tcToReqMap.get(tcid);
            if (reqs != null) {
              for (String reqId : reqs) {
                if (!defectToReqMap.get(defectId).contains(reqId)) {
                  defectToReqMap.get(defectId).add(reqId);
                }
              }
            }
          }
        }
      }
    }
    
    System.out.println("===== 缺陷汇总 =====");
    System.out.println("缺陷总数: " + defectToTcMap.size());
    for (Map.Entry<String, List<String>> entry : defectToTcMap.entrySet()) {
      String defectId = entry.getKey();
      List<String> affectedTcs = entry.getValue();
      List<String> affectedReqs = defectToReqMap.get(defectId);
      System.out.println("  缺陷[" + defectId + "]: 影响用例" + affectedTcs.size() + "个, 影响需求" + 
          (affectedReqs != null ? affectedReqs.size() : 0) + "个");
    }
    
    // 输出测试结论建议
    System.out.println("===== 测试结论建议 =====");
    if (passRate >= 90.0 && reqPassCount == reqIdToTitleMap.size()) {
      System.out.println("建议: 测试通过，所有需求通过率100%");
    } else if (passRate >= 80.0) {
      System.out.println("建议: 测试基本通过，建议补充测试覆盖未通过的需求");
    } else {
      System.out.println("建议: 测试未通过，需要重新测试");
    }
  }

  /**
   * 查找用例的最新测试结果
   */
  private TestResult findLatestResult(List<TestResult> results, String tcid) {
    if (results == null || tcid == null) return null;
    
    TestResult latest = null;
    for (TestResult tr : results) {
      if (tcid.equals(tr.getTestCaseId())) {
        if (latest == null) {
          latest = tr;
        } else {
          // 简单比较：用后面的结果作为最新
          latest = tr;
        }
      }
    }
    return latest;
  }

  private record StrStatistics(int total, int passed, int failed, int others, Map<String, Integer> verdictCounter) {
  }
}