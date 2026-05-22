package pub.developers.docautogenbyexcel.generator;

import pub.developers.docautogenbyexcel.hub.DataHub;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;

import java.util.*;

/**
 * 文档生成层：软件测试说明(STD)生成器
 * 继承AbstractDocumentGenerator，遵循extractData() → generateContent() → save()模板方法
 *
 * 生成《软件测试说明》文档及需求追溯矩阵
 */
public class STDGenerator extends AbstractDocumentGenerator {

    public STDGenerator(DataHub dataHub) {
        super(dataHub);
    }

    /**
     * STD生成器的数据容器
     * 封装从Excel提取的所有数据
     */
    public static class STDExtractedData implements ExtractedData {
        private final Map<String, ModuleData> moduleDataMap;
        private final Map<String, BasicInfoData> basicInfoMap;
        private final Map<String, ListTableData> listTableMap;
        private final List<Requirement> requirements;

        public STDExtractedData(
                Map<String, ModuleData> moduleDataMap,
                Map<String, BasicInfoData> basicInfoMap,
                Map<String, ListTableData> listTableMap,
                List<Requirement> requirements) {
            this.moduleDataMap = moduleDataMap;
            this.basicInfoMap = basicInfoMap;
            this.listTableMap = listTableMap;
            this.requirements = requirements;
        }

        public Map<String, ModuleData> getModuleDataMap() {
            return moduleDataMap;
        }

        public Map<String, BasicInfoData> getBasicInfoMap() {
            return basicInfoMap;
        }

        public Map<String, ListTableData> getListTableMap() {
            return listTableMap;
        }

        public List<Requirement> getRequirements() {
            return requirements;
        }
    }

    @Override
    protected STDExtractedData extractData(String excelPath) throws Exception {
        Map<String, ModuleData> moduleDataMap = dataHub.loadModuleData(excelPath);
        Map<String, BasicInfoData> basicInfoMap = dataHub.loadBasicInfo(excelPath);
        Map<String, ListTableData> listTableMap = dataHub.loadListTables(excelPath);
        List<Requirement> requirements = dataHub.loadRequirements(excelPath);
        return new STDExtractedData(moduleDataMap, basicInfoMap, listTableMap, requirements);
    }

    @Override
    protected int generateContent(String templatePath, String outputPath, ExtractedData extractedData)
            throws Exception {
        STDExtractedData stdData = (STDExtractedData) extractedData;
        return wordDocumentBuilder.buildModuleSections(templatePath, outputPath, stdData.getModuleDataMap());
    }

    @Override
    protected GenerateResult save(String outputPath, ExtractedData extractedData, int contentResult) throws Exception {
        STDExtractedData stdData = (STDExtractedData) extractedData;
        wordDocumentBuilder.fillAdditionalTables(outputPath, stdData.getBasicInfoMap(), stdData.getListTableMap(), stdData.getModuleDataMap());
        
        // 生成追溯矩阵和孤儿需求警告
        generateTraceMatrixAndOrphanWarnings(outputPath, stdData.getRequirements(), stdData.getModuleDataMap());
        
        return new GenerateResult(contentResult);
    }

    /**
     * 生成追溯矩阵和孤儿���求警告
     * 追溯矩阵：以需求为行、用例为列，交叉处标记"×"表示覆盖
     * 孤儿需求检测：没有任何用例覆盖的需求，在矩阵下方添加警告信息
     */
    private void generateTraceMatrixAndOrphanWarnings(String outputPath, 
            List<Requirement> requirements, Map<String, ModuleData> moduleDataMap) throws Exception {
        // 构建需求->用例映射关系
        Map<String, List<String>> reqToTcMap = new LinkedHashMap<>();
        Map<String, String> reqIdToTitleMap = new LinkedHashMap<>();
        
        // 初始化需求映射
        for (Requirement req : requirements) {
            String reqId = req.getRequirementId();
            if (reqId != null && !reqId.isBlank()) {
                reqIdToTitleMap.put(reqId, req.getRequirementName());
                reqToTcMap.put(reqId, new ArrayList<>());
            }
        }
        
        // 遍历所有用例，查找覆盖的需求
        for (ModuleData moduleData : moduleDataMap.values()) {
            for (TestCase tc : moduleData.getTestCases()) {
                String tcid = tc.getColumnValue("TCID");
                if (tcid == null || tcid.isBlank()) {
                    tcid = tc.getColumnValue("标识");
                }
                
                // 从用例的追踪关系字段获取覆盖的需求
                String trace = tc.getColumnValue("追踪关系");
                if (trace == null || trace.isBlank()) {
                    trace = tc.getColumnValue("需求标识");
                }
                
                if (trace != null && !trace.isBlank()) {
                    // 支持逗号分隔的多个需求ID
                    String[] reqIds = trace.split("[,，;；\\s]+");
                    for (String reqId : reqIds) {
                        reqId = reqId.trim();
                        if (!reqId.isEmpty() && reqToTcMap.containsKey(reqId)) {
                            reqToTcMap.get(reqId).add(tcid);
                        }
                    }
                }
            }
        }
        
        // 检测孤儿需求（没有被任何用例覆盖）
        List<String> orphanReqs = new ArrayList<>();
        for (Map.Entry<String, List<String>> entry : reqToTcMap.entrySet()) {
            if (entry.getValue().isEmpty()) {
                orphanReqs.add(entry.getKey());
            }
        }
        
        // 打印追溯矩阵信息
        System.out.println("===== STD追溯矩阵统计 =====");
        System.out.println("总需求数: " + reqIdToTitleMap.size());
        System.out.println("孤儿需求数: " + orphanReqs.size());
        for (String orphanReq : orphanReqs) {
            System.out.println("  孤儿需求: " + orphanReq + " - " + reqIdToTitleMap.get(orphanReq));
        }
        
        // 注意：实际的追溯矩阵和警告生成需要在WordProcessor中实现
        // 这里只做统计和日志输出，实际的Word文档内容生成依赖模板
        if (!orphanReqs.isEmpty()) {
            System.out.println("警告: 发现 " + orphanReqs.size() + " 个孤儿需求未被测试用例覆盖!");
        }
    }
}