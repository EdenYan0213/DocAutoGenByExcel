package pub.developers.docautogenbyexcel.config;

import java.util.*;
import java.util.stream.Collectors;

/**
 * 表格格式配置加载器
 * 从配置文件加载表格格式配置，提高通用性
 */
public class TableFormatConfigLoader {
    
    private final Properties properties;
    
    public TableFormatConfigLoader(Properties properties) {
        this.properties = properties;
    }
    
    /**
     * 加载测试用例表格配置
     */
    public TableFormatConfig loadTestCaseTableConfig() {
        // 读取识别关键字
        String keywordsStr = properties.getProperty("testcase.table.identifier.keywords", "测试项名称,测试项");
        List<String> keywords = Arrays.stream(keywordsStr.split(","))
            .map(String::trim)
            .filter(s -> !s.isEmpty())
            .collect(Collectors.toList());
        
        // 读取最小行数和列数
        int minRows = Integer.parseInt(properties.getProperty("testcase.table.min.rows", "5"));
        int minCols = Integer.parseInt(properties.getProperty("testcase.table.min.cols", "2"));
        
        // 读取第一行配置
        String firstRowLabelColsStr = properties.getProperty("testcase.table.firstrow.label.cols", "0,2");
        String firstRowValueColsStr = properties.getProperty("testcase.table.firstrow.value.cols", "1,3");
        List<Integer> firstRowLabelCols = parseIntegerList(firstRowLabelColsStr);
        List<Integer> firstRowValueCols = parseIntegerList(firstRowValueColsStr);
        TableFormatConfig.FirstRowConfig firstRow = new TableFormatConfig.FirstRowConfig(
            firstRowLabelCols, firstRowValueCols);
        
        // 读取数据行配置
        int dataRowLabelCol = Integer.parseInt(properties.getProperty("testcase.table.datarow.label.col", "0"));
        int dataRowValueCol = Integer.parseInt(properties.getProperty("testcase.table.datarow.value.col", "2"));
        TableFormatConfig.DataRowConfig dataRow = new TableFormatConfig.DataRowConfig(
            dataRowLabelCol, dataRowValueCol);
        
        // 创建表格结构
        String structureTypeStr = properties.getProperty("testcase.table.structure.type", "LABEL_VALUE");
        TableFormatConfig.StructureType structureType = TableFormatConfig.StructureType.valueOf(structureTypeStr);
        TableFormatConfig.TableStructure structure = new TableFormatConfig.TableStructure(
            structureType, firstRow, dataRow);
        
        return new TableFormatConfig(keywords, minRows, minCols, structure);
    }
    
    /**
     * 获取表格评分关键字
     */
    public Map<String, Integer> getScoringKeywords() {
        Map<String, Integer> scoringMap = new HashMap<>();
        String scoringStr = properties.getProperty("testcase.table.scoring.keywords", 
            "测试内容=5,测试策略=3,判定准则=3,追踪关系=2");
        
        for (String pair : scoringStr.split(",")) {
            String[] parts = pair.trim().split("=");
            if (parts.length == 2) {
                try {
                    scoringMap.put(parts[0].trim(), Integer.parseInt(parts[1].trim()));
                } catch (NumberFormatException e) {
                    // 忽略无效的配置
                }
            }
        }
        
        return scoringMap;
    }
    
    /**
     * 获取优先表格行数
     */
    public int getPreferredRows() {
        return Integer.parseInt(properties.getProperty("testcase.table.preferred.rows", "6"));
    }
    
    /**
     * 解析整数列表
     */
    private List<Integer> parseIntegerList(String str) {
        return Arrays.stream(str.split(","))
            .map(String::trim)
            .filter(s -> !s.isEmpty())
            .map(Integer::parseInt)
            .collect(Collectors.toList());
    }
}

