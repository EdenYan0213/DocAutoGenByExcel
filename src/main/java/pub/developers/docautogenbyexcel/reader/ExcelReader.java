package pub.developers.docautogenbyexcel.reader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import pub.developers.docautogenbyexcel.config.TableConfig;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.TestCase;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * Excel数据读取模块
 * 负责读取Excel文件，解析测试用例数据，并按模块分组
 * 自动搜索包含"模块编号"列的Sheet作为测试用例数据源
 */
public class ExcelReader {
    
    private final TableConfig config = TableConfig.getInstance();

    /**
     * 读取Excel文件并返回按模块分组的数据
     *
     * @param excelPath Excel文件路径
     * @return Map<模块编号, ModuleData>，同时返回列名列表
     * @throws Exception 读取异常
     */
    public Map<String, ModuleData> readExcel(String excelPath) throws Exception {
        File file = new File(excelPath);
        if (!file.exists() || !file.canRead()) {
            throw new Exception("Excel文件路径错误或文件损坏: " + excelPath);
        }

        // 检查文件格式
        if (!excelPath.toLowerCase().endsWith(".xlsx")) {
            throw new Exception("仅支持.xlsx格式的Excel文件，不支持.xls格式");
        }

        Map<String, ModuleData> moduleDataMap = new LinkedHashMap<>();
        String requiredColumn = config.getTestCaseRequiredColumn();
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            // 自动搜索包含必填列的Sheet
            Sheet sheet = null;
            String foundSheetName = null;
            List<String> columnNames = new ArrayList<>();
            Map<String, Integer> columnIndexMap = new HashMap<>();
            
            for (int sheetIdx = 0; sheetIdx < workbook.getNumberOfSheets(); sheetIdx++) {
                Sheet candidateSheet = workbook.getSheetAt(sheetIdx);
                if (candidateSheet == null || candidateSheet.getPhysicalNumberOfRows() == 0) {
                    continue;
                }
                
                Row headerRow = candidateSheet.getRow(0);
                if (headerRow == null) {
                    continue;
                }
                
                // 读取列名
                List<String> candidateColumns = new ArrayList<>();
                Map<String, Integer> candidateColumnIndex = new HashMap<>();
                
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    String cellValue = getCellValueAsString(cell);
                    if (cellValue != null && !cellValue.trim().isEmpty()) {
                        String columnName = cellValue.trim();
                        candidateColumns.add(columnName);
                        candidateColumnIndex.put(columnName, i);
                    }
                }
                
                // 检查是否包含必填列
                if (candidateColumnIndex.containsKey(requiredColumn)) {
                    sheet = candidateSheet;
                    foundSheetName = candidateSheet.getSheetName();
                    columnNames = candidateColumns;
                    columnIndexMap = candidateColumnIndex;
                    System.out.println("找到测试用例Sheet: " + foundSheetName + " (包含'" + requiredColumn + "'列)");
                    break;
                }
            }
            
            if (sheet == null) {
                System.out.println("未找到包含'" + requiredColumn + "'列的Sheet，跳过测试用例处理（将只处理基本信息、列表型表格等）");
                return moduleDataMap; // 返回空的Map，继续处理其他类型的表格
            }
            
            Row headerRow = sheet.getRow(0);

            // 读取数据行
            int totalRows = sheet.getPhysicalNumberOfRows();
            int dataCount = 0;
            
            for (int i = 1; i < totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }

                // 检查是否为空行
                if (isRowEmpty(row)) {
                    continue;
                }

                // 读取测试用例数据（包含所有列）
                TestCase testCase = readTestCase(row, columnIndexMap, columnNames);
                if (testCase == null) {
                    continue;
                }

                // 按模块分组
                String moduleNumber = testCase.getModuleNumber();
                if (moduleNumber == null || moduleNumber.trim().isEmpty()) {
                    continue;
                }

                moduleDataMap.computeIfAbsent(moduleNumber, ModuleData::new)
                        .addTestCase(testCase);
                dataCount++;
            }

            // 读取测试步骤Sheet并关联到TestCase
            readTestSteps(workbook, moduleDataMap);
            
            System.out.println("读取完成（共" + dataCount + "条数据，" + moduleDataMap.size() + "个模块，共" + columnNames.size() + "列）");
            return moduleDataMap;

        } catch (IOException e) {
            throw new Exception("读取Excel文件失败: " + e.getMessage(), e);
        }
    }
    
    /** 读取测试步骤Sheet并关联到TestCase */
    private void readTestSteps(Workbook workbook, Map<String, ModuleData> moduleDataMap) {
        // 查找测试步骤Sheet
        Sheet stepsSheet = null;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            String name = workbook.getSheetName(i);
            if (name.contains("测试步骤") || name.toLowerCase().contains("step")) {
                stepsSheet = workbook.getSheetAt(i);
                break;
            }
        }
        if (stepsSheet == null || stepsSheet.getPhysicalNumberOfRows() < 2) return;
        
        // 读取表头
        Row header = stepsSheet.getRow(0);
        Map<String, Integer> colMap = new HashMap<>();
        for (int i = 0; i < header.getLastCellNum(); i++) {
            String val = getCellValueAsString(header.getCell(i));
            if (val != null) colMap.put(val.trim(), i);
        }
        
        // 必须包含这些列
        Integer idCol = findColumn(colMap, "测试用例标识", "用例标识", "标识", "ID");
        Integer stepNoCol = findColumn(colMap, "步骤序号", "序号", "StepNo");
        Integer actionCol = findColumn(colMap, "测试步骤", "步骤", "操作", "输入及操作", "Action");
        Integer expectedCol = findColumn(colMap, "预期结果", "期望结果", "期望结果与评估标准", "Expected");
        Integer resultCol = findColumn(colMap, "实测结果", "测试结果", "Result");
        
        if (idCol == null || actionCol == null) {
            System.out.println("测试步骤Sheet缺少必要列（测试用例标识、测试步骤），跳过");
            return;
        }
        
        // 建立测试用例标识 -> TestCase 的映射
        Map<String, TestCase> caseMap = new HashMap<>();
        for (ModuleData module : moduleDataMap.values()) {
            for (TestCase tc : module.getTestCases()) {
                String caseId = tc.getColumnValue("测试用例标识");
                if (caseId == null || caseId.isEmpty()) caseId = tc.getColumnValue("标识");
                if (caseId != null && !caseId.isEmpty()) caseMap.put(caseId, tc);
            }
        }
        
        // 读取测试步骤数据
        int stepCount = 0;
        for (int i = 1; i < stepsSheet.getPhysicalNumberOfRows(); i++) {
            Row row = stepsSheet.getRow(i);
            if (row == null || isRowEmpty(row)) continue;
            
            String caseId = getCellValueAsString(row.getCell(idCol));
            if (caseId == null || caseId.isEmpty()) continue;
            
            TestCase tc = caseMap.get(caseId.trim());
            if (tc == null) continue;
            
            int stepNo = 0;
            if (stepNoCol != null) {
                String stepNoStr = getCellValueAsString(row.getCell(stepNoCol));
                try { stepNo = Integer.parseInt(stepNoStr); } catch (Exception e) { stepNo = tc.getTestSteps().size() + 1; }
            } else {
                stepNo = tc.getTestSteps().size() + 1;
            }
            
            String action = getCellValueAsString(row.getCell(actionCol));
            String expected = expectedCol != null ? getCellValueAsString(row.getCell(expectedCol)) : "";
            String result = resultCol != null ? getCellValueAsString(row.getCell(resultCol)) : "";
            
            tc.addTestStep(stepNo, action != null ? action : "", expected != null ? expected : "", result != null ? result : "");
            stepCount++;
        }
        
        if (stepCount > 0) {
            System.out.println("读取测试步骤: " + stepCount + " 条");
        }
    }
    
    /** 在多个可能的列名中查找 */
    private Integer findColumn(Map<String, Integer> colMap, String... names) {
        for (String name : names) {
            if (colMap.containsKey(name)) return colMap.get(name);
        }
        return null;
    }

    /**
     * 读取一行数据，转换为TestCase对象（包含所有列）
     */
    private TestCase readTestCase(Row row, Map<String, Integer> columnIndexMap, List<String> columnNames) {
        String requiredColumn = config.getTestCaseRequiredColumn();
        String moduleNumber = getCellValue(row, columnIndexMap.get(requiredColumn));
        
        // 验证必填字段
        if (moduleNumber == null || moduleNumber.trim().isEmpty()) {
            return null;
        }

        TestCase testCase = new TestCase(moduleNumber.trim());
        
        // 读取所有列的数据
        for (String columnName : columnNames) {
            // 跳过模块编号列（已经设置）
            if (requiredColumn.equals(columnName)) {
                continue;
            }
            
            String value = getCellValue(row, columnIndexMap.get(columnName));
            testCase.addColumnData(columnName, value != null ? value.trim() : "");
        }

        return testCase;
    }

    /**
     * 获取单元格值（字符串格式）
     */
    private String getCellValue(Row row, Integer columnIndex) {
        if (columnIndex == null || row == null) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        return getCellValueAsString(cell);
    }

    /**
     * 将单元格值转换为字符串
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // 避免科学计数法，保留原始格式
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == (long) numericValue) {
                        return String.valueOf((long) numericValue);
                    } else {
                        return String.valueOf(numericValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BLANK:
                return null;
            default:
                return null;
        }
    }

    /**
     * 判断行是否为空
     */
    private boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                String value = getCellValueAsString(cell);
                if (value != null && !value.trim().isEmpty()) {
                    return false;
                }
            }
        }
        return true;
    }
}

