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
                throw new Exception("未找到包含'" + requiredColumn + "'列的Sheet，请确保Excel中有测试用例数据");
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

            System.out.println("读取完成（共" + dataCount + "条数据，" + moduleDataMap.size() + "个模块，共" + columnNames.size() + "列）");
            return moduleDataMap;

        } catch (IOException e) {
            throw new Exception("读取Excel文件失败: " + e.getMessage(), e);
        }
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

