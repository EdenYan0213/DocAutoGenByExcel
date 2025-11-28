package pub.developers.docautogenbyexcel.reader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.TestCase;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * Excel数据读取模块
 * 负责读取Excel文件，解析测试用例数据，并按模块分组
 */
public class ExcelReader {
    
    // 必填列名
    private static final String COL_MODULE_NUMBER = "模块编号";
    private static final String COL_TEST_NAME = "testName";
    private static final String COL_ID = "id";
    private static final String COL_CONTENT = "content";
    private static final String COL_STRATEGY = "strategy";
    private static final String COL_CRITERIA = "criteria";
    private static final String COL_STOP_CONDITION = "stopCondition";
    private static final String COL_TRACE = "trace";
    
    private static final Set<String> REQUIRED_COLUMNS = new HashSet<>(Arrays.asList(
            COL_MODULE_NUMBER, COL_TEST_NAME, COL_ID, COL_CONTENT,
            COL_STRATEGY, COL_CRITERIA, COL_STOP_CONDITION, COL_TRACE
    ));

    /**
     * 读取Excel文件并返回按模块分组的数据
     *
     * @param excelPath Excel文件路径
     * @return Map<模块编号, ModuleData>
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
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
                throw new Exception("Excel文件为空或没有数据");
            }

            // 读取表头，确定列索引
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new Exception("Excel文件缺少表头行");
            }

            Map<String, Integer> columnIndexMap = new HashMap<>();
            for (Cell cell : headerRow) {
                String cellValue = getCellValueAsString(cell);
                if (cellValue != null && !cellValue.trim().isEmpty()) {
                    columnIndexMap.put(cellValue.trim(), cell.getColumnIndex());
                }
            }

            // 验证必填列
            List<String> missingColumns = new ArrayList<>();
            for (String requiredCol : REQUIRED_COLUMNS) {
                if (!columnIndexMap.containsKey(requiredCol)) {
                    missingColumns.add(requiredCol);
                }
            }
            if (!missingColumns.isEmpty()) {
                throw new Exception("Excel缺少必填列: " + String.join(", ", missingColumns));
            }

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

                // 读取测试用例数据
                TestCase testCase = readTestCase(row, columnIndexMap);
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

            System.out.println("读取完成（共" + dataCount + "条数据，" + moduleDataMap.size() + "个模块）");
            return moduleDataMap;

        } catch (IOException e) {
            throw new Exception("读取Excel文件失败: " + e.getMessage(), e);
        }
    }

    /**
     * 读取一行数据，转换为TestCase对象
     */
    private TestCase readTestCase(Row row, Map<String, Integer> columnIndexMap) {
        String moduleNumber = getCellValue(row, columnIndexMap.get(COL_MODULE_NUMBER));
        String testName = getCellValue(row, columnIndexMap.get(COL_TEST_NAME));
        String id = getCellValue(row, columnIndexMap.get(COL_ID));
        String content = getCellValue(row, columnIndexMap.get(COL_CONTENT));
        String strategy = getCellValue(row, columnIndexMap.get(COL_STRATEGY));
        String criteria = getCellValue(row, columnIndexMap.get(COL_CRITERIA));
        String stopCondition = getCellValue(row, columnIndexMap.get(COL_STOP_CONDITION));
        String trace = getCellValue(row, columnIndexMap.get(COL_TRACE));

        // 验证必填字段
        if (moduleNumber == null || moduleNumber.trim().isEmpty() ||
            testName == null || testName.trim().isEmpty() ||
            id == null || id.trim().isEmpty()) {
            return null;
        }

        return new TestCase(
                moduleNumber.trim(),
                testName.trim(),
                id.trim(),
                content != null ? content.trim() : "",
                strategy != null ? strategy.trim() : "",
                criteria != null ? criteria.trim() : "",
                stopCondition != null ? stopCondition.trim() : "",
                trace != null ? trace.trim() : ""
        );
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

