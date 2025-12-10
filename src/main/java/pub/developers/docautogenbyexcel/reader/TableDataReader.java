package pub.developers.docautogenbyexcel.reader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import pub.developers.docautogenbyexcel.config.TableConfig;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

/**
 * 表格数据读取器
 * 读取Excel中的基本信息和列表型表格数据
 * 根据Sheet内容自动识别数据类型，不依赖Sheet名称
 * 
 * 基本信息Sheet格式（根据配置文件识别）:
 * | 表格名称 | 字段名 | 字段值 |
 * | 表1.1 被测软件基本信息 | 软件等级 | D |
 * 
 * 列表型Sheet格式（根据配置文件识别）:
 * | 表格名称 | 序号 | 接口类型 | 方向 | 说明 |
 * | 表1.2 被测软件接口信息 | 1 | 串口 | 输入 | 接收控制命令 |
 */
public class TableDataReader {
    
    private final TableConfig config = TableConfig.getInstance();
    
    /**
     * 基本信息表格数据
     * Map<表格名称, Map<字段名, 字段值>>
     */
    public static class BasicInfoData {
        private String tableName;
        private Map<String, String> fields = new LinkedHashMap<>();
        
        public BasicInfoData(String tableName) {
            this.tableName = tableName;
        }
        
        public String getTableName() {
            return tableName;
        }
        
        public void addField(String fieldName, String fieldValue) {
            fields.put(fieldName, fieldValue);
        }
        
        public String getFieldValue(String fieldName) {
            return fields.get(fieldName);
        }
        
        public Map<String, String> getAllFields() {
            return fields;
        }
    }
    
    /**
     * 列表型表格数据（如接口信息）
     * 每行是一个Map<列名, 值>
     */
    public static class ListTableData {
        private String tableName;
        private List<String> columnNames = new ArrayList<>();
        private List<Map<String, String>> rows = new ArrayList<>();
        
        public ListTableData(String tableName) {
            this.tableName = tableName;
        }
        
        public String getTableName() {
            return tableName;
        }
        
        public void setColumnNames(List<String> names) {
            this.columnNames = names;
        }
        
        public List<String> getColumnNames() {
            return columnNames;
        }
        
        public void addRow(Map<String, String> row) {
            rows.add(row);
        }
        
        public List<Map<String, String>> getRows() {
            return rows;
        }
    }
    
    /**
     * 读取基本信息表格数据
     * 从Sheet "基本信息" 读取
     * 
     * @param excelPath Excel文件路径
     * @return Map<表格名称, BasicInfoData>
     */
    public Map<String, BasicInfoData> readBasicInfo(String excelPath) throws Exception {
        Map<String, BasicInfoData> result = new LinkedHashMap<>();
        
        File file = new File(excelPath);
        if (!file.exists()) {
            System.out.println("Excel文件不存在: " + excelPath);
            return result;
        }
        
        String expectedCol0 = config.getBasicInfoTableNameColumn();
        String expectedCol1 = config.getBasicInfoFieldNameColumn();
        String expectedCol2 = config.getBasicInfoFieldValueColumn();
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            // 自动搜索符合基本信息格式的 Sheet
            for (int sheetIdx = 0; sheetIdx < workbook.getNumberOfSheets(); sheetIdx++) {
                Sheet sheet = workbook.getSheetAt(sheetIdx);
                if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
                    continue;
                }
                
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) {
                    continue;
                }
                
                // 验证列名
                String col0 = getCellValue(headerRow.getCell(0));
                String col1 = getCellValue(headerRow.getCell(1));
                String col2 = getCellValue(headerRow.getCell(2));
                
                if (!expectedCol0.equals(col0) || !expectedCol1.equals(col1) || !expectedCol2.equals(col2)) {
                    continue;
                }
                
                System.out.println("发现基本信息Sheet: " + sheet.getSheetName());
                
                // 读取数据行
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;
                    
                    String tableName = getCellValue(row.getCell(0));
                    String fieldName = getCellValue(row.getCell(1));
                    String fieldValue = getCellValue(row.getCell(2));
                    
                    if (tableName == null || tableName.trim().isEmpty()) continue;
                    if (fieldName == null || fieldName.trim().isEmpty()) continue;
                    
                    tableName = tableName.trim();
                    fieldName = fieldName.trim();
                    fieldValue = fieldValue != null ? fieldValue.trim() : "";
                    
                    // 获取或创建 BasicInfoData
                    BasicInfoData data = result.get(tableName);
                    if (data == null) {
                        data = new BasicInfoData(tableName);
                        result.put(tableName, data);
                    }
                    data.addField(fieldName, fieldValue);
                }
            }
            
            if (!result.isEmpty()) {
                System.out.println("读取基本信息完成，共" + result.size() + "个表格");
            }
        }
        
        return result;
    }
    
    /**
     * 读取列表型表格数据
     * 从Sheet "接口信息" 或其他列表型Sheet读取
     * 
     * @param excelPath Excel文件路径
     * @param sheetName Sheet名称
     * @return Map<表格名称, ListTableData>
     */
    public Map<String, ListTableData> readListTableData(String excelPath, String sheetName) throws Exception {
        Map<String, ListTableData> result = new LinkedHashMap<>();
        
        File file = new File(excelPath);
        if (!file.exists()) {
            System.out.println("Excel文件不存在: " + excelPath);
            return result;
        }
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                System.out.println("未找到 '" + sheetName + "' Sheet，跳过读取");
                return result;
            }
            
            // 读取表头（第一行）
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                return result;
            }
            
            // 读取列名（从第二列开始，第一列是表格名称）
            List<String> columnNames = new ArrayList<>();
            for (int c = 1; c < headerRow.getLastCellNum(); c++) {
                String colName = getCellValue(headerRow.getCell(c));
                if (colName != null && !colName.trim().isEmpty()) {
                    columnNames.add(colName.trim());
                }
            }
            
            if (columnNames.isEmpty()) {
                System.out.println(sheetName + " Sheet没有数据列");
                return result;
            }
            
            // 读取数据行
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                
                String tableName = getCellValue(row.getCell(0));
                if (tableName == null || tableName.trim().isEmpty()) continue;
                tableName = tableName.trim();
                
                // 获取或创建 ListTableData
                ListTableData data = result.get(tableName);
                if (data == null) {
                    data = new ListTableData(tableName);
                    data.setColumnNames(columnNames);
                    result.put(tableName, data);
                }
                
                // 读取行数据
                Map<String, String> rowData = new LinkedHashMap<>();
                for (int c = 0; c < columnNames.size(); c++) {
                    String value = getCellValue(row.getCell(c + 1));
                    rowData.put(columnNames.get(c), value != null ? value.trim() : "");
                }
                data.addRow(rowData);
            }
            
            System.out.println("读取" + sheetName + "完成，共" + result.size() + "个表格");
        }
        
        return result;
    }
    
    /**
     * 读取所有列表型表格数据
     * 自动识别列表型Sheet（第一列是配置的表格名称列）
     * 跳过测试用例Sheet（包含模块编号列）和基本信息Sheet（包含字段名、字段值列）
     * 
     * @param excelPath Excel文件路径
     * @return Map<表格名称, ListTableData>
     */
    public Map<String, ListTableData> readAllListTableData(String excelPath) throws Exception {
        Map<String, ListTableData> result = new LinkedHashMap<>();
        
        File file = new File(excelPath);
        if (!file.exists()) {
            System.out.println("Excel文件不存在: " + excelPath);
            return result;
        }
        
        String testCaseRequiredColumn = config.getTestCaseRequiredColumn();
        String basicInfoFieldNameColumn = config.getBasicInfoFieldNameColumn();
        String basicInfoFieldValueColumn = config.getBasicInfoFieldValueColumn();
        String listDataTableNameColumn = config.getListDataTableNameColumn();
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            // 遍历所有Sheet
            for (int sheetIdx = 0; sheetIdx < workbook.getNumberOfSheets(); sheetIdx++) {
                Sheet sheet = workbook.getSheetAt(sheetIdx);
                String sheetName = sheet.getSheetName();
                
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) {
                    continue;
                }
                
                // 读取所有列名
                List<String> headerColumns = new ArrayList<>();
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    String colName = getCellValue(headerRow.getCell(c));
                    headerColumns.add(colName != null ? colName.trim() : "");
                }
                
                // 跳过测试用例Sheet（包含模块编号列）
                if (headerColumns.contains(testCaseRequiredColumn)) {
                    continue;
                }
                
                // 跳过基本信息Sheet（第二列是字段名，第三列是字段值）
                if (headerColumns.size() >= 3 && 
                    headerColumns.get(1).equals(basicInfoFieldNameColumn) &&
                    headerColumns.get(2).equals(basicInfoFieldValueColumn)) {
                    continue;
                }
                
                // 检查是否是列表型Sheet（第一列是表格名称列）
                String firstColName = headerColumns.isEmpty() ? "" : headerColumns.get(0);
                if (!firstColName.equals(listDataTableNameColumn)) {
                    continue;
                }
                
                System.out.println("发现列表型Sheet: " + sheetName);
                
                // 读取列名（从第二列开始）
                List<String> columnNames = new ArrayList<>();
                for (int c = 1; c < headerRow.getLastCellNum(); c++) {
                    String colName = getCellValue(headerRow.getCell(c));
                    if (colName != null && !colName.trim().isEmpty()) {
                        columnNames.add(colName.trim());
                    }
                }
                
                if (columnNames.isEmpty()) {
                    continue;
                }
                
                // 读取数据行
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;
                    
                    String tableName = getCellValue(row.getCell(0));
                    if (tableName == null || tableName.trim().isEmpty()) continue;
                    tableName = tableName.trim();
                    
                    // 获取或创建 ListTableData
                    ListTableData data = result.get(tableName);
                    if (data == null) {
                        data = new ListTableData(tableName);
                        data.setColumnNames(columnNames);
                        result.put(tableName, data);
                    }
                    
                    // 读取行数据
                    Map<String, String> rowData = new LinkedHashMap<>();
                    for (int c = 0; c < columnNames.size(); c++) {
                        String value = getCellValue(row.getCell(c + 1));
                        rowData.put(columnNames.get(c), value != null ? value.trim() : "");
                    }
                    data.addRow(rowData);
                }
            }
            
            System.out.println("共读取 " + result.size() + " 个列表型表格");
        }
        
        return result;
    }
    
    /**
     * 获取单元格值（转为字符串）
     */
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                double numValue = cell.getNumericCellValue();
                if (numValue == Math.floor(numValue)) {
                    return String.valueOf((long) numValue);
                }
                return String.valueOf(numValue);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e2) {
                        return "";
                    }
                }
            default:
                return "";
        }
    }
}

