package pub.developers.docautogenbyexcel.reader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import pub.developers.docautogenbyexcel.model.Requirement;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * 需求Excel读取器
 * 负责从Excel文件读取需求数据
 */
public class RequirementExcelReader {
    
    // 必填列名
    private static final String COL_REQUIREMENT_ID = "需求ID";
    private static final String COL_REQUIREMENT_NAME = "需求名称";
    
    // 可选列名
    private static final String COL_REQUIREMENT_NUMBER = "需求编号";
    private static final String COL_DESCRIPTION = "需求描述";
    private static final String COL_TYPE = "需求类型";
    private static final String COL_PRIORITY = "优先级";
    private static final String COL_PARENT_ID = "父需求ID";
    private static final String COL_STATUS = "状态";
    
    /**
     * 读取需求Excel文件
     * 
     * @param excelPath Excel文件路径
     * @return 需求列表
     * @throws Exception 读取异常
     */
    public List<Requirement> readRequirements(String excelPath) throws Exception {
        File file = new File(excelPath);
        if (!file.exists() || !file.canRead()) {
            throw new Exception("Excel文件路径错误或文件损坏: " + excelPath);
        }
        
        if (!excelPath.toLowerCase().endsWith(".xlsx")) {
            throw new Exception("仅支持.xlsx格式的Excel文件");
        }
        
        List<Requirement> requirements = new ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
                throw new Exception("Excel文件为空或没有数据");
            }
            
            // 读取表头
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new Exception("Excel文件缺少表头行");
            }
            
            // 建立列名到索引的映射
            Map<String, Integer> columnIndexMap = new HashMap<>();
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                String cellValue = getCellValueAsString(cell);
                if (cellValue != null && !cellValue.trim().isEmpty()) {
                    columnIndexMap.put(cellValue.trim(), i);
                }
            }
            
            // 验证必填列
            if (!columnIndexMap.containsKey(COL_REQUIREMENT_ID)) {
                throw new Exception("Excel缺少必填列: " + COL_REQUIREMENT_ID);
            }
            if (!columnIndexMap.containsKey(COL_REQUIREMENT_NAME)) {
                throw new Exception("Excel缺少必填列: " + COL_REQUIREMENT_NAME);
            }
            
            // 读取数据行
            int totalRows = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < totalRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null || isRowEmpty(row)) {
                    continue;
                }
                
                Requirement requirement = readRequirement(row, columnIndexMap);
                if (requirement != null) {
                    requirements.add(requirement);
                }
            }
            
            System.out.println("读取完成（共" + requirements.size() + "条需求）");
            return requirements;
            
        } catch (IOException e) {
            throw new Exception("读取Excel文件失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 读取一行需求数据
     */
    private Requirement readRequirement(Row row, Map<String, Integer> columnIndexMap) {
        String requirementId = getCellValue(row, columnIndexMap.get(COL_REQUIREMENT_ID));
        String requirementName = getCellValue(row, columnIndexMap.get(COL_REQUIREMENT_NAME));
        
        if (requirementId == null || requirementId.trim().isEmpty() ||
            requirementName == null || requirementName.trim().isEmpty()) {
            return null;
        }
        
        Requirement requirement = new Requirement();
        requirement.setRequirementId(requirementId.trim());
        requirement.setRequirementName(requirementName.trim());
        
        // 读取可选字段
        String requirementNumber = getCellValue(row, columnIndexMap.get(COL_REQUIREMENT_NUMBER));
        if (requirementNumber != null && !requirementNumber.trim().isEmpty()) {
            requirement.setRequirementNumber(requirementNumber.trim());
        } else {
            requirement.setRequirementNumber(requirementId.trim());
        }
        
        String description = getCellValue(row, columnIndexMap.get(COL_DESCRIPTION));
        if (description != null) {
            requirement.setDescription(description.trim());
        }
        
        String type = getCellValue(row, columnIndexMap.get(COL_TYPE));
        if (type != null && !type.trim().isEmpty()) {
            requirement.setType(parseRequirementType(type.trim()));
        }
        
        String priority = getCellValue(row, columnIndexMap.get(COL_PRIORITY));
        if (priority != null && !priority.trim().isEmpty()) {
            requirement.setPriority(parsePriority(priority.trim()));
        }
        
        String parentId = getCellValue(row, columnIndexMap.get(COL_PARENT_ID));
        if (parentId != null && !parentId.trim().isEmpty()) {
            requirement.setParentRequirementId(parentId.trim());
        }
        
        String status = getCellValue(row, columnIndexMap.get(COL_STATUS));
        if (status != null && !status.trim().isEmpty()) {
            requirement.setStatus(parseStatus(status.trim()));
        }
        
        // 读取其他扩展属性
        for (Map.Entry<String, Integer> entry : columnIndexMap.entrySet()) {
            String columnName = entry.getKey();
            if (!isStandardColumn(columnName)) {
                String value = getCellValue(row, entry.getValue());
                requirement.addAttribute(columnName, value != null ? value.trim() : "");
            }
        }
        
        return requirement;
    }
    
    /**
     * 解析需求类型
     */
    private Requirement.RequirementType parseRequirementType(String type) {
        for (Requirement.RequirementType rt : Requirement.RequirementType.values()) {
            if (rt.name().equalsIgnoreCase(type) || 
                rt.getDescription().equals(type)) {
                return rt;
            }
        }
        return Requirement.RequirementType.FUNCTIONAL;
    }
    
    /**
     * 解析优先级
     */
    private Requirement.Priority parsePriority(String priority) {
        for (Requirement.Priority p : Requirement.Priority.values()) {
            if (p.name().equalsIgnoreCase(priority) || 
                p.getDescription().equals(priority)) {
                return p;
            }
        }
        return Requirement.Priority.MEDIUM;
    }
    
    /**
     * 解析状态
     */
    private Requirement.RequirementStatus parseStatus(String status) {
        for (Requirement.RequirementStatus s : Requirement.RequirementStatus.values()) {
            if (s.name().equalsIgnoreCase(status) || 
                s.getDescription().equals(status)) {
                return s;
            }
        }
        return Requirement.RequirementStatus.DRAFT;
    }
    
    /**
     * 判断是否为标准列
     */
    private boolean isStandardColumn(String columnName) {
        return COL_REQUIREMENT_ID.equals(columnName) ||
               COL_REQUIREMENT_NUMBER.equals(columnName) ||
               COL_REQUIREMENT_NAME.equals(columnName) ||
               COL_DESCRIPTION.equals(columnName) ||
               COL_TYPE.equals(columnName) ||
               COL_PRIORITY.equals(columnName) ||
               COL_PARENT_ID.equals(columnName) ||
               COL_STATUS.equals(columnName);
    }
    
    /**
     * 获取单元格值
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

