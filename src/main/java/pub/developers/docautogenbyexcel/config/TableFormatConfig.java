package pub.developers.docautogenbyexcel.config;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.*;

/**
 * 表格格式配置
 * 用于提高表格处理的通用性，支持配置化识别和填充规则
 */
public class TableFormatConfig {
    
    // 表格识别配置
    private final List<String> identifierKeywords;
    private final int minRows;
    private final int minCols;
    
    // 表格结构配置
    private final TableStructure structure;
    
    public TableFormatConfig(List<String> identifierKeywords, int minRows, int minCols, TableStructure structure) {
        this.identifierKeywords = identifierKeywords;
        this.minRows = minRows;
        this.minCols = minCols;
        this.structure = structure;
    }
    
    /**
     * 判断表格是否匹配此配置
     */
    public boolean matches(XWPFTable table) {
        if (table == null || table.getNumberOfRows() < minRows) {
            return false;
        }
        
        XWPFTableRow firstRow = table.getRow(0);
        if (firstRow == null || firstRow.getTableCells().size() < minCols) {
            return false;
        }
        
        String firstCellText = getCellText(firstRow.getCell(0));
        return identifierKeywords.stream()
            .anyMatch(firstCellText::contains);
    }
    
    /**
     * 填充表格数据
     */
    public void fillTable(XWPFTable table, Map<String, String> data) {
        if (table == null || data == null || data.isEmpty()) {
            return;
        }
        
        structure.fillTable(table, data);
    }
    
    private String getCellText(org.apache.poi.xwpf.usermodel.XWPFTableCell cell) {
        if (cell == null) return "";
        StringBuilder sb = new StringBuilder();
        for (org.apache.poi.xwpf.usermodel.XWPFParagraph p : cell.getParagraphs()) {
            sb.append(p.getText());
        }
        return sb.toString().trim();
    }
    
    /**
     * 表格结构定义
     */
    public static class TableStructure {
        private final StructureType type;
        private final FirstRowConfig firstRow;
        private final DataRowConfig dataRow;
        
        public TableStructure(StructureType type, FirstRowConfig firstRow, DataRowConfig dataRow) {
            this.type = type;
            this.firstRow = firstRow;
            this.dataRow = dataRow;
        }
        
        public void fillTable(XWPFTable table, Map<String, String> data) {
            switch (type) {
                case LABEL_VALUE:
                    fillLabelValueTable(table, data);
                    break;
                case HEADER_DATA:
                    fillHeaderDataTable(table, data);
                    break;
            }
        }
        
        private void fillLabelValueTable(XWPFTable table, Map<String, String> data) {
            int rowCount = table.getNumberOfRows();
            if (rowCount == 0) return;
            
            // 填充第一行（如果有特殊配置）
            if (firstRow != null && rowCount > 0) {
                XWPFTableRow firstRowObj = table.getRow(0);
                firstRow.fillRow(firstRowObj, data);
            }
            
            // 填充数据行
            int startRow = firstRow != null ? 1 : 0;
            for (int i = startRow; i < rowCount; i++) {
                XWPFTableRow row = table.getRow(i);
                if (row != null) {
                    dataRow.fillRow(row, data);
                }
            }
        }
        
        private void fillHeaderDataTable(XWPFTable table, Map<String, String> data) {
            // 表头数据格式的填充逻辑
            // 可以根据需要实现
        }
    }
    
    /**
     * 结构类型
     */
    public enum StructureType {
        LABEL_VALUE,  // 标签-值格式（如：测试项名称 | 值）
        HEADER_DATA   // 表头-数据格式（如：表头行 + 数据行）
    }
    
    /**
     * 第一行配置
     */
    public static class FirstRowConfig {
        private final List<Integer> labelCols;  // 标签列索引
        private final List<Integer> valueCols;  // 值列索引
        
        public FirstRowConfig(List<Integer> labelCols, List<Integer> valueCols) {
            this.labelCols = labelCols;
            this.valueCols = valueCols;
        }
        
        public void fillRow(XWPFTableRow row, Map<String, String> data) {
            if (row == null || labelCols.size() != valueCols.size()) return;
            
            for (int i = 0; i < labelCols.size(); i++) {
                int labelCol = labelCols.get(i);
                int valueCol = valueCols.get(i);
                
                if (labelCol < row.getTableCells().size() && 
                    valueCol < row.getTableCells().size()) {
                    String label = getCellText(row.getCell(labelCol));
                    String match = findMatchingColumn(label, data.keySet());
                    if (match != null) {
                        setCellText(row.getCell(valueCol), data.get(match));
                    }
                }
            }
        }
        
        private String getCellText(org.apache.poi.xwpf.usermodel.XWPFTableCell cell) {
            if (cell == null) return "";
            StringBuilder sb = new StringBuilder();
            for (org.apache.poi.xwpf.usermodel.XWPFParagraph p : cell.getParagraphs()) {
                sb.append(p.getText());
            }
            return sb.toString().trim();
        }
        
        private String findMatchingColumn(String label, Set<String> columns) {
            if (label == null || label.isEmpty()) return null;
            for (String col : columns) {
                if (col.equals(label) || label.contains(col) || col.contains(label)) {
                    return col;
                }
            }
            return null;
        }
        
        private void setCellText(org.apache.poi.xwpf.usermodel.XWPFTableCell cell, String text) {
            if (cell == null) return;
            while (!cell.getParagraphs().isEmpty()) {
                cell.removeParagraph(0);
            }
            org.apache.poi.xwpf.usermodel.XWPFParagraph para = cell.addParagraph();
            org.apache.poi.xwpf.usermodel.XWPFRun run = para.createRun();
            run.setText(text != null ? text : "");
        }
    }
    
    /**
     * 数据行配置
     */
    public static class DataRowConfig {
        private final int labelCol;  // 标签列索引
        private final int valueCol;  // 值列索引（考虑合并单元格）
        
        public DataRowConfig(int labelCol, int valueCol) {
            this.labelCol = labelCol;
            this.valueCol = valueCol;
        }
        
        public void fillRow(XWPFTableRow row, Map<String, String> data) {
            if (row == null || row.getTableCells().size() < Math.max(labelCol, valueCol) + 1) {
                return;
            }
            
            String label = getCellText(row.getCell(labelCol));
            String match = findMatchingColumn(label, data.keySet());
            if (match != null) {
                int actualValueCol = row.getTableCells().size() >= valueCol + 1 ? valueCol : labelCol + 1;
                setCellText(row.getCell(actualValueCol), data.get(match));
            }
        }
        
        private String getCellText(org.apache.poi.xwpf.usermodel.XWPFTableCell cell) {
            if (cell == null) return "";
            StringBuilder sb = new StringBuilder();
            for (org.apache.poi.xwpf.usermodel.XWPFParagraph p : cell.getParagraphs()) {
                sb.append(p.getText());
            }
            return sb.toString().trim();
        }
        
        private String findMatchingColumn(String label, Set<String> columns) {
            if (label == null || label.isEmpty()) return null;
            for (String col : columns) {
                if (col.equals(label) || label.contains(col) || col.contains(label)) {
                    return col;
                }
            }
            return null;
        }
        
        private void setCellText(org.apache.poi.xwpf.usermodel.XWPFTableCell cell, String text) {
            if (cell == null) return;
            while (!cell.getParagraphs().isEmpty()) {
                cell.removeParagraph(0);
            }
            org.apache.poi.xwpf.usermodel.XWPFParagraph para = cell.addParagraph();
            org.apache.poi.xwpf.usermodel.XWPFRun run = para.createRun();
            run.setText(text != null ? text : "");
        }
    }
    
    /**
     * 创建默认的测试用例表格配置
     */
    public static TableFormatConfig createDefaultTestCaseConfig() {
        List<String> keywords = Arrays.asList("测试项名称", "测试项");
        
        // 第一行配置：4列格式（标签0，值1，标签2，值3）
        FirstRowConfig firstRow = new FirstRowConfig(
            Arrays.asList(0, 2),
            Arrays.asList(1, 3)
        );
        
        // 数据行配置：标签列0，值列2（考虑合并单元格）
        DataRowConfig dataRow = new DataRowConfig(0, 2);
        
        TableStructure structure = new TableStructure(
            StructureType.LABEL_VALUE,
            firstRow,
            dataRow
        );
        
        return new TableFormatConfig(keywords, 5, 2, structure);
    }
}

