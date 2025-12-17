package pub.developers.docautogenbyexcel.processor;

import org.apache.poi.xwpf.usermodel.*;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 表格填充处理器
 * 根据Excel数据填充Word文档中的各种表格
 * 
 * 支持两种表格类型：
 * 1. 基本信息表格（键值对形式）：如"表1.1 被测软件基本信息"
 * 2. 列表型表格：如"表1.2 被测软件接口信息"
 */
public class TableFillProcessor {
    
    // 表格标题匹配模式
    private static final Pattern TABLE_CAPTION_PATTERN = Pattern.compile("表\\s*([\\d.]+)\\s*(.+)");
    
    /**
     * 填充基本信息表格
     * 
     * @param document Word文档
     * @param basicInfoMap 基本信息数据
     * @return 填充的表格数量
     */
    public int fillBasicInfoTables(XWPFDocument document, Map<String, BasicInfoData> basicInfoMap) {
        if (basicInfoMap == null || basicInfoMap.isEmpty()) {
            return 0;
        }
        
        int filledCount = 0;
        List<XWPFTable> tables = document.getTables();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        // 遍历所有表格
        for (int tableIdx = 0; tableIdx < tables.size(); tableIdx++) {
            XWPFTable table = tables.get(tableIdx);
            
            // 查找表格前面的Caption
            String tableCaption = findTableCaption(document, table);
            if (tableCaption == null || tableCaption.isEmpty()) {
                continue;
            }
            
            System.out.println("检查表格Caption: " + tableCaption);
            
            // 检查是否匹配任何基本信息表格
            for (Map.Entry<String, BasicInfoData> entry : basicInfoMap.entrySet()) {
                String targetName = entry.getKey();
                BasicInfoData data = entry.getValue();
                
                // 匹配表格名称（可以是完全匹配或包含匹配）
                if (captionMatches(tableCaption, targetName)) {
                    System.out.println("找到匹配的基本信息表格: " + targetName);
                    fillBasicInfoTable(table, data);
                    filledCount++;
                    break;
                }
            }
        }
        
        System.out.println("基本信息表格填充完成，共填充" + filledCount + "个表格");
        return filledCount;
    }
    
    /**
     * 填充列表型表格
     * 
     * @param document Word文档
     * @param listTableMap 列表型表格数据
     * @return 填充的表格数量
     */
    public int fillListTables(XWPFDocument document, Map<String, ListTableData> listTableMap) {
        if (listTableMap == null || listTableMap.isEmpty()) {
            return 0;
        }
        
        int filledCount = 0;
        List<XWPFTable> tables = document.getTables();
        
        // 遍历所有表格
        for (XWPFTable table : tables) {
            // 查找表格前面的Caption
            String tableCaption = findTableCaption(document, table);
            if (tableCaption == null || tableCaption.isEmpty()) {
                continue;
            }
            
            // 检查是否匹配任何列表型表格
            for (Map.Entry<String, ListTableData> entry : listTableMap.entrySet()) {
                String targetName = entry.getKey();
                ListTableData data = entry.getValue();
                
                if (captionMatches(tableCaption, targetName)) {
                    System.out.println("找到匹配的列表型表格: " + targetName);
                    fillListTable(table, data);
                    filledCount++;
                    break;
                }
            }
        }
        
        System.out.println("列表型表格填充完成，共填充" + filledCount + "个表格");
        return filledCount;
    }
    
    /**
     * 查找表格的Caption（表格标题）
     * Caption通常在表格前面的段落中
     */
    private String findTableCaption(XWPFDocument document, XWPFTable table) {
        List<IBodyElement> elements = document.getBodyElements();
        
        for (int i = 0; i < elements.size(); i++) {
            if (elements.get(i) instanceof XWPFTable) {
                XWPFTable currentTable = (XWPFTable) elements.get(i);
                if (currentTable.getCTTbl() == table.getCTTbl()) {
                    // 找到表格，向前查找Caption
                    for (int j = i - 1; j >= 0 && j >= i - 3; j--) {
                        if (elements.get(j) instanceof XWPFParagraph) {
                            XWPFParagraph para = (XWPFParagraph) elements.get(j);
                            String text = para.getText().trim();
                            String style = para.getStyle();
                            
                            // 检查是否是Caption样式或以"表"开头
                            if ((style != null && (style.equals("11") || 
                                style.toLowerCase().contains("caption"))) ||
                                text.startsWith("表")) {
                                return text;
                            }
                        }
                    }
                    break;
                }
            }
        }
        
        return null;
    }
    
    /**
     * 检查Caption是否匹配目标表格名称
     */
    private boolean captionMatches(String caption, String targetName) {
        if (caption == null || targetName == null) {
            return false;
        }
        
        // 移除空格后比较
        String normalizedCaption = caption.replaceAll("\\s+", "");
        String normalizedTarget = targetName.replaceAll("\\s+", "");
        
        // 完全匹配
        if (normalizedCaption.equals(normalizedTarget)) {
            return true;
        }
        
        // 包含匹配（Caption包含目标名称，或目标名称包含Caption的关键部分）
        if (normalizedCaption.contains(normalizedTarget) || 
            normalizedTarget.contains(normalizedCaption)) {
            return true;
        }
        
        // 提取表格编号和名称进行匹配
        Matcher captionMatcher = TABLE_CAPTION_PATTERN.matcher(caption);
        Matcher targetMatcher = TABLE_CAPTION_PATTERN.matcher(targetName);
        
        if (captionMatcher.find() && targetMatcher.find()) {
            String captionNum = captionMatcher.group(1);
            String targetNum = targetMatcher.group(1);
            return captionNum.equals(targetNum);
        }
        
        return false;
    }
    
    /**
     * 填充基本信息表格
     * 根据字段名找到对应的单元格并填充值
     */
    private void fillBasicInfoTable(XWPFTable table, BasicInfoData data) {
        Map<String, String> fields = data.getAllFields();
        
        for (XWPFTableRow row : table.getRows()) {
            List<XWPFTableCell> cells = row.getTableCells();
            
            // 遍历每个单元格，查找匹配的字段名
            for (int i = 0; i < cells.size(); i++) {
                XWPFTableCell cell = cells.get(i);
                String cellText = cell.getText().trim();
                
                // 检查该单元格是否是字段名
                String fieldValue = findFieldValue(fields, cellText);
                if (fieldValue != null) {
                    // 找到对应的数据单元格（通常是右边的单元格）
                    int dataColIndex = findDataCellIndex(cells, i, cellText);
                    if (dataColIndex >= 0 && dataColIndex < cells.size()) {
                        XWPFTableCell dataCell = cells.get(dataColIndex);
                        setCellText(dataCell, fieldValue);
                        System.out.println("  填充字段: " + cellText + " = " + fieldValue);
                    }
                }
            }
        }
    }
    
    /**
     * 根据字段名查找字段值（支持模糊匹配）
     */
    private String findFieldValue(Map<String, String> fields, String cellText) {
        if (cellText == null || cellText.isEmpty()) {
            return null;
        }
        
        // 精确匹配
        if (fields.containsKey(cellText)) {
            return fields.get(cellText);
        }
        
        // 模糊匹配（去除空格和标点）
        String normalizedCellText = cellText.replaceAll("[\\s\\p{Punct}]", "");
        for (Map.Entry<String, String> entry : fields.entrySet()) {
            String normalizedKey = entry.getKey().replaceAll("[\\s\\p{Punct}]", "");
            if (normalizedCellText.equals(normalizedKey) || 
                normalizedCellText.contains(normalizedKey) ||
                normalizedKey.contains(normalizedCellText)) {
                return entry.getValue();
            }
        }
        
        return null;
    }
    
    /**
     * 查找数据单元格的索引
     * 通常是标签单元格右边的单元格
     */
    private int findDataCellIndex(List<XWPFTableCell> cells, int labelIndex, String labelText) {
        // 检查右边是否有单元格
        if (labelIndex + 1 < cells.size()) {
            return labelIndex + 1;
        }
        return -1;
    }
    
    /**
     * 填充列表型表格
     * 根据列名找到对应的列并填充数据
     */
    private void fillListTable(XWPFTable table, ListTableData data) {
        List<Map<String, String>> rows = data.getRows();
        if (rows.isEmpty()) {
            return;
        }
        
        // 获取表格的表头行（第一行）
        XWPFTableRow headerRow = table.getRow(0);
        if (headerRow == null) {
            return;
        }
        
        // 建立列名到列索引的映射
        Map<String, Integer> columnIndexMap = new LinkedHashMap<>();
        List<XWPFTableCell> headerCells = headerRow.getTableCells();
        for (int i = 0; i < headerCells.size(); i++) {
            String colName = headerCells.get(i).getText().trim();
            if (!colName.isEmpty()) {
                columnIndexMap.put(colName, i);
            }
        }
        
        System.out.println("  表格列: " + columnIndexMap.keySet());
        
        // 填充数据行
        int dataRowIndex = 1; // 从第二行开始
        for (Map<String, String> rowData : rows) {
            XWPFTableRow tableRow = table.getRow(dataRowIndex);
            
            // 如果行不存在，创建新行
            if (tableRow == null) {
                tableRow = table.createRow();
            }
            
            // 填充各列数据
            for (Map.Entry<String, String> entry : rowData.entrySet()) {
                String colName = entry.getKey();
                String value = entry.getValue();
                
                // 查找列索引（支持模糊匹配）
                Integer colIndex = findColumnIndex(columnIndexMap, colName);
                if (colIndex != null && colIndex < tableRow.getTableCells().size()) {
                    XWPFTableCell cell = tableRow.getCell(colIndex);
                    setCellText(cell, value);
                }
            }
            
            dataRowIndex++;
        }
        
        System.out.println("  填充了" + rows.size() + "行数据");
    }
    
    /**
     * 查找列索引（支持模糊匹配）
     */
    private Integer findColumnIndex(Map<String, Integer> columnIndexMap, String colName) {
        // 精确匹配
        if (columnIndexMap.containsKey(colName)) {
            return columnIndexMap.get(colName);
        }
        
        // 模糊匹配
        String normalizedColName = colName.replaceAll("[\\s\\p{Punct}]", "");
        for (Map.Entry<String, Integer> entry : columnIndexMap.entrySet()) {
            String normalizedKey = entry.getKey().replaceAll("[\\s\\p{Punct}]", "");
            if (normalizedColName.equals(normalizedKey) ||
                normalizedColName.contains(normalizedKey) ||
                normalizedKey.contains(normalizedColName)) {
                return entry.getValue();
            }
        }
        
        return null;
    }
    
    /**
     * 设置单元格文本（保持格式）
     */
    private void setCellText(XWPFTableCell cell, String text) {
        if (cell == null || text == null) {
            return;
        }
        
        // 清除现有内容
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        for (int i = paragraphs.size() - 1; i > 0; i--) {
            cell.removeParagraph(i);
        }
        
        if (paragraphs.isEmpty()) {
            cell.addParagraph();
            paragraphs = cell.getParagraphs();
        }
        
        XWPFParagraph para = paragraphs.get(0);
        
        // 清除现有runs
        while (!para.getRuns().isEmpty()) {
            para.removeRun(0);
        }
        
        // 添加新文本
        XWPFRun run = para.createRun();
        run.setText(text);
    }
}

