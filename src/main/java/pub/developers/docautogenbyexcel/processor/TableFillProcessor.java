package pub.developers.docautogenbyexcel.processor;

import org.apache.poi.xwpf.usermodel.*;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

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
     * @param document     Word文档
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
     * @param document     Word文档
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
     * 填充测试项追踪表（如：表9.1 测试依据到测试项的追踪）
     * 数据来源：测试用例Sheet
     *
     * 映射规则：
     * 序号 -> 自增
     * 合同指标编号 -> Excel列"合同指标编号"
     * 合同指标内容 -> Excel列"合同指标内容"
     * 测试项 -> Excel列"测试项名称"
     */
    public int fillTestTraceabilityTables(XWPFDocument document, Map<String, ModuleData> moduleDataMap) {
        if (moduleDataMap == null || moduleDataMap.isEmpty()) {
            return 0;
        }

        List<Map<String, String>> traceRows = buildTraceabilityRows(moduleDataMap);
        if (traceRows.isEmpty()) {
            return 0;
        }

        int filledCount = 0;
        for (XWPFTable table : document.getTables()) {
            String tableCaption = findTableCaption(document, table);
            if (!isTraceabilityCaption(tableCaption)) {
                continue;
            }

            if (fillTraceabilityTable(table, traceRows)) {
                filledCount++;
            }
        }

        if (filledCount > 0) {
            System.out.println("测试项追踪表填充完成，共填充" + filledCount + "个表格，" + traceRows.size() + "行数据");
        }
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

    /** 从测试用例数据构建追踪表行 */
    private List<Map<String, String>> buildTraceabilityRows(Map<String, ModuleData> moduleDataMap) {
        List<Map<String, String>> rows = new ArrayList<>();

        for (ModuleData moduleData : moduleDataMap.values()) {
            if (moduleData == null || moduleData.getTestCases() == null) {
                continue;
            }

            for (TestCase testCase : moduleData.getTestCases()) {
                if (testCase == null) {
                    continue;
                }

                String contractNo = firstNonBlank(testCase,
                        "合同指标编号", "合同指标编号（合同）", "合同指标编号(合同)");
                String contractContent = firstNonBlank(testCase,
                        "合同指标内容", "合同指标内容（合同）", "合同指标内容(合同)");
                String testItem = firstNonBlank(testCase, "测试项名称", "测试项");
                // 读取追踪关系列
                String traceRelation = firstNonBlank(testCase, "追踪关系", "trace", "追踪");

                if (isBlank(contractNo) && isBlank(contractContent) && isBlank(testItem) && isBlank(traceRelation)) {
                    continue;
                }

                Map<String, String> row = new LinkedHashMap<>();
                row.put("合同指标编号", contractNo != null ? contractNo : "");
                row.put("合同指标内容", contractContent != null ? contractContent : "");
                row.put("测试项", testItem != null ? testItem : "");
                // 添加追踪关系列
                row.put("追踪关系", traceRelation != null ? traceRelation : "");
                rows.add(row);
            }
        }

        return rows;
    }

    /** 判断Caption是否是测试项追踪表 */
    private boolean isTraceabilityCaption(String caption) {
        if (caption == null || caption.trim().isEmpty()) {
            return false;
        }

        String normalized = caption.replaceAll("\\s+", "");
        return normalized.contains("测试依据到测试项的追踪") || normalized.contains("测试项的追踪");
    }

    /** 填充测试项追踪表 */
    private boolean fillTraceabilityTable(XWPFTable table, List<Map<String, String>> traceRows) {
        if (table.getNumberOfRows() <= 0) {
            return false;
        }

        XWPFTableRow headerRow = table.getRow(0);
        if (headerRow == null) {
            return false;
        }

        Map<String, Integer> columnIndexMap = new LinkedHashMap<>();
        List<XWPFTableCell> headerCells = headerRow.getTableCells();
        for (int i = 0; i < headerCells.size(); i++) {
            String colName = headerCells.get(i).getText().trim();
            if (!colName.isEmpty()) {
                columnIndexMap.put(colName, i);
            }
        }

        Integer serialCol = findColumnIndex(columnIndexMap, "序号", "序");
        Integer contractNoCol = findColumnIndex(columnIndexMap, "合同指标编号");
        Integer contractContentCol = findColumnIndex(columnIndexMap, "合同指标内容");
        Integer testItemCol = findColumnIndex(columnIndexMap, "测试项", "测试项名称");
        // 追踪关系列索引
        Integer traceRelationCol = findColumnIndex(columnIndexMap, "追踪关系", "trace", "追踪");

        if (serialCol == null && contractNoCol == null && contractContentCol == null && testItemCol == null && traceRelationCol == null) {
            System.out.println("测试项追踪表列头未识别，跳过该表");
            return false;
        }

        if (serialCol != null) {
            normalizeTokenColumnLayout(table, serialCol, 1200);
        }
        if (contractNoCol != null) {
            normalizeTokenColumnLayout(table, contractNoCol, 1800);
        }
        // 追踪关系列宽设置
        if (traceRelationCol != null) {
            normalizeTokenColumnLayout(table, traceRelationCol, 3000);
        }

        // 清空旧数据行，仅保留表头和一行模板数据行（若存在）
        if (table.getNumberOfRows() > 2) {
            for (int i = table.getNumberOfRows() - 1; i >= 2; i--) {
                table.removeRow(i);
            }
        }

        XWPFTableRow firstDataRow;
        if (table.getNumberOfRows() > 1) {
            firstDataRow = table.getRow(1);
            clearRowText(firstDataRow);
        } else {
            firstDataRow = table.createRow();
        }

        int serial = 1;
        for (int i = 0; i < traceRows.size(); i++) {
            XWPFTableRow row = (i == 0) ? firstDataRow : table.createRow();
            Map<String, String> rowData = traceRows.get(i);

            if (serialCol != null) {
                // 序号按1..N自动递增，并强制居中显示。
                setCellValueByIndex(row, serialCol, String.valueOf(serial), true, true);
                serial++;
            }
            if (contractNoCol != null) {
                setCellValueByIndex(row, contractNoCol, rowData.getOrDefault("合同指标编号", ""), false, true);
            }
            if (contractContentCol != null) {
                setCellValueByIndex(row, contractContentCol, rowData.getOrDefault("合同指标内容", ""), false, false);
            }
            if (testItemCol != null) {
                setCellValueByIndex(row, testItemCol, rowData.getOrDefault("测试项", ""), false, false);
            }
            // 填充追踪关系列
            if (traceRelationCol != null) {
                setCellValueByIndex(row, traceRelationCol, rowData.getOrDefault("追踪关系", ""), false, false);
            }
        }

        return true;
    }

    /** 按列索引写入单元格文本 */
    private void setCellValueByIndex(XWPFTableRow row, int columnIndex, String value,
            boolean centerText, boolean preventWrap) {
        if (row == null || columnIndex < 0) {
            return;
        }

        while (row.getTableCells().size() <= columnIndex) {
            row.createCell();
        }

        XWPFTableCell cell = row.getCell(columnIndex);
        setTraceabilityCellText(cell, value != null ? value : "", centerText, preventWrap);
    }

    /**
     * 追踪表专用写入：重建段落，避免模板残留的编号/缩进格式导致序号显示异常。
     */
    private void setTraceabilityCellText(XWPFTableCell cell, String text,
            boolean centerText, boolean preventWrap) {
        if (cell == null) {
            return;
        }

        while (!cell.getParagraphs().isEmpty()) {
            cell.removeParagraph(0);
        }

        XWPFParagraph para = cell.addParagraph();
        para.setAlignment(centerText ? ParagraphAlignment.CENTER : ParagraphAlignment.LEFT);
        para.setSpacingBefore(0);
        para.setSpacingAfter(0);

        XWPFRun run = para.createRun();
        String content = text != null ? text : "";
        run.setText(preventWrap ? toUnbreakableToken(content) : content);
        run.setColor("000000");

        if (centerText || preventWrap) {
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
            if (!tcPr.isSetNoWrap()) {
                tcPr.addNewNoWrap();
            }
        }
    }

    /**
     * 编号类列布局兜底：防止列过窄导致 token 被拆行。
     */
    private void normalizeTokenColumnLayout(XWPFTable table, int colIndex, int minWidthDxa) {
        if (table == null || colIndex < 0) {
            return;
        }

        for (XWPFTableRow row : table.getRows()) {
            if (row == null || row.getTableCells().size() <= colIndex) {
                continue;
            }

            XWPFTableCell cell = row.getCell(colIndex);
            if (cell == null) {
                continue;
            }

            CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();

            // 设置最小宽度，避免 token 过窄被拆行。
            if (!tcPr.isSetTcW()) {
                tcPr.addNewTcW().setType(STTblWidth.DXA);
                tcPr.getTcW().setW(java.math.BigInteger.valueOf(minWidthDxa));
            } else {
                tcPr.getTcW().setType(STTblWidth.DXA);
                tcPr.getTcW().setW(java.math.BigInteger.valueOf(minWidthDxa));
            }

            if (!tcPr.isSetNoWrap()) {
                tcPr.addNewNoWrap();
            }
        }
    }

    /**
     * 把 token 变成不可在字符间断行的文本。
     * 使用 WORD JOINER(U+2060) 连接每个字符，避免 13/5.2.3 被拆成多行。
     */
    private String toUnbreakableToken(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }
        if (text.length() == 1) {
            return text;
        }

        final char WORD_JOINER = '\u2060';
        StringBuilder sb = new StringBuilder(text.length() * 2);
        for (int i = 0; i < text.length(); i++) {
            sb.append(text.charAt(i));
            if (i < text.length() - 1) {
                sb.append(WORD_JOINER);
            }
        }
        return sb.toString();
    }

    /** 清空整行文本，保留单元格结构 */
    private void clearRowText(XWPFTableRow row) {
        if (row == null) {
            return;
        }
        for (XWPFTableCell cell : row.getTableCells()) {
            setCellText(cell, "");
        }
    }

    /** 在多个候选列名中查找列索引 */
    private Integer findColumnIndex(Map<String, Integer> columnIndexMap, String... candidates) {
        if (columnIndexMap == null || columnIndexMap.isEmpty()) {
            return null;
        }

        for (String candidate : candidates) {
            Integer exact = findColumnIndex(columnIndexMap, candidate);
            if (exact != null) {
                return exact;
            }
        }
        return null;
    }

    private String firstNonBlank(TestCase testCase, String... columnNames) {
        for (String columnName : columnNames) {
            String value = testCase.getColumnValue(columnName);
            if (!isBlank(value)) {
                return value.trim();
            }
        }
        return "";
    }

    private boolean isBlank(String value) {
        return value == null || value.trim().isEmpty();
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

    /** 设置单元格文本（设置黑色字体） */
    private void setCellText(XWPFTableCell cell, String text) {
        if (cell == null || text == null)
            return;

        // 清除多余段落，只保留一个
        while (cell.getParagraphs().size() > 1)
            cell.removeParagraph(cell.getParagraphs().size() - 1);
        if (cell.getParagraphs().isEmpty())
            cell.addParagraph();

        XWPFParagraph para = cell.getParagraphs().get(0);
        while (!para.getRuns().isEmpty())
            para.removeRun(0);

        // 添加新文本并设置黑色字体
        XWPFRun run = para.createRun();
        run.setText(text);
        run.setColor("000000"); // 设置字体颜色为黑色
    }
}
