package pub.developers.docautogenbyexcel.processor;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.TestCase;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Word模板处理模块
 * 负责读取Word模板，定位章节，插入表格和子标题
 */
public class WordProcessor {
    
    // 章节标题匹配模式：X.X 模块名称（X为数字）
    private static final Pattern SECTION_PATTERN = Pattern.compile("^(\\d+\\.\\d+)\\s+(.+)$");
    // 占位符匹配模式：X.x 表示该章节下有多个子章节
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("^(\\d+)\\.x\\s*(.+)?$", Pattern.CASE_INSENSITIVE);

    /**
     * 处理Word文档，在指定章节插入测试表格
     *
     * @param templatePath Word模板路径
     * @param outputPath   输出文件路径
     * @param moduleDataMap 模块数据Map
     * @return 处理成功的模块数量
     * @throws Exception 处理异常
     */
    public int processWord(String templatePath, String outputPath, 
                          java.util.Map<String, ModuleData> moduleDataMap) throws Exception {
        // 检查文件格式
        String lowerPath = templatePath.toLowerCase();
        if (lowerPath.endsWith(".doc") && !lowerPath.endsWith(".docx")) {
            throw new Exception("不支持旧版Word格式(.doc文件)。请将文件转换为.docx格式后再使用。\n" +
                    "转换方法：\n" +
                    "1. 使用Microsoft Word打开.doc文件\n" +
                    "2. 选择'文件' -> '另存为'\n" +
                    "3. 在'文件类型'中选择'Word文档(*.docx)'\n" +
                    "4. 保存后使用新的.docx文件");
        }
        
        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // 先扫描Word文档中的所有章节编号、占位符和已存在的子章节
            List<String> wordSectionNumbers = scanWordSections(document);
            List<PlaceholderInfo> placeholders = scanPlaceholders(document);
            List<SubSectionInfo> existingSubSections = scanExistingSubSections(document);
            
            System.out.println("Word文档中找到的章节编号: " + wordSectionNumbers);
            System.out.println("Word文档中找到的占位符: " + placeholders);
            System.out.println("Word文档中找到的已存在子章节: " + existingSubSections);
            
            // 获取Excel的列名列表（从第一个测试用例获取）
            List<String> excelColumnNames = new java.util.ArrayList<>();
            if (!moduleDataMap.isEmpty()) {
                ModuleData firstModule = moduleDataMap.values().iterator().next();
                if (!firstModule.getTestCases().isEmpty()) {
                    TestCase firstCase = firstModule.getTestCases().get(0);
                    excelColumnNames.addAll(firstCase.getColumnData().keySet());
                    excelColumnNames.remove("模块编号"); // 排除模块编号列
                    System.out.println("Excel列名: " + excelColumnNames);
                }
            }
            
            CTBody body = document.getDocument().getBody();
            
            // 保存全局模板表格（用于没有模板的章节）
            XWPFTable globalTemplateTable = null;
            
            // 查找并处理每个模块
            int successCount = 0;
            
            // 1. 处理已存在的子章节：填充表格并更新章节名称
            for (SubSectionInfo subSection : existingSubSections) {
                String moduleNumber = subSection.moduleNumber; // 如 "5.3.1"
                ModuleData moduleData = moduleDataMap.get(moduleNumber);
                
                if (moduleData == null) {
                    System.out.println("Excel中未找到模块: " + moduleNumber);
                    continue;
                }
                
                // 更新章节名称（如果Excel中的testName与Word中的不一致）
                String excelTestName = moduleData.getTestCases().get(0).getTestName();
                String expectedTitle = moduleNumber + " " + excelTestName + "测试";
                String currentTitle = subSection.paragraph.getText().trim();
                
                if (!currentTitle.equals(expectedTitle)) {
                    // 更新章节名称
                    updateParagraphText(subSection.paragraph, expectedTitle);
                    System.out.println("更新章节名称: " + currentTitle + " -> " + expectedTitle);
                }
                
                // 填充表格
                CTTbl existingTable = findTableCttblAfterParagraph(body, subSection.paragraph.getCTP());
                if (existingTable != null) {
                    XWPFTable table = new XWPFTable(existingTable, document);
                    fillTableData(table, moduleData.getTestCases().get(0));
                    System.out.println("模块" + moduleNumber + "表格填充完成");
                } else {
                    // 如果没有表格，创建新表格（使用全局模板）
                    insertTestCaseTableAfterParagraph(document, subSection.paragraph.getCTP(), moduleData.getTestCases().get(0), globalTemplateTable);
                    System.out.println("模块" + moduleNumber + "创建新表格完成");
                }
                
                successCount++;
            }
            
            // 2. 处理占位符：自动创建子章节
            for (PlaceholderInfo placeholder : placeholders) {
                String parentNumber = placeholder.parentNumber; // 如 "5"
                String placeholderText = placeholder.text; // 如 "5.x" 或 "5.x 功能测试"
                
                // 在Excel中查找所有匹配的子模块（如5.1、5.2、5.3）
                List<String> matchedModules = new java.util.ArrayList<>();
                for (String moduleNumber : moduleDataMap.keySet()) {
                    if (moduleNumber.startsWith(parentNumber + ".") && 
                        !moduleNumber.equals(parentNumber + ".x") &&
                        moduleNumber.matches("^" + parentNumber + "\\.\\d+$")) {
                        matchedModules.add(moduleNumber);
                    }
                }
                
                // 按编号排序
                matchedModules.sort((a, b) -> {
                    double numA = Double.parseDouble(a);
                    double numB = Double.parseDouble(b);
                    return Double.compare(numA, numB);
                });
                
                if (matchedModules.isEmpty()) {
                    System.out.println("占位符 " + placeholderText + " 在Excel中未找到匹配的子模块");
                    continue;
                }
                
                System.out.println("占位符 " + placeholderText + " 匹配到 " + matchedModules.size() + " 个子模块: " + matchedModules);
                
                // 在占位符位置创建子章节
                for (int i = 0; i < matchedModules.size(); i++) {
                    String moduleNumber = matchedModules.get(i);
                    ModuleData moduleData = moduleDataMap.get(moduleNumber);
                    
                    // 获取当前插入位置（第一次在占位符后，后续在前一个子章节后）
                    XWPFParagraph insertAfterPara;
                    if (i == 0) {
                        insertAfterPara = placeholder.paragraph;
                    } else {
                        // 需要找到前一个子章节的段落
                        // 由于我们已经创建了前一个子章节，需要重新查找
                        String prevModuleNumber = matchedModules.get(i - 1);
                        insertAfterPara = findSectionParagraph(document, prevModuleNumber);
                        if (insertAfterPara == null) {
                            // 如果找不到，使用占位符段落
                            insertAfterPara = placeholder.paragraph;
                        }
                    }
                    
                    // 创建子章节标题段落
                    XWPFParagraph subSectionPara = createSubSectionParagraph(
                        document, insertAfterPara, moduleNumber, moduleData.getTestCases().get(0).getTestName());
                    
                    // 在子章节后插入内容（使用全局模板）
                    insertModuleContent(document, subSectionPara, moduleNumber, moduleData, globalTemplateTable);
                    successCount++;
                    System.out.println("模块" + moduleNumber + "处理完成（生成" + 
                                     moduleData.getTestCaseCount() + "个表格）");
                }
                
                // 删除占位符段落（在所有子章节创建完成后）
                int pIndex = -1;
                CTP placeholderCTP = placeholder.paragraph.getCTP();
                for (int i = 0; i < body.sizeOfPArray(); i++) {
                    if (body.getPArray(i) == placeholderCTP) {
                        pIndex = i;
                        break;
                    }
                }
                if (pIndex != -1) {
                    body.removeP(pIndex);
                }
            }
            
            // 2. 处理占位符：自动创建子章节
            wordSectionNumbers = scanWordSections(document);
            for (String sectionNumber : wordSectionNumbers) {
                // 跳过已经被占位符处理的章节
                boolean isPlaceholderChild = false;
                for (PlaceholderInfo placeholder : placeholders) {
                    if (sectionNumber.startsWith(placeholder.parentNumber + ".")) {
                        isPlaceholderChild = true;
                        break;
                    }
                }
                if (isPlaceholderChild) {
                    continue;
                }
                
                ModuleData moduleData = moduleDataMap.get(sectionNumber);
                if (moduleData == null) {
                    System.out.println("Excel中未找到模块: " + sectionNumber);
                    continue;
                }
                
                // 查找章节标题位置
                XWPFParagraph sectionPara = findSectionParagraph(document, sectionNumber);
                if (sectionPara == null) {
                    System.out.println("Word模板中未找到模块: " + sectionNumber);
                    continue;
                }

                // 检查章节后是否已有表格
                CTTbl existingTable = findTableCttblAfterParagraph(body, sectionPara.getCTP());
                
                // 找到该章节的结束位置（用于插入子章节）
                XWPFParagraph lastElementInSection = findLastElementInSection(document, sectionPara, sectionNumber);
                System.out.println("模块" + sectionNumber + "的最后元素: " + (lastElementInSection != null ? lastElementInSection.getText() : "null"));
                
                // 保存第一个表格作为模板（如果存在）
                XWPFTable templateTable = null;
                if (existingTable != null) {
                    templateTable = new XWPFTable(existingTable, document);
                    if (globalTemplateTable == null) {
                        globalTemplateTable = templateTable;
                        System.out.println("保存模块" + sectionNumber + "的表格作为全局模板");
                    }
                }
                
                // 处理所有测试用例，按顺序创建子章节和表格
                List<TestCase> testCases = moduleData.getTestCases();
                
                // 找到该章节下已存在的子章节（模板子章节，如 XX测试）
                List<XWPFParagraph> existingSubSectionParas = findExistingSubSectionsInSection(
                    document, sectionPara, sectionNumber);
                System.out.println("模块" + sectionNumber + "下已有" + existingSubSectionParas.size() + "个子章节");
                
                XWPFParagraph currentInsertPoint = lastElementInSection != null ? lastElementInSection : sectionPara;
                System.out.println("初始插入点: " + (currentInsertPoint != null ? currentInsertPoint.getText() : "null"));
                
                for (int i = 0; i < testCases.size(); i++) {
                    TestCase testCase = testCases.get(i);
                    int sequenceNumber = i + 1;
                    String subSectionNumber = sectionNumber + "." + sequenceNumber;
                    
                    // 检查是否有可复用的已存在子章节（按顺序复用）
                    if (i < existingSubSectionParas.size()) {
                        // 复用已存在的子章节：修改标题和填充表格
                        XWPFParagraph existingSubSection = existingSubSectionParas.get(i);
                        System.out.println("替换已存在的子章节为: " + subSectionNumber + " " + testCase.getTestName() + "测试");
                        
                        // 修改子章节标题
                        updateParagraphText(existingSubSection, subSectionNumber + " " + testCase.getTestName() + "测试");
                        
                        // 填充表格
                        CTTbl tableAfterSub = findTableCttblAfterParagraph(body, existingSubSection.getCTP());
                        if (tableAfterSub != null) {
                            XWPFTable table = new XWPFTable(tableAfterSub, document);
                            fillTableData(table, testCase);
                            System.out.println("子章节" + subSectionNumber + "表格填充完成");
                            
                            // 更新表格前的Caption（如果有的话）
                            updateTableCaption(document, existingSubSection, subSectionNumber, testCase.getTestName());
                        }
                        // 更新插入点为表格后面（确保新子章节在表格后面创建）
                        currentInsertPoint = findInsertPointAfterTable(document, existingSubSection);
                    } else {
                        // 没有更多已存在的子章节，创建新的
                        System.out.println("创建子章节: " + subSectionNumber + " " + testCase.getTestName() + "测试");
                        System.out.println("当前插入点: " + (currentInsertPoint != null ? currentInsertPoint.getText() : "null"));
                        XWPFParagraph subSectionPara = createSubSectionParagraph(
                            document, currentInsertPoint, subSectionNumber, testCase.getTestName());
                        
                        // 始终为新创建的子章节创建新表格（不复用已存在的表格）
                        if (templateTable != null) {
                            // 如果有模板，复制模板表格
                            System.out.println("复制模板表格到子章节" + subSectionNumber + "后");
                            CTTbl newCttbl = copyTable(document, subSectionPara.getCTP(), templateTable.getCTTbl());
                            if (newCttbl != null) {
                                XWPFTable newTable = new XWPFTable(newCttbl, document);
                                System.out.println("模板表格复制成功，行数: " + newTable.getNumberOfRows());
                                
                                // 清空数据列（只保留标签列）
                                for (int rowIdx = 0; rowIdx < newTable.getNumberOfRows(); rowIdx++) {
                                    XWPFTableRow row = newTable.getRow(rowIdx);
                                    if (row == null) continue;
                                    
                                    int cellCount = row.getTableCells().size();
                                    if (rowIdx == 0 && cellCount >= 4) {
                                        // 第一行：清空单元格1和3（数据列）
                                        clearCellContent(row.getCell(1));
                                        clearCellContent(row.getCell(3));
                                    } else if (cellCount >= 2) {
                                        // 其他行：清空单元格1及后续单元格（数据列）
                                        for (int j = 1; j < cellCount; j++) {
                                            clearCellContent(row.getCell(j));
                                        }
                                    }
                                }
                                
                                // 填充表格数据
                                fillTableData(newTable, testCase);
                                System.out.println("表格数据填充完成");
                            } else {
                                System.err.println("复制模板表格失败");
                            }
                        } else {
                            // 如果没有模板，创建新表格
                            insertNewTableAfterParagraph(document, subSectionPara.getCTP(), testCase);
                        }
                        
                        // 更新插入点为当前子章节
                        currentInsertPoint = subSectionPara;
                    }
                }
                
                System.out.println("模块" + sectionNumber + "处理完成（生成" + testCases.size() + "个表格）");
                successCount++;
            }
            
            // 3. 处理Excel中有但Word中没有的模块（可选）
            for (java.util.Map.Entry<String, ModuleData> entry : moduleDataMap.entrySet()) {
                String moduleNumber = entry.getKey();
                boolean found = wordSectionNumbers.contains(moduleNumber);
                for (PlaceholderInfo placeholder : placeholders) {
                    if (moduleNumber.startsWith(placeholder.parentNumber + ".")) {
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    System.out.println("警告：Word文档中未找到章节 " + moduleNumber + "，跳过处理");
                }
            }

            // 保存文档
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }

            return successCount;

        } catch (OLE2NotOfficeXmlFileException e) {
            throw new Exception("不支持旧版Word格式(.doc文件)。请将文件转换为.docx格式后再使用。\n" +
                    "转换方法：\n" +
                    "1. 使用Microsoft Word打开.doc文件\n" +
                    "2. 选择'文件' -> '另存为'\n" +
                    "3. 在'文件类型'中选择'Word文档(*.docx)'\n" +
                    "4. 保存后使用新的.docx文件", e);
        } catch (IOException e) {
            throw new Exception("处理Word文档失败: " + e.getMessage(), e);
        }
    }

    /**
     * 子章节信息
     */
    private static class SubSectionInfo {
        String moduleNumber; // 模块编号，如 "5.3.1"
        XWPFParagraph paragraph; // 子章节段落
        
        SubSectionInfo(String moduleNumber, XWPFParagraph paragraph) {
            this.moduleNumber = moduleNumber;
            this.paragraph = paragraph;
        }
        
        @Override
        public String toString() {
            return moduleNumber;
        }
    }

    /**
     * 占位符信息
     */
    private static class PlaceholderInfo {
        String parentNumber; // 父章节编号，如 "5"
        String text; // 占位符文本，如 "5.x 功能测试"
        XWPFParagraph paragraph; // 占位符段落
        
        PlaceholderInfo(String parentNumber, String text, XWPFParagraph paragraph) {
            this.parentNumber = parentNumber;
            this.text = text;
            this.paragraph = paragraph;
        }
        
        @Override
        public String toString() {
            return text;
        }
    }
    
    /**
     * 扫描Word文档中已存在的子章节（如5.3.1、5.3.2等）
     * 从目录读取子章节编号，然后在正文中查找对应段落
     *
     * @param document Word文档
     * @return 子章节信息列表
     */
    private List<SubSectionInfo> scanExistingSubSections(XWPFDocument document) {
        List<SubSectionInfo> subSections = new java.util.ArrayList<>();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        // 目录样式 ID
        java.util.Set<String> tocStyles = new java.util.HashSet<>();
        tocStyles.add("22");
        tocStyles.add("25");
        tocStyles.add("16");
        
        // 正文标题样式 ID: 2=Heading1, 3=Heading2, 4=Heading3, 5=Heading4
        java.util.Set<String> headingStyles = new java.util.HashSet<>();
        headingStyles.add("2");  // Heading 1
        headingStyles.add("3");  // Heading 2
        headingStyles.add("4");  // Heading 3 (子章节 X.X.X)
        headingStyles.add("5");  // Heading 4
        
        // 存储目录中的子章节: 编号 -> 名称
        java.util.Map<String, String> tocSubSections = new java.util.LinkedHashMap<>();
        
        // 子章节匹配模式：X.X.X 模块名称
        Pattern subSectionPattern = Pattern.compile("^(\\d+\\.\\d+\\.\\d+)\\s+(.+)$");
        
        // 第一遍：从目录中读取子章节编号和名称
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }

                String styleName = para.getStyle();
                boolean isToc = styleName != null && 
                    (tocStyles.contains(styleName) || styleName.toLowerCase().startsWith("toc"));
                
                if (!isToc) {
                    continue;
                }

                // 匹配子章节格式：X.X.X 模块名称
                Matcher matcher = subSectionPattern.matcher(text.trim());
                if (matcher.matches()) {
                    String sectionNumber = matcher.group(1);
                    String sectionName = matcher.group(2).trim();
                    // 去除页码
                    int tabIndex = sectionName.indexOf('\t');
                    if (tabIndex > 0) {
                        sectionName = sectionName.substring(0, tabIndex).trim();
                    }
                    tocSubSections.put(sectionNumber, sectionName);
                }
            } catch (Exception e) {
                continue;
            }
        }
        
        // 第二遍：在正文中查找对应的段落
        for (java.util.Map.Entry<String, String> entry : tocSubSections.entrySet()) {
            String subNumber = entry.getKey();
            String subName = entry.getValue();
            
            for (XWPFParagraph para : paragraphs) {
                try {
                    String text = para.getText();
                    if (text == null || text.trim().isEmpty()) {
                        continue;
                    }
                    
                    String styleName = para.getStyle();
                    
                    // 跳过目录项
                    boolean isToc = styleName != null && 
                        (tocStyles.contains(styleName) || styleName.toLowerCase().startsWith("toc"));
                    if (isToc) {
                        continue;
                    }
                    
                    // 只在标题样式中查找
                    boolean isHeading = styleName != null && 
                        (headingStyles.contains(styleName) || 
                         styleName.toLowerCase().contains("heading") ||
                         styleName.contains("标题") ||
                         styleName.contains("程序标题"));
                    
                    String trimmed = text.trim();
                    // 匹配带编号或仅名称
                    if ((trimmed.startsWith(subNumber + " ") || trimmed.equals(subName)) && isHeading) {
                        subSections.add(new SubSectionInfo(subNumber, para));
                        System.out.println("找到正文中的子章节: " + subNumber + " -> " + trimmed);
                        break;
                    }
                } catch (Exception e) {
                    continue;
                }
            }
        }
        
        return subSections;
    }
    
    /**
     * 更新段落文本
     */
    private void updateParagraphText(XWPFParagraph paragraph, String newText) {
        // 清除原有内容
        while (paragraph.getRuns().size() > 0) {
            paragraph.removeRun(0);
        }
        
        // 添加新文本
        XWPFRun run = paragraph.createRun();
        run.setText(newText);
        run.setBold(true);
        run.setFontSize(12);
    }
    
    /**
     * 更新表格标题（Caption）
     * 查找子章节后面的Caption样式段落并更新
     */
    private void updateTableCaption(XWPFDocument document, XWPFParagraph subSectionPara, 
                                     String subSectionNumber, String testName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        // 找到子章节在列表中的位置
        int subSectionIndex = -1;
        for (int i = 0; i < paragraphs.size(); i++) {
            if (paragraphs.get(i).getCTP() == subSectionPara.getCTP()) {
                subSectionIndex = i;
                break;
            }
        }
        
        if (subSectionIndex == -1) {
            return;
        }
        
        // 查找子章节后的第一个Caption样式段落
        for (int i = subSectionIndex + 1; i < paragraphs.size() && i < subSectionIndex + 5; i++) {
            XWPFParagraph para = paragraphs.get(i);
            String styleName = para.getStyle();
            
            // 检查是否是Caption样式 (style_id="11")
            if (styleName != null && (styleName.equalsIgnoreCase("Caption") || 
                styleName.equals("11") || styleName.contains("题注"))) {
                // 更新Caption文本
                String newCaption = "表" + subSectionNumber + " " + testName + "测试";
                updateParagraphText(para, newCaption);
                System.out.println("更新表格标题为: " + newCaption);
                break;
            }
            
            // 如果遇到表格或其他Heading，停止查找
            String text = para.getText();
            if (text != null && !text.trim().isEmpty()) {
                if (styleName != null && (styleName.equals("3") || styleName.equals("4") || 
                    styleName.toLowerCase().contains("heading"))) {
                    break;
                }
            }
        }
    }

    /**
     * 扫描Word文档中的占位符（如 "5.x"）
     *
     * @param document Word文档
     * @return 占位符信息列表
     */
    private List<PlaceholderInfo> scanPlaceholders(XWPFDocument document) {
        List<PlaceholderInfo> placeholders = new java.util.ArrayList<>();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph para : paragraphs) {
            String text = para.getText();
            
            if (text == null || text.trim().isEmpty()) {
                continue;
            }

            // 匹配占位符格式：X.x 或 X.x 模块名称
            Matcher matcher = PLACEHOLDER_PATTERN.matcher(text.trim());
            if (matcher.matches()) {
                String parentNumber = matcher.group(1); // 如 "5"
                placeholders.add(new PlaceholderInfo(parentNumber, text.trim(), para));
            }
        }
        
        return placeholders;
    }
    
    /**
     * 创建子章节标题段落
     * 设置 Heading 3 样式以便后续识别
     *
     * @param document Word文档
     * @param afterPara 在哪个段落后插入
     * @param moduleNumber 模块编号，如 "5.2.2"
     * @param testName 测试名称，如 "登录功能"
     * @return 创建的子章节段落
     */
    private XWPFParagraph createSubSectionParagraph(XWPFDocument document, XWPFParagraph afterPara, 
                                                     String moduleNumber, String testName) {
        CTBody body = document.getDocument().getBody();
        CTP afterCTP = afterPara.getCTP();
        
        // 找到afterPara在body中的位置
        int afterIndex = -1;
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i) == afterCTP) {
                afterIndex = i;
                break;
            }
        }
        
        if (afterIndex == -1) {
            System.err.println("找不到插入点段落，使用末尾插入");
            afterIndex = body.sizeOfPArray() - 1;
        }
        
        // 在afterIndex+1位置插入新段落
        CTP ctp = body.insertNewP(afterIndex + 1);
        XWPFParagraph para = new XWPFParagraph(ctp, document);

        // 不使用Heading样式（避免自动编号导致双编号问题）
        // 直接设置格式为加粗、小四号字体
        
        // 设置标题内容：编号 + 名称
        String subTitle = moduleNumber + " " + testName + "测试";
        XWPFRun run = para.createRun();
        run.setText(subTitle);
        run.setBold(true);
        run.setFontSize(12); // 小四号字体
        run.setFontFamily("宋体"); // 设置字体

        // 设置段落格式
        para.setAlignment(ParagraphAlignment.LEFT);
        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(120)); // 段后间距
        
        System.out.println("在位置 " + (afterIndex + 1) + " 插入子章节: " + subTitle);
        
        return para;
    }

    /**
     * 扫描Word文档中的所有章节编号
     * 从目录（TOC）中读取章节编号（因为正文中的章节标题可能没有编号）
     *
     * @param document Word文档
     * @return 章节编号列表，如 ["4.3", "5.3", "6.2"]
     */
    private List<String> scanWordSections(XWPFDocument document) {
        List<String> sectionNumbers = new java.util.ArrayList<>();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        // 只从目录项（样式 ID 22, 25, 16 等）中读取章节编号
        // 同时建立章节编号到章节名称的映射
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }

                String styleName = para.getStyle();
                
                // 目录样式 ID: 22=toc1, 25=toc2, 16=toc3
                boolean isToc = styleName != null && 
                    (styleName.equals("22") || styleName.equals("25") || styleName.equals("16") ||
                     styleName.toLowerCase().startsWith("toc"));
                
                if (!isToc) {
                    continue;
                }
                
                String trimmed = text.trim();

                // 匹配章节标题格式：X.X 模块名称
                Matcher matcher = SECTION_PATTERN.matcher(trimmed);
                if (matcher.matches()) {
                    String sectionNumber = matcher.group(1);
                    // 只添加主章节（X.X格式），不添加子章节（X.X.X格式）
                    if (sectionNumber.matches("^\\d+\\.\\d+$")) {
                        if (!sectionNumbers.contains(sectionNumber)) {
                            sectionNumbers.add(sectionNumber);
                        }
                    }
                }
            } catch (org.apache.xmlbeans.impl.values.XmlValueDisconnectedException e) {
                continue;
            } catch (Exception e) {
                continue;
            }
        }
        
        return sectionNumbers;
    }

    /**
     * 查找章节标题段落
     * 策略：1. 从目录(TOC)获取章节名称 2. 在正文中通过名称和Heading样式匹配
     *
     * @param document     Word文档
     * @param moduleNumber 模块编号，如"5.3"或"4.3"
     * @return 段落对象，未找到返回null
     */
    private XWPFParagraph findSectionParagraph(XWPFDocument document, String moduleNumber) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        String sectionNameFromToc = null;  // 从目录中提取的章节名称
        
        // 目录样式 ID: 22=toc1, 25=toc2, 16=toc3
        java.util.Set<String> tocStyles = new java.util.HashSet<>();
        tocStyles.add("22");
        tocStyles.add("25");
        tocStyles.add("16");
        
        // 正文标题样式 ID: 2=Heading1, 3=Heading2, 4=Heading3, 5=Heading4
        java.util.Set<String> headingStyles = new java.util.HashSet<>();
        headingStyles.add("2");  // Heading 1
        headingStyles.add("3");  // Heading 2 (主章节 X.X)
        headingStyles.add("4");  // Heading 3 (子章节 X.X.X)
        headingStyles.add("5");  // Heading 4
        
        // 第一遍：从目录中提取章节名称
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }

                String styleName = para.getStyle();
                boolean isToc = styleName != null && 
                    (tocStyles.contains(styleName) || styleName.toLowerCase().startsWith("toc"));
                
                if (isToc) {
                    String trimmedText = text.trim();
                    Matcher matcher = SECTION_PATTERN.matcher(trimmedText);
                    if (matcher.matches()) {
                        String foundModuleNumber = matcher.group(1);
                        if (moduleNumber.equals(foundModuleNumber)) {
                            sectionNameFromToc = matcher.group(2).trim();
                            // 去除页码（通常以\t分隔）
                            int tabIndex = sectionNameFromToc.indexOf('\t');
                            if (tabIndex > 0) {
                                sectionNameFromToc = sectionNameFromToc.substring(0, tabIndex).trim();
                            }
                            System.out.println("从目录提取章节名称: " + moduleNumber + " -> " + sectionNameFromToc);
                            break;
                        }
                    }
                }
            } catch (Exception e) {
                continue;
            }
        }
        
        // 第二遍：在正文中查找匹配的段落（只在Heading样式中查找）
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }
                
                String styleName = para.getStyle();
                
                // 跳过目录项
                if (styleName != null && 
                    (tocStyles.contains(styleName) || styleName.toLowerCase().startsWith("toc"))) {
                    continue;
                }
                
                // 只在正文标题样式中查找
                boolean isHeading = styleName != null && 
                    (headingStyles.contains(styleName) || 
                     styleName.toLowerCase().contains("heading") || 
                     styleName.contains("标题"));
                
                String trimmedText = text.trim();
                
                // 方法1：尝试通过编号匹配（如果正文有编号）
                Matcher matcher = SECTION_PATTERN.matcher(trimmedText);
                if (matcher.matches()) {
                    String foundModuleNumber = matcher.group(1);
                    if (moduleNumber.equals(foundModuleNumber)) {
                        System.out.println("在正文中找到章节(通过编号): " + trimmedText + " [" + styleName + "]");
                        return para;
                    }
                }
                
                // 方法2：通过章节名称 + Heading样式匹配
                if (sectionNameFromToc != null && trimmedText.equals(sectionNameFromToc) && isHeading) {
                    System.out.println("在正文中找到章节(通过名称+样式): " + trimmedText + " [" + styleName + "]");
                    return para;
                }
            } catch (org.apache.xmlbeans.impl.values.XmlValueDisconnectedException e) {
                continue;
            } catch (Exception e) {
                continue;
            }
        }
        
        return null;
    }

    /**
     * 在指定章节段落后插入模块内容（子标题和表格）
     *
     * @param document     Word文档
     * @param sectionPara 章节段落
     * @param moduleNumber 模块编号
     * @param moduleData   模块数据
     */
    private void insertModuleContent(XWPFDocument document, XWPFParagraph sectionPara,
                                     String moduleNumber, ModuleData moduleData, XWPFTable templateTable) {
        List<TestCase> testCases = moduleData.getTestCases();

        if (testCases.isEmpty()) {
            return;
        }

        // 获取body和sectionPara的CTP
        CTBody body = document.getDocument().getBody();
        
        // 使用顺序插入，而不是倒序
        XWPFParagraph lastInsertPara = sectionPara;
        
        for (int i = 0; i < testCases.size(); i++) {
            TestCase testCase = testCases.get(i);
            int sequenceNumber = i + 1;

            // 1. 创建子标题
            String subSectionNumber = moduleNumber + "." + sequenceNumber;
            String subTitle = subSectionNumber + " " + testCase.getTestName() + "测试";
            XWPFParagraph subSectionPara = createSubSectionParagraph(
                document, lastInsertPara, subSectionNumber, testCase.getTestName());

            // 2. 然后在子标题后插入表格（使用模板复制）
            insertTestCaseTableAfterParagraph(document, subSectionPara.getCTP(), testCase, templateTable);
            
            // 3. 找到表格后的位置作为下次插入点
            CTTbl tableAfterSub = findTableCttblAfterParagraph(body, subSectionPara.getCTP());
            if (tableAfterSub != null) {
                // 在表格后创建空段落作为下次插入点
                org.apache.xmlbeans.XmlCursor tableCursor = tableAfterSub.newCursor();
                tableCursor.toEndToken();
                tableCursor.toNextToken();
                
                tableCursor.beginElement(
                    new javax.xml.namespace.QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p", "w"));
                tableCursor.toParent();
                
                CTP newEmptyPara = null;
                if (tableCursor.getObject() instanceof CTP) {
                    newEmptyPara = (CTP) tableCursor.getObject();
                }
                
                tableCursor.close();
                
                if (newEmptyPara != null) {
                    lastInsertPara = new XWPFParagraph(newEmptyPara, document);
                } else {
                    lastInsertPara = subSectionPara;
                }
            } else {
                lastInsertPara = subSectionPara;
            }
        }
    }

    /**
     * 在指定位置插入子标题
     * @return 插入的段落CTP对象
     */
    private CTP insertSubTitleAt(XWPFDocument document, CTBody body, int position, String title) {
        // 确保位置不超过P数组长度
        int maxPos = body.sizeOfPArray();
        int insertPos = Math.min(position, maxPos);
        
        CTP ctp = body.insertNewP(insertPos);
        XWPFParagraph para = new XWPFParagraph(ctp, document);

        // 设置标题样式：小四、加粗
        XWPFRun run = para.createRun();
        run.setText(title);
        run.setBold(true);
        run.setFontSize(12); // 小四号字体

        // 设置段落格式
        para.setAlignment(ParagraphAlignment.LEFT);
        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(120)); // 段后间距
        
        return ctp;
    }

    /**
     * 在段落后强制创建新表格（不查找已存在的表格）
     */
    private void insertNewTableAfterParagraph(XWPFDocument document, CTP paragraph, TestCase testCase) {
        CTBody body = document.getDocument().getBody();
        
        // 使用XmlCursor在段落后直接插入表格元素
        org.apache.xmlbeans.XmlCursor paraCursor = paragraph.newCursor();
        paraCursor.toEndToken();
        paraCursor.toNextToken();
        
        // 在cursor位置插入新的表格元素
        paraCursor.beginElement(
            new javax.xml.namespace.QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "tbl", "w"));
        paraCursor.toParent();
        
        // 获取刚插入的CTTbl对象
        CTTbl cttbl = null;
        if (paraCursor.getObject() instanceof CTTbl) {
            cttbl = (CTTbl) paraCursor.getObject();
        }
        
        paraCursor.close();
        
        if (cttbl == null) {
            System.err.println("创建表格失败");
            return;
        }
        
        // 创建XWPFTable对象
        XWPFTable table = new XWPFTable(cttbl, document);
        System.out.println("创建新表格，段落: " + (paragraph != null ? "非空" : "空"));
        
        // 创建表格结构：第1行4列，后续行2列
        // 第1行：4列（测试项名称、testName值、标识、id值）
        XWPFTableRow firstRow = table.createRow();
        for (int i = 0; i < 4; i++) {
            firstRow.createCell();
        }
        setCellValue(firstRow.getCell(0), "测试项名称");
        setCellValue(firstRow.getCell(2), "标识");
        // testName和id的值会在fillTableData中填充
        
        // 后续行：2列（标签、数据，数据列需要合并3列）
        List<String> labels = new java.util.ArrayList<>();
        labels.add("测试内容");
        labels.add("测试策略与方法");
        labels.add("判定准则");
        labels.add("测试终止条件");
        labels.add("追踪关系");
        
        for (String label : labels) {
            XWPFTableRow row = table.createRow();
            // 创建第1列（标签列）
            XWPFTableCell labelCell = row.createCell();
            setCellValue(labelCell, label);
            // 创建第2列（数据列，需要合并3列）
            XWPFTableCell dataCell = row.createCell();
            // 设置gridSpan=3，使第2列跨越3个网格列
            CTTcPr tcPr = dataCell.getCTTc().getTcPr();
            if (tcPr == null) {
                tcPr = dataCell.getCTTc().addNewTcPr();
            }
            if (!tcPr.isSetGridSpan()) {
                tcPr.addNewGridSpan().setVal(BigInteger.valueOf(3));
            } else {
                tcPr.getGridSpan().setVal(BigInteger.valueOf(3));
            }
        }
        
        // 填充表格数据
        System.out.println("开始填充表格数据");
        fillTableData(table, testCase);
        System.out.println("表格数据填充完成，行数: " + table.getNumberOfRows());
        
        // 设置表格样式
        styleTable(cttbl);
        
        // 调试信息
        try {
            XWPFParagraph para = new XWPFParagraph(paragraph, document);
            String paraText = para.getText();
            System.out.println("表格已创建在段落后: " + (paraText != null ? paraText : "段落"));
        } catch (Exception e) {
            System.out.println("表格已创建在段落后");
        }
    }
    
    /**
     * 在段落后插入测试用例表格（会查找已存在的表格，或复制模板表格）
     */
    private void insertTestCaseTableAfterParagraph(XWPFDocument document, CTP paragraph, TestCase testCase, XWPFTable templateTable) {
        CTBody body = document.getDocument().getBody();
        
        // 检查段落后是否已经存在表格
        CTTbl existingCttbl = findTableCttblAfterParagraph(body, paragraph);
        
        if (existingCttbl != null) {
            // 如果已存在表格，直接填充数据
            try {
                XWPFParagraph para = new XWPFParagraph(paragraph, document);
                System.out.println("找到已存在的表格，段落: " + para.getText() + "，填充数据");
            } catch (Exception e) {
                System.out.println("找到已存在的表格，填充数据");
            }
            XWPFTable existingTable = new XWPFTable(existingCttbl, document);
            System.out.println("表格行数: " + existingTable.getNumberOfRows());
            fillTableData(existingTable, testCase);
            System.out.println("表格数据填充完成");
            // 不重新设置样式，保留原有格式
        } else if (templateTable != null) {
            // 如果不存在表格但有模板，复制模板表格
            try {
                XWPFParagraph para = new XWPFParagraph(paragraph, document);
                System.out.println("复制模板表格并填充数据，段落: " + para.getText());
            } catch (Exception e) {
                System.out.println("复制模板表格并填充数据");
            }
            
            // 复制模板表格
            CTTbl newCttbl = copyTable(document, paragraph, templateTable.getCTTbl());
            if (newCttbl != null) {
                XWPFTable newTable = new XWPFTable(newCttbl, document);
                System.out.println("模板表格复制成功，行数: " + newTable.getNumberOfRows());
                
                // 清空数据列（只保留标签列）
                for (int i = 0; i < newTable.getNumberOfRows(); i++) {
                    XWPFTableRow row = newTable.getRow(i);
                    if (row == null) continue;
                    
                    int cellCount = row.getTableCells().size();
                    if (i == 0 && cellCount >= 4) {
                        // 第一行：清空单元格1和3（数据列）
                        clearCellContent(row.getCell(1));
                        clearCellContent(row.getCell(3));
                    } else if (cellCount >= 2) {
                        // 其他行：清空单元格1及后续单元格（数据列）
                        for (int j = 1; j < cellCount; j++) {
                            clearCellContent(row.getCell(j));
                        }
                    }
                }
                
                fillTableData(newTable, testCase);
                System.out.println("表格数据填充完成");
            } else {
                System.err.println("复制模板表格失败");
            }
        } else {
            // 如果不存在表格且没有模板，输出警告
            try {
                XWPFParagraph para = new XWPFParagraph(paragraph, document);
                System.err.println("警告：段落后未找到表格且没有模板，无法填充数据。段落: " + para.getText());
            } catch (Exception e) {
                System.err.println("警告：段落后未找到表格且没有模板，无法填充数据");
            }
        }
    }
    
    /**
     * 复制表格到指定段落后
     */
    private CTTbl copyTable(XWPFDocument document, CTP afterParagraph, CTTbl sourceTable) {
        try {
            // 使用XmlCursor在段落后插入表格
            org.apache.xmlbeans.XmlCursor paraCursor = afterParagraph.newCursor();
            paraCursor.toEndToken();
            paraCursor.toNextToken();
            
            // 插入新的表格元素
            paraCursor.beginElement(
                new javax.xml.namespace.QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "tbl", "w"));
            paraCursor.toParent();
            
            // 获取刚插入的CTTbl对象
            CTTbl newTable = null;
            if (paraCursor.getObject() instanceof CTTbl) {
                newTable = (CTTbl) paraCursor.getObject();
            }
            
            paraCursor.close();
            
            if (newTable == null) {
                return null;
            }
            
            // 复制表格内容（行和单元格结构）
            for (int i = 0; i < sourceTable.sizeOfTrArray(); i++) {
                CTRow sourceRow = sourceTable.getTrArray(i);
                CTRow newRow = newTable.addNewTr();
                
                // 复制行的属性
                if (sourceRow.isSetTrPr()) {
                    newRow.setTrPr((CTTrPr) sourceRow.getTrPr().copy());
                }
                
                // 复制单元格
                for (int j = 0; j < sourceRow.sizeOfTcArray(); j++) {
                    CTTc sourceCell = sourceRow.getTcArray(j);
                    CTTc newCell = newRow.addNewTc();
                    
                    // 复制单元格属性（包括合并信息）
                    if (sourceCell.getTcPr() != null) {
                        newCell.setTcPr((CTTcPr) sourceCell.getTcPr().copy());
                    }
                    
                    // 只复制标签列的内容
                    // 对于第一行：复制单元格0和2（测试项名称、标识）
                    // 对于其他行：复制单元格0（标签）
                    boolean shouldCopyContent = (i == 0 && (j == 0 || j == 2)) || (i > 0 && j == 0);
                    
                    if (shouldCopyContent) {
                        // 复制内容
                        for (int k = 0; k < sourceCell.sizeOfPArray(); k++) {
                            CTP sourceP = sourceCell.getPArray(k);
                            CTP newP = newCell.addNewP();
                            
                            // 复制段落属性
                            if (sourceP.getPPr() != null) {
                                newP.setPPr((CTPPr) sourceP.getPPr().copy());
                            }
                            
                            // 复制段落内容（run）
                            for (int m = 0; m < sourceP.sizeOfRArray(); m++) {
                                CTR sourceR = sourceP.getRArray(m);
                                CTR newR = newP.addNewR();
                                
                                // 复制run属性
                                if (sourceR.getRPr() != null) {
                                    newR.setRPr((CTRPr) sourceR.getRPr().copy());
                                }
                                
                                // 复制文本内容
                                for (int n = 0; n < sourceR.sizeOfTArray(); n++) {
                                    CTText sourceT = sourceR.getTArray(n);
                                    CTText newT = newR.addNewT();
                                    newT.setStringValue(sourceT.getStringValue());
                                    if (sourceT.getSpace() != null) {
                                        newT.setSpace(sourceT.getSpace());
                                    }
                                }
                            }
                        }
                    } else {
                        // 创建空段落
                        newCell.addNewP();
                    }
                }
            }
            
            // 复制表格属性
            if (sourceTable.getTblPr() != null) {
                newTable.setTblPr((CTTblPr) sourceTable.getTblPr().copy());
            }
            
            // 复制表格网格
            if (sourceTable.getTblGrid() != null) {
                newTable.setTblGrid((CTTblGrid) sourceTable.getTblGrid().copy());
            }
            
            return newTable;
        } catch (Exception e) {
            System.err.println("复制表格失败: " + e.getMessage());
            e.printStackTrace();
            return null;
        }
    }
    
    /**
     * 查找段落后表格的CTTbl对象
     * 支持跳过空行段落查找表格
     * 只查找紧邻段落的表格，不查找距离较远的表格
     */
    private CTTbl findTableCttblAfterParagraph(CTBody body, CTP paragraph) {
        // 找到段落在body中的位置
        int pIndex = -1;
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i) == paragraph) {
                pIndex = i;
                break;
            }
        }
        
        if (pIndex == -1) {
            return null;
        }
        
        // 调试信息：检查段落内容
        try {
            XWPFParagraph para = new XWPFParagraph(paragraph, null);
            String paraText = para.getText();
            // 如果段落是子章节标题（如"4.3.2 注册功能测试"），且是新创建的，后面不应该有表格
            if (paraText != null && paraText.trim().matches("^\\d+\\.\\d+\\.\\d+\\s+.*")) {
                // 这是子章节，检查是否是刚创建的（通过检查后面是否有表格）
                // 如果是新创建的子章节，后面不应该有表格，应该返回null
            }
        } catch (Exception e) {
            // 忽略异常
        }
        
        // 方法1：使用XML cursor查找紧邻的表格（只查找直接相邻的表格，跳过最多1个空行段落）
        org.apache.xmlbeans.XmlCursor cursor = paragraph.newCursor();
        cursor.toEndToken();
        
        int emptyParaCount = 0;
        // 移动到下一个兄弟元素（跳过空行段落，但最多跳过1个）
        while (cursor.toNextSibling()) {
            org.apache.xmlbeans.XmlObject obj = cursor.getObject();
            
            // 检查是否是表格
            if (obj instanceof CTTbl) {
                cursor.close();
                return (CTTbl) obj;
            }
            
            // 如果是段落，检查是否为空行
            if (obj instanceof CTP) {
                CTP nextPara = (CTP) obj;
                String paraText = "";
                try {
                    XWPFParagraph xwpfPara = new XWPFParagraph(nextPara, null);
                    paraText = xwpfPara.getText();
                } catch (Exception e) {
                    // 忽略异常，继续查找
                }
                // 如果段落不为空，停止查找（表格应该在空行之后，但中间不应该有其他内容段落）
                if (paraText != null && !paraText.trim().isEmpty()) {
                    break;
                } else {
                    // 空行段落，计数
                    emptyParaCount++;
                    // 如果已经有空行段落，继续查找表格
                    if (emptyParaCount > 1) {
                        // 如果已经有多个空行段落，停止查找（表格应该更近）
                        break;
                    }
                }
            } else {
                // 遇到其他类型的元素，停止查找
                break;
            }
        }
        cursor.close();
        
        // 方法2：在body的表格数组中查找，检查是否在段落后（跳过空行段落）
        // 找到段落后的第一个表格，但只返回紧邻的表格（中间最多只有空行段落）
        CTTbl closestTable = null;
        org.apache.xmlbeans.XmlCursor paraEndCursor = paragraph.newCursor();
        paraEndCursor.toEndToken();
        
        for (int i = 0; i < body.sizeOfTblArray(); i++) {
            CTTbl tbl = body.getTblArray(i);
            org.apache.xmlbeans.XmlCursor tblCursor = tbl.newCursor();
            
            // 如果表格在段落后
            if (tblCursor.comparePosition(paraEndCursor) > 0) {
                // 检查中间是否有其他段落（非空行）
                boolean hasNonEmptyParaBetween = false;
                for (int j = pIndex + 1; j < body.sizeOfPArray(); j++) {
                    CTP checkPara = body.getPArray(j);
                    org.apache.xmlbeans.XmlCursor checkCursor = checkPara.newCursor();
                    
                    // 如果这个段落在表格之前
                    if (checkCursor.comparePosition(tblCursor) < 0) {
                        String checkText = "";
                        try {
                            XWPFParagraph xwpfCheckPara = new XWPFParagraph(checkPara, null);
                            checkText = xwpfCheckPara.getText();
                        } catch (Exception e) {
                            // 忽略异常
                        }
                        // 如果段落不为空，说明表格不是紧邻的
                        if (checkText != null && !checkText.trim().isEmpty()) {
                            hasNonEmptyParaBetween = true;
                            checkCursor.close();
                            break;
                        }
                        checkCursor.close();
                    } else {
                        checkCursor.close();
                        break;
                    }
                }
                
                if (!hasNonEmptyParaBetween) {
                    // 找到紧邻的表格，记录它
                    if (closestTable == null) {
                        closestTable = tbl;
                    } else {
                        // 如果有多个紧邻的表格，选择最近的（位置最小的）
                        org.apache.xmlbeans.XmlCursor closestCursor = closestTable.newCursor();
                        if (tblCursor.comparePosition(closestCursor) < 0) {
                            closestTable = tbl;
                        }
                        closestCursor.close();
                    }
                }
            }
            
            tblCursor.close();
        }
        
        paraEndCursor.close();
        
        if (closestTable != null) {
            return closestTable;
        }
        
        return null;
    }
    
    /**
     * 查找段落后的表格
     */
    private XWPFTable findTableAfterParagraph(CTBody body, CTP paragraph) {
        // 找到段落在body中的位置
        int pIndex = -1;
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i) == paragraph) {
                pIndex = i;
                break;
            }
        }
        
        if (pIndex == -1) {
            return null;
        }
        
        // 检查段落后是否有表格
        // 使用 XML cursor 检查下一个元素
        org.apache.xmlbeans.XmlCursor cursor = paragraph.newCursor();
        cursor.toEndToken();
        
        // 移动到下一个兄弟元素
        if (cursor.toNextSibling()) {
            org.apache.xmlbeans.XmlObject obj = cursor.getObject();
            cursor.close();
            
            // 检查是否是表格
            if (obj instanceof CTTbl) {
                CTTbl cttbl = (CTTbl) obj;
                // 在 body 的表格数组中查找这个表格
                for (int i = 0; i < body.sizeOfTblArray(); i++) {
                    if (body.getTblArray(i) == cttbl) {
                        // 找到了，需要创建 XWPFTable 对象
                        // 由于我们需要 document 引用，我们需要从 document 中查找
                        return null; // 返回 null，让调用者知道找到了表格但需要特殊处理
                    }
                }
            }
        } else {
            cursor.close();
        }
        
        return null;
    }
    
    /**
     * 找到章节的最后一个子章节，用于确定插入点
     * 样式ID对应：2=Heading1, 3=Heading2, 4=Heading3
     * X.X 章节用 Heading 2 (样式ID=3)，X.X.X 子章节用 Heading 3 (样式ID=4)
     */
    private XWPFParagraph findLastElementInSection(XWPFDocument document, XWPFParagraph sectionPara, String sectionNumber) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        // 目录样式 ID
        java.util.Set<String> tocStyles = new java.util.HashSet<>();
        tocStyles.add("22");
        tocStyles.add("25");
        tocStyles.add("16");
        
        // 找到章节段落在列表中的位置
        int sectionIndex = -1;
        for (int i = 0; i < paragraphs.size(); i++) {
            if (paragraphs.get(i).getCTP() == sectionPara.getCTP()) {
                sectionIndex = i;
                break;
            }
        }
        
        if (sectionIndex == -1) {
            return sectionPara;
        }
        
        // 从章节段落后开始查找，找到最后一个属于该章节的子章节
        XWPFParagraph lastSubSection = null;
        
        for (int i = sectionIndex + 1; i < paragraphs.size(); i++) {
            XWPFParagraph para = paragraphs.get(i);
            String text = para.getText();
            String styleName = para.getStyle();
            
            if (text == null || text.trim().isEmpty()) {
                continue;
            }
            
            // 跳过目录项
            boolean isToc = styleName != null && 
                (tocStyles.contains(styleName) || styleName.toLowerCase().startsWith("toc"));
            if (isToc) {
                continue;
            }
            
            // 通过样式识别章节边界
            if (styleName != null) {
                // 样式 "2" = Heading 1 (一级标题) -> 遇到一级标题，停止
                // 样式 "3" = Heading 2 (二级标题/主章节) -> 遇到下一个主章节，停止
                if (styleName.equals("2") || styleName.equals("3") || 
                    styleName.toLowerCase().contains("heading 1") ||
                    styleName.toLowerCase().contains("heading 2") ||
                    styleName.contains("标题 1") || styleName.contains("标题 2")) {
                    break;
                }
                // 样式 "4" = Heading 3 (三级标题/子章节) -> 记录为 lastSubSection
                // 也包括 "程序标题 3" 等变体
                if (styleName.equals("4") || styleName.equals("5") ||
                    styleName.toLowerCase().contains("heading 3") || 
                    styleName.toLowerCase().contains("heading 4") ||
                    styleName.contains("标题 3") || styleName.contains("程序标题")) {
                    lastSubSection = para;
                    continue;
                }
            }
        }
        
        // 如果找到了子章节，返回最后一个子章节；否则返回主章节
        return lastSubSection != null ? lastSubSection : sectionPara;
    }
    
    /**
     * 找到章节下所有已存在的子章节段落
     * 用于复用模板中的子章节（如 XX测试）
     */
    private List<XWPFParagraph> findExistingSubSectionsInSection(XWPFDocument document, 
                                                                   XWPFParagraph sectionPara, 
                                                                   String sectionNumber) {
        List<XWPFParagraph> subSections = new java.util.ArrayList<>();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        // 目录样式 ID
        java.util.Set<String> tocStyles = new java.util.HashSet<>();
        tocStyles.add("22");
        tocStyles.add("25");
        tocStyles.add("16");
        
        // 找到章节段落在列表中的位置
        int sectionIndex = -1;
        for (int i = 0; i < paragraphs.size(); i++) {
            if (paragraphs.get(i).getCTP() == sectionPara.getCTP()) {
                sectionIndex = i;
                break;
            }
        }
        
        if (sectionIndex == -1) {
            return subSections;
        }
        
        // 从章节段落后开始查找子章节
        for (int i = sectionIndex + 1; i < paragraphs.size(); i++) {
            XWPFParagraph para = paragraphs.get(i);
            String text = para.getText();
            String styleName = para.getStyle();
            
            if (text == null || text.trim().isEmpty()) {
                continue;
            }
            
            // 跳过目录项
            boolean isToc = styleName != null && 
                (tocStyles.contains(styleName) || styleName.toLowerCase().startsWith("toc"));
            if (isToc) {
                continue;
            }
            
            // 检查是否是主章节（Heading 2），如果是则停止
            if (styleName != null) {
                if (styleName.equals("2") || styleName.equals("3") || 
                    styleName.toLowerCase().contains("heading 1") ||
                    styleName.toLowerCase().contains("heading 2") ||
                    styleName.contains("标题 1") || styleName.contains("标题 2")) {
                    break;
                }
                // 子章节样式（Heading 3/4）
                if (styleName.equals("4") || styleName.equals("5") ||
                    styleName.toLowerCase().contains("heading 3") || 
                    styleName.toLowerCase().contains("heading 4") ||
                    styleName.contains("标题 3") || styleName.contains("程序标题")) {
                    subSections.add(para);
                }
            }
        }
        
        return subSections;
    }
    
    /**
     * 找到表格后的插入点（用于插入下一个子章节）
     */
    private XWPFParagraph findInsertPointAfterTable(XWPFDocument document, XWPFParagraph beforeTablePara) {
        CTBody body = document.getDocument().getBody();
        CTTbl table = findTableCttblAfterParagraph(body, beforeTablePara.getCTP());
        
        if (table == null) {
            return beforeTablePara;
        }
        
        // 在表格后创建空段落作为插入点
        org.apache.xmlbeans.XmlCursor tableCursor = table.newCursor();
        tableCursor.toEndToken();
        tableCursor.toNextToken();
        
        // 找到表格后的段落位置
        int insertPos = -1;
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            CTP checkPara = body.getPArray(i);
            org.apache.xmlbeans.XmlCursor checkCursor = checkPara.newCursor();
            if (checkCursor.comparePosition(tableCursor) > 0) {
                insertPos = i;
                checkCursor.close();
                break;
            }
            checkCursor.close();
        }
        tableCursor.close();
        
        if (insertPos >= 0) {
            // 检查该位置是否已有段落，如果有空段落则使用，否则创建新的
            CTP existingPara = body.getPArray(insertPos);
            XWPFParagraph existingXwpfPara = new XWPFParagraph(existingPara, document);
            String paraText = existingXwpfPara.getText();
            if (paraText == null || paraText.trim().isEmpty()) {
                // 使用现有的空段落
                return existingXwpfPara;
            } else {
                // 在当前位置插入新段落
                CTP newPara = body.insertNewP(insertPos);
                return new XWPFParagraph(newPara, document);
            }
        } else {
            // 如果找不到位置，在body末尾添加
            CTP newPara = body.addNewP();
            return new XWPFParagraph(newPara, document);
        }
    }
    
    /**
     * 填充表格数据（识别Word表格格式并填充）
     * 支持两种格式：
     * 1. 2列格式：左列是标签，右列是数据（从Word表格读取标签）
     * 2. 多列表格式：第1行是表头，后续行是数据（如果Word表格是空的或格式不匹配）
     */
    private void fillTableData(XWPFTable table, TestCase testCase) {
        int rowCount = table.getNumberOfRows();
        if (rowCount == 0) {
            System.err.println("警告：表格为空，无法填充数据");
            return;
        }
        
        // 检查表格格式：读取第一行，判断是2列格式还是多列格式
        XWPFTableRow firstRow = table.getRow(0);
        int firstRowCellCount = firstRow != null ? firstRow.getTableCells().size() : 0;
        
        // 如果第一行只有2列，且第二行也是2列，则认为是2列格式（标签-数据格式）
        boolean isTwoColumnFormat = false;
        if (firstRowCellCount == 2 && rowCount > 1) {
            XWPFTableRow secondRow = table.getRow(1);
            if (secondRow != null && secondRow.getTableCells().size() == 2) {
                // 读取第一行和第二行的左列内容，判断是否是标签格式
                String firstRowLeft = getCellText(firstRow.getCell(0));
                String secondRowLeft = getCellText(secondRow.getCell(0));
                
                // 如果左列包含常见的标签关键词，则认为是2列格式
                // 或者左列是Excel列名（如testName, id等），也认为是2列格式
                if (firstRowLeft.contains("测试项名称") || firstRowLeft.contains("标识") ||
                    secondRowLeft.contains("测试内容") || secondRowLeft.contains("测试项名称") ||
                    secondRowLeft.contains("测试策略") || secondRowLeft.contains("判定准则") ||
                    firstRowLeft.equals("testName") || firstRowLeft.equals("id") ||
                    secondRowLeft.equals("content") || secondRowLeft.equals("strategy") ||
                    secondRowLeft.equals("criteria") || secondRowLeft.equals("stopCondition") ||
                    secondRowLeft.equals("trace")) {
                    isTwoColumnFormat = true;
                }
            }
        }
        
        if (isTwoColumnFormat) {
            // 2列格式：读取Word表格左列的标签，从Excel中找到对应的数据填充到右列
            fillTwoColumnTable(table, testCase);
        } else {
            // 多列格式：使用Excel的列名作为表头
            fillMultiColumnTable(table, testCase);
        }
    }
    
    /**
     * 填充2列格式的表格（左列标签，右列数据）
     */
    private void fillTwoColumnTable(XWPFTable table, TestCase testCase) {
        // 获取Excel中所有可用的列名
        Map<String, String> columnData = testCase.getColumnData();
        
        int rowCount = table.getNumberOfRows();
        if (rowCount == 0) {
            return;
        }
        
        XWPFTableRow firstRow = table.getRow(0);
        int firstRowCellCount = firstRow != null ? firstRow.getTableCells().size() : 0;
        
        // 如果第一行是4列格式，特殊处理
        if (firstRowCellCount == 4) {
            // 填充第一行的4列数据
            String firstCol = getCellText(firstRow.getCell(0)).trim();
            String thirdCol = getCellText(firstRow.getCell(2)).trim();
            
            // 在Excel列名中查找匹配的列
            String firstColMatch = findMatchingColumn(firstCol, columnData.keySet());
            if (firstColMatch != null) {
                setCellValue(firstRow.getCell(1), testCase.getColumnValue(firstColMatch));
            }
            
            String thirdColMatch = findMatchingColumn(thirdCol, columnData.keySet());
            if (thirdColMatch != null) {
                setCellValue(firstRow.getCell(3), testCase.getColumnValue(thirdColMatch));
            }
            
            // 从第2行开始处理2列格式
            for (int i = 1; i < rowCount; i++) {
                XWPFTableRow row = table.getRow(i);
                if (row == null || row.getTableCells().size() < 2) {
                    continue;
                }
                
                // 读取左列的标签
                String label = getCellText(row.getCell(0)).trim();
                
                // 在Excel列名中查找匹配的列
                String matchedColumn = findMatchingColumn(label, columnData.keySet());
                
                // 如果找到了对应的列，填充数据到右列
                if (matchedColumn != null) {
                    String value = testCase.getColumnValue(matchedColumn);
                    // 注意：由于单元格合并，单元格0-1是一个物理单元格（标签列），单元格2-3是另一个物理单元格（数据列）
                    // 所以应该填充单元格2，而不是单元格1
                    if (row.getTableCells().size() >= 3) {
                        setCellValue(row.getCell(2), value);
                    } else {
                        setCellValue(row.getCell(1), value);
                    }
                }
            }
        } else {
            // 标准的2列格式，从第1行开始处理
            for (int i = 0; i < rowCount; i++) {
                XWPFTableRow row = table.getRow(i);
                if (row == null || row.getTableCells().size() < 2) {
                    continue;
                }
                
                // 读取左列的标签
                String label = getCellText(row.getCell(0)).trim();
                
                // 在Excel列名中查找匹配的列
                String matchedColumn = findMatchingColumn(label, columnData.keySet());
                
                // 如果找到了对应的列，填充数据到右列
                if (matchedColumn != null) {
                    String value = testCase.getColumnValue(matchedColumn);
                    // 注意：由于单元格合并，可能单元格0-1是一个物理单元格，单元格2-3是另一个
                    // 所以应该填充最后一个单元格或单元格2
                    if (row.getTableCells().size() >= 3) {
                        setCellValue(row.getCell(2), value);
                    } else {
                        setCellValue(row.getCell(1), value);
                    }
                }
            }
        }
    }
    
    /**
     * 清空单元格内容
     */
    private void clearCellContent(XWPFTableCell cell) {
        if (cell == null) {
            return;
        }
        
        // 清空所有段落
        while (cell.getParagraphs().size() > 0) {
            cell.removeParagraph(0);
        }
        
        // 创建一个空段落
        cell.addParagraph();
    }
    
    /**
     * 在Excel列名集合中查找匹配的列名
     * 支持完全匹配和模糊匹配
     */
    private String findMatchingColumn(String label, java.util.Set<String> columnNames) {
        if (label == null || label.isEmpty()) {
            return null;
        }
        
        // 1. 尝试完全匹配
        for (String columnName : columnNames) {
            if (columnName.equals(label)) {
                return columnName;
            }
        }
        
        // 2. 尝试包含匹配（Word标签包含Excel列名，或Excel列名包含Word标签）
        for (String columnName : columnNames) {
            if (label.contains(columnName) || columnName.contains(label)) {
                return columnName;
            }
        }
        
        // 3. 尝试去除空格后匹配
        String labelNoSpace = label.replace(" ", "").replace("　", "");
        for (String columnName : columnNames) {
            String columnNameNoSpace = columnName.replace(" ", "").replace("　", "");
            if (labelNoSpace.equals(columnNameNoSpace)) {
                return columnName;
            }
        }
        
        return null;
    }
    
    /**
     * 填充多列格式的表格（第1行是表头，后续行是数据）
     */
    private void fillMultiColumnTable(XWPFTable table, TestCase testCase) {
        Map<String, String> columnData = testCase.getColumnData();
        
        if (columnData.isEmpty()) {
            System.err.println("警告：测试用例没有列数据");
            return;
        }
        
        int currentRowCount = table.getNumberOfRows();
        if (currentRowCount == 0) {
            System.err.println("警告：表格没有行，无法填充数据");
            return;
        }
        
        // 不创建新行，只填充已有的行
        // 假设表格已经有了正确的结构
        
        // 填充表头行（第1行）
        XWPFTableRow headerRow = table.getRow(0);
        if (headerRow != null && headerRow.getTableCells().size() >= 2) {
            // 第1行：测试项名称 | testName值 | 标识 | id值
            if (headerRow.getTableCells().size() >= 4) {
                // 读取第1列和第3列的标签，在Excel中查找匹配的列
                String firstCol = getCellText(headerRow.getCell(0)).trim();
                String thirdCol = getCellText(headerRow.getCell(2)).trim();
                
                String firstColMatch = findMatchingColumn(firstCol, columnData.keySet());
                if (firstColMatch != null) {
                    setCellValue(headerRow.getCell(1), testCase.getColumnValue(firstColMatch));
                }
                
                String thirdColMatch = findMatchingColumn(thirdCol, columnData.keySet());
                if (thirdColMatch != null) {
                    setCellValue(headerRow.getCell(3), testCase.getColumnValue(thirdColMatch));
                }
            }
        }
        
        // 填充数据行（第2行及以后）
        for (int i = 1; i < currentRowCount; i++) {
            XWPFTableRow row = table.getRow(i);
            if (row == null || row.getTableCells().size() < 2) {
                continue;
            }
            
            // 读取左列的标签
            String label = getCellText(row.getCell(0)).trim();
            
            // 在Excel列名中查找匹配的列
            String matchedColumn = findMatchingColumn(label, columnData.keySet());
            if (matchedColumn != null) {
                String value = testCase.getColumnValue(matchedColumn);
                // 填充到右列（考虑单元格合并的情况）
                if (row.getTableCells().size() >= 3) {
                    setCellValue(row.getCell(2), value);
                } else {
                    setCellValue(row.getCell(1), value);
                }
            }
        }
    }
    
    /**
     * 获取单元格文本内容
     */
    private String getCellText(XWPFTableCell cell) {
        if (cell == null) {
            return "";
        }
        StringBuilder text = new StringBuilder();
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        for (XWPFParagraph para : paragraphs) {
            text.append(para.getText());
        }
        return text.toString().trim();
    }
    
    /**
     * 设置单元格值（直接操作单元格）
     */
    private void setCellValue(XWPFTableCell cell, String value) {
        if (cell == null) {
            return;
        }
        
        // 清空现有内容
        cell.removeParagraph(0);
        while (cell.getParagraphs().size() > 0) {
            cell.removeParagraph(0);
        }
        
        // 创建新段落并设置文本
        XWPFParagraph para = cell.addParagraph();
        XWPFRun run = para.createRun();
        run.setText(value != null ? value : "");
    }
    
    /**
     * 设置列宽
     */
    private void setColumnWidth(XWPFTable table, int row, int col, int width) {
        XWPFTableRow tableRow = table.getRow(row);
        if (tableRow == null) {
            return;
        }
        
        XWPFTableCell cell = tableRow.getCell(col);
        if (cell == null) {
            return;
        }
        
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr == null) {
            tcPr = cell.getCTTc().addNewTcPr();
        }
        
        CTTblWidth cellWidth = tcPr.getTcW();
        if (cellWidth == null) {
            cellWidth = tcPr.addNewTcW();
        }
        cellWidth.setType(STTblWidth.DXA);
        cellWidth.setW(BigInteger.valueOf(width));
    }
    
    /**
     * 横向合并单元格 - 使用 hMerge 标准方式并完全隐藏被合并单元格
     */
    private void mergeCellsHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        XWPFTableRow tableRow = table.getRow(row);
        if (tableRow == null) {
            return;
        }
        
        for (int col = fromCol; col <= toCol; col++) {
            XWPFTableCell cell = tableRow.getCell(col);
            if (cell == null) {
                continue;
            }
            
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr == null) {
                tcPr = cell.getCTTc().addNewTcPr();
            }
            
            if (col == fromCol) {
                // 第一个单元格：设置为合并的起始单元格
                if (!tcPr.isSetHMerge()) {
                    tcPr.addNewHMerge().setVal(STMerge.RESTART);
                } else {
                    tcPr.getHMerge().setVal(STMerge.RESTART);
                }
            } else {
                // 后续单元格：设置为继续合并
                if (!tcPr.isSetHMerge()) {
                    tcPr.addNewHMerge().setVal(STMerge.CONTINUE);
                } else {
                    tcPr.getHMerge().setVal(STMerge.CONTINUE);
                }
                
                // 设置单元格宽度为0
                CTTblWidth cellWidth = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
                cellWidth.setType(STTblWidth.DXA);
                cellWidth.setW(BigInteger.ZERO);
                
                // 完全隐藏边框
                CTTcBorders borders = tcPr.isSetTcBorders() ? tcPr.getTcBorders() : tcPr.addNewTcBorders();
                
                // 创建无边框 - 使用 NONE 而不是 NIL
                CTBorder noBorder = CTBorder.Factory.newInstance();
                noBorder.setVal(STBorder.NONE);
                noBorder.setSz(BigInteger.ZERO);
                noBorder.setSpace(BigInteger.ZERO);
                noBorder.setColor("FFFFFF"); // 白色
                
                borders.setTop(noBorder);
                borders.setBottom(noBorder);
                borders.setLeft(noBorder);
                borders.setRight(noBorder);
                borders.setInsideH(noBorder);
                borders.setInsideV(noBorder);
                
                // 设置单元格底纹为白色（完全隐藏）
                CTShd shd = tcPr.isSetShd() ? tcPr.getShd() : tcPr.addNewShd();
                shd.setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd.CLEAR);
                shd.setColor("auto");
                shd.setFill("FFFFFF");
            }
        }
    }


    /**
     * 设置单元格值
     */
    private void setCellValue(XWPFTable table, int row, int col, String value) {
        XWPFTableRow tableRow = table.getRow(row);
        if (tableRow == null) {
            return;
        }
        
        XWPFTableCell cell = tableRow.getCell(col);
        if (cell == null) {
            return;
        }
        
        // 清除原有内容
        try {
            // 清除所有段落
            while (cell.getParagraphs().size() > 0) {
                cell.removeParagraph(0);
            }
        } catch (Exception e) {
            // 忽略错误
        }
        
        // 只有当值不为空时才添加新段落
        if (value != null && !value.isEmpty()) {
            XWPFParagraph para = cell.addParagraph();
            XWPFRun run = para.createRun();
            run.setText(value);
            
            // 设置单元格对齐方式：垂直居中、左对齐
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            para.setAlignment(ParagraphAlignment.LEFT);
            
            // 设置字体大小：五号
            run.setFontSize(10);
        }
    }

    /**
     * 设置表格样式：边框、对齐等
     */
    private void styleTable(CTTbl cttbl) {
        CTTblPr tblPr = cttbl.getTblPr();
        if (tblPr == null) {
            tblPr = cttbl.addNewTblPr();
        }

        // 设置表格边框
        CTTblBorders borders = tblPr.isSetTblBorders() ? 
                tblPr.getTblBorders() : tblPr.addNewTblBorders();
        
        // 设置所有边框为黑色细线（2pt）
        CTBorder border = CTBorder.Factory.newInstance();
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(16)); // 2pt = 16/8
        border.setColor("000000");

        borders.setTop(border);
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setInsideH(border);
        borders.setInsideV(border);

        // 设置单元格边框
        for (CTRow row : cttbl.getTrArray()) {
            for (CTTc cell : row.getTcArray()) {
                CTTcPr tcPr = cell.isSetTcPr() ? cell.getTcPr() : cell.addNewTcPr();
                
                CTTcBorders tcBorders = tcPr.isSetTcBorders() ? 
                        tcPr.getTcBorders() : tcPr.addNewTcBorders();
                
                tcBorders.setTop(border);
                tcBorders.setBottom(border);
                tcBorders.setLeft(border);
                tcBorders.setRight(border);
            }
        }
    }
}

