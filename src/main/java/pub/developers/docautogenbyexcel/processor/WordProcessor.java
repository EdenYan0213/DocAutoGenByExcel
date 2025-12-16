package pub.developers.docautogenbyexcel.processor;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.util.SectionNumberUtil;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.*;

/**
 * Word模板处理模块（精简版）
 * 支持任意级别的章节编号（如 5.3、5.3.1、5.3.1.1 等）
 */
public class WordProcessor {
    
    // 格式缓存
    private TextFormat subSectionFormat;
    private TextFormat captionFormat;
    private XWPFTable templateTable;

    /**
     * 处理Word文档
     */
    public int processWord(String templatePath, String outputPath, 
                          Map<String, ModuleData> moduleDataMap) throws Exception {
        validateFormat(templatePath);
        
        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // 扫描文档结构
            Map<String, SectionInfo> sections = scanSections(document);
            List<PlaceholderInfo> placeholders = scanPlaceholders(document);
            
            System.out.println("找到章节: " + sections.keySet());
            if (!placeholders.isEmpty()) {
                System.out.println("找到占位符: " + placeholders);
            }

            int successCount = 0;
            
            // 1. 处理已存在的章节
            for (Map.Entry<String, SectionInfo> entry : sections.entrySet()) {
                String sectionNumber = entry.getKey();
                SectionInfo sectionInfo = entry.getValue();
                ModuleData moduleData = moduleDataMap.get(sectionNumber);
                
                if (moduleData == null) continue;
                
                // 提取格式（只提取一次）
                extractFormatsIfNeeded(document, sectionInfo);
                
                // 处理该章节
                processSection(document, sectionInfo, moduleData);
                successCount++;
            }
            
            // 2. 处理占位符
            for (PlaceholderInfo placeholder : placeholders) {
                successCount += processPlaceholder(document, placeholder, moduleDataMap, sections);
            }

            // 保存文档
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }

            return successCount;

        } catch (OLE2NotOfficeXmlFileException e) {
            throw new Exception("不支持旧版Word格式(.doc)，请转换为.docx格式", e);
        } catch (IOException e) {
            throw new Exception("处理Word文档失败: " + e.getMessage(), e);
        }
    }

    // ==================== 扫描方法 ====================
    
    /**
     * 扫描文档中的所有章节（支持任意级别）
     * 支持两种模式：
     * 1. 从目录（TOC）读取章节
     * 2. 直接从正文中查找章节编号（不依赖样式）
     */
    private Map<String, SectionInfo> scanSections(XWPFDocument document) {
        Map<String, SectionInfo> sections = new LinkedHashMap<>();
        Map<String, String> tocSections = new LinkedHashMap<>();
        
        // 第一遍：从目录读取章节编号和名称
        for (XWPFParagraph para : document.getParagraphs()) {
            String text = para.getText();
            if (text == null || text.trim().isEmpty()) continue;
            
            if (SectionNumberUtil.isTocStyle(para.getStyle())) {
                String number = SectionNumberUtil.extractSectionNumber(text.trim());
                String name = SectionNumberUtil.extractSectionName(text.trim());
                if (number != null && name != null) {
                    tocSections.put(number, name);
                }
            }
        }
        
        // 如果有目录，从目录中查找对应段落
        if (!tocSections.isEmpty()) {
            for (Map.Entry<String, String> entry : tocSections.entrySet()) {
                String number = entry.getKey();
                String name = entry.getValue();
                
                for (XWPFParagraph para : document.getParagraphs()) {
                    String text = para.getText();
                    if (text == null || text.trim().isEmpty()) continue;
                    
                    String style = para.getStyle();
                    if (SectionNumberUtil.isTocStyle(style)) continue;
                    
                    String paraNumber = SectionNumberUtil.extractSectionNumber(text.trim());
                    if ((paraNumber != null && paraNumber.equals(number)) ||
                        (SectionNumberUtil.isHeadingStyle(style) && text.trim().equals(name))) {
                        sections.put(number, new SectionInfo(number, name, para));
                        break;
                    }
                }
            }
        } else {
            // 没有目录，直接从正文中查找章节（不依赖样式）
            for (XWPFParagraph para : document.getParagraphs()) {
                String text = para.getText();
                if (text == null || text.trim().isEmpty()) continue;
                
                String style = para.getStyle();
                if (SectionNumberUtil.isTocStyle(style)) continue;
                
                String number = SectionNumberUtil.extractSectionNumber(text.trim());
                String name = SectionNumberUtil.extractSectionName(text.trim());
                
                if (number != null && name != null) {
                    sections.put(number, new SectionInfo(number, name, para));
                }
            }
        }
        
        return sections;
    }
    
    /**
     * 扫描占位符
     */
    private List<PlaceholderInfo> scanPlaceholders(XWPFDocument document) {
        List<PlaceholderInfo> placeholders = new ArrayList<>();
        
        for (XWPFParagraph para : document.getParagraphs()) {
            String text = para.getText();
            if (text == null || text.trim().isEmpty()) continue;
            
            if (SectionNumberUtil.isPlaceholder(text.trim())) {
                String parent = SectionNumberUtil.extractPlaceholderParent(text.trim());
                if (parent != null) {
                    placeholders.add(new PlaceholderInfo(parent, text.trim(), para));
                }
            }
        }
        
        return placeholders;
    }

    // ==================== 处理方法 ====================
    
    /**
     * 处理单个章节
     * 确保每个子章节后紧跟其Caption和表格
     */
    private void processSection(XWPFDocument document, SectionInfo section, ModuleData moduleData) {
        CTBody body = document.getDocument().getBody();
        List<TestCase> testCases = moduleData.getTestCases();
        
        if (testCases.isEmpty()) return;
        
        // 查找该章节下的已有子章节
        List<XWPFParagraph> existingSubSections = findSubSections(document, section);
        
        // 找到下一个主章节的边界
        int boundary = findNextSectionBoundary(document, section);
        
        // 处理已存在的子章节（更新标题和填充表格）
        for (int i = 0; i < Math.min(testCases.size(), existingSubSections.size()); i++) {
            TestCase testCase = testCases.get(i);
            String subNumber = SectionNumberUtil.generateChildNumber(section.number, i + 1);
            XWPFParagraph existingPara = existingSubSections.get(i);
            
            // 更新子章节标题
            updateParagraphText(existingPara, subNumber + " " + testCase.getTestName() + "测试");
            
            // 更新Caption和填充表格
            updateCaption(document, existingPara, subNumber, testCase.getTestName());
            CTTbl table = findTableAfterParagraph(body, existingPara.getCTP());
            if (table != null) {
                fillTableData(new XWPFTable(table, document), testCase);
            }
        }
        
        // 如果Excel数据比已有子章节多，需要创建新的子章节
        if (testCases.size() > existingSubSections.size()) {
            // 找到最后一个已存在子章节的表格后面作为插入点
            XWPFParagraph insertPoint;
            if (!existingSubSections.isEmpty()) {
                insertPoint = findInsertPointAfterTable(document, existingSubSections.get(existingSubSections.size() - 1));
            } else {
                insertPoint = section.paragraph;
            }
            
            // 创建新的子章节，每个子章节后紧跟Caption和表格
            for (int i = existingSubSections.size(); i < testCases.size(); i++) {
                TestCase testCase = testCases.get(i);
                String subNumber = SectionNumberUtil.generateChildNumber(section.number, i + 1);
                
                // 1. 创建子章节标题
                XWPFParagraph subPara = createSubSection(document, insertPoint, subNumber, 
                                                          testCase.getTestName(), boundary);
                
                // 2. 紧接着创建Caption
                XWPFParagraph captionPara = createCaption(document, subPara, subNumber, testCase.getTestName());
                
                // 3. 紧接着创建表格
                if (templateTable != null) {
                    CTTbl newTable = copyTable(document, captionPara.getCTP(), templateTable.getCTTbl());
                    if (newTable != null) {
                        clearTableData(new XWPFTable(newTable, document));
                        fillTableData(new XWPFTable(newTable, document), testCase);
                    }
                } else {
                    insertNewTable(document, captionPara.getCTP(), testCase);
                }
                
                // 4. 更新插入点到刚创建的表格后面，为下一个子章节做准备
                insertPoint = findInsertPointAfterTable(document, captionPara);
                if (boundary > 0) boundary += 4;
            }
        }
        
    }
    
    /**
     * 处理占位符
     */
    private int processPlaceholder(XWPFDocument document, PlaceholderInfo placeholder,
                                    Map<String, ModuleData> moduleDataMap,
                                    Map<String, SectionInfo> existingSections) {
        // 查找匹配的子模块
        List<String> matchedModules = new ArrayList<>();
        for (String moduleNumber : moduleDataMap.keySet()) {
            if (SectionNumberUtil.isChildOf(moduleNumber, placeholder.parentNumber) &&
                !existingSections.containsKey(moduleNumber)) {
                matchedModules.add(moduleNumber);
            }
        }
        
        matchedModules.sort(SectionNumberUtil::compare);
        
        if (matchedModules.isEmpty()) {
            return 0;
        }
        
        XWPFParagraph insertPoint = placeholder.paragraph;
        
        for (String moduleNumber : matchedModules) {
            ModuleData moduleData = moduleDataMap.get(moduleNumber);
            TestCase testCase = moduleData.getTestCases().get(0);
            
            // 创建子章节
            XWPFParagraph subPara = createSubSection(document, insertPoint, moduleNumber, 
                                                      testCase.getTestName(), -1);
            
            // 插入内容
            insertModuleContent(document, subPara, moduleNumber, moduleData);
            
            insertPoint = findInsertPointAfterTable(document, subPara);
        }
        
        // 删除占位符
        removeParagraph(document, placeholder.paragraph);
        
        return matchedModules.size();
    }
    
    /**
     * 插入模块内容
     */
    private void insertModuleContent(XWPFDocument document, XWPFParagraph sectionPara,
                                     String moduleNumber, ModuleData moduleData) {
        List<TestCase> testCases = moduleData.getTestCases();
        if (testCases.isEmpty()) return;

        XWPFParagraph insertPoint = sectionPara;
        
        for (int i = 0; i < testCases.size(); i++) {
            TestCase testCase = testCases.get(i);
            String subNumber = SectionNumberUtil.generateChildNumber(moduleNumber, i + 1);
            
            XWPFParagraph subPara = createSubSection(document, insertPoint, subNumber, 
                                                      testCase.getTestName(), -1);
            
            // 创建表格
            if (templateTable != null) {
                XWPFParagraph captionPara = createCaption(document, subPara, subNumber, testCase.getTestName());
                CTTbl newTable = copyTable(document, captionPara.getCTP(), templateTable.getCTTbl());
                if (newTable != null) {
                    clearTableData(new XWPFTable(newTable, document));
                    fillTableData(new XWPFTable(newTable, document), testCase);
                }
            } else {
                insertNewTable(document, subPara.getCTP(), testCase);
            }
            
            insertPoint = findInsertPointAfterTable(document, subPara);
        }
    }

    // ==================== 辅助方法 ====================
    
    /**
     * 提取格式（只提取一次）
     */
    private void extractFormatsIfNeeded(XWPFDocument document, SectionInfo section) {
        if (subSectionFormat != null) return;
        
        // 方法1：从已存在的子章节提取
        List<XWPFParagraph> subSections = findSubSections(document, section);
        if (!subSections.isEmpty()) {
            XWPFParagraph firstSub = subSections.get(0);
            subSectionFormat = extractTextFormat(firstSub);
            
            // 查找Caption
            XWPFParagraph caption = findCaptionAfterParagraph(document, firstSub);
            if (caption != null) {
                captionFormat = extractTextFormat(caption);
            }
            
            // 查找模板表格（子章节后面的表格）
            CTTbl table = findTableAfterParagraph(document.getDocument().getBody(), firstSub.getCTP());
            if (table != null && isTestCaseTable(new XWPFTable(table, document))) {
                templateTable = new XWPFTable(table, document);
            }
        }
        
        // 方法2：如果没有子章节，直接从章节后面找表格
        if (templateTable == null) {
            CTTbl table = findTableAfterParagraph(document.getDocument().getBody(), section.paragraph.getCTP());
            if (table != null && isTestCaseTable(new XWPFTable(table, document))) {
                templateTable = new XWPFTable(table, document);
            }
        }
        
        // 方法3：遍历文档找最佳测试用例表格作为模板
        if (templateTable == null) {
            templateTable = findBestTemplateTable(document);
        }
        
        // 设置默认格式
        if (subSectionFormat == null) {
            subSectionFormat = new TextFormat("黑体", 12, false, "4");
        }
        if (captionFormat == null) {
            captionFormat = new TextFormat("黑体", 12, false, "11");
        }
    }
    
    /**
     * 判断是否是测试用例表格
     * 测试用例表格特征：
     * 1. 至少有5行（优先选择6行的表格）
     * 2. 第一行第一个单元格包含"测试项名称"
     * 3. 至少有2列
     * 4. 包含"测试内容"或"测试策略"等关键字
     */
    private boolean isTestCaseTable(XWPFTable table) {
        if (table == null || table.getNumberOfRows() < 5) return false;
        
        XWPFTableRow firstRow = table.getRow(0);
        if (firstRow == null || firstRow.getTableCells().size() < 2) return false;
        
        String firstCellText = getCellText(firstRow.getCell(0));
        if (!firstCellText.contains("测试项名称") && !firstCellText.contains("测试项")) {
            return false;
        }
        
        // 检查是否有"测试内容"行（优先选择包含测试内容的表格）
        for (int i = 1; i < table.getNumberOfRows(); i++) {
            XWPFTableRow row = table.getRow(i);
            if (row != null && row.getTableCells().size() > 0) {
                String cellText = getCellText(row.getCell(0));
                if (cellText.contains("测试内容")) {
                    return true;  // 优先返回包含"测试内容"的表格
                }
            }
        }
        
        return true;
    }
    
    /**
     * 在文档中查找最佳的测试用例表格模板
     * 优先选择6行且包含"测试内容"的表格
     */
    private XWPFTable findBestTemplateTable(XWPFDocument document) {
        XWPFTable bestTable = null;
        int bestScore = 0;
        
        for (XWPFTable table : document.getTables()) {
            if (!isTestCaseTable(table)) continue;
            
            int score = 0;
            int rows = table.getNumberOfRows();
            
            // 优先选择6行的表格
            if (rows == 6) score += 10;
            else if (rows >= 5) score += 5;
            
            // 检查是否包含关键字段
            for (int i = 0; i < rows; i++) {
                XWPFTableRow row = table.getRow(i);
                if (row == null || row.getTableCells().isEmpty()) continue;
                
                String cellText = getCellText(row.getCell(0));
                if (cellText.contains("测试内容")) score += 5;
                if (cellText.contains("测试策略")) score += 3;
                if (cellText.contains("判定准则")) score += 3;
                if (cellText.contains("追踪关系")) score += 2;
            }
            
            if (score > bestScore) {
                bestScore = score;
                bestTable = table;
            }
        }
        
        return bestTable;
    }
    
    /**
     * 查找章节下的所有子章节
     */
    private List<XWPFParagraph> findSubSections(XWPFDocument document, SectionInfo section) {
        List<XWPFParagraph> subSections = new ArrayList<>();
        boolean foundSection = false;
        
        for (XWPFParagraph para : document.getParagraphs()) {
            if (para.getCTP() == section.paragraph.getCTP()) {
                foundSection = true;
                continue;
            }
            
            if (!foundSection) continue;
            
            String text = para.getText();
            if (text == null || text.trim().isEmpty()) continue;
            
            String style = para.getStyle();
            if (SectionNumberUtil.isTocStyle(style)) continue;
            
            String paraNumber = SectionNumberUtil.extractSectionNumber(text.trim());
            
            // 遇到同级或更高级章节，停止
            if (paraNumber != null && !SectionNumberUtil.isDescendantOf(paraNumber, section.number)) {
                if (SectionNumberUtil.isHeadingStyle(style)) break;
            }
            
            // 找到子章节
            if (paraNumber != null && SectionNumberUtil.isChildOf(paraNumber, section.number)) {
                subSections.add(para);
            }
        }
        
        return subSections;
    }
    
    /**
     * 查找章节下最后一个子章节
     */
    private XWPFParagraph findLastSubSection(XWPFDocument document, SectionInfo section) {
        List<XWPFParagraph> subSections = findSubSections(document, section);
        return subSections.isEmpty() ? null : subSections.get(subSections.size() - 1);
    }
    
    /**
     * 查找下一个主章节的边界位置
     */
    private int findNextSectionBoundary(XWPFDocument document, SectionInfo section) {
        CTBody body = document.getDocument().getBody();
        boolean foundSection = false;
        
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            CTP ctp = body.getPArray(i);
            XWPFParagraph para = new XWPFParagraph(ctp, document);
            
            if (para.getCTP() == section.paragraph.getCTP()) {
                foundSection = true;
                continue;
            }
            
            if (!foundSection) continue;
            
            String text = para.getText();
            if (text == null || text.trim().isEmpty()) continue;
            
            String style = para.getStyle();
            if (SectionNumberUtil.isTocStyle(style)) continue;
            
            String paraNumber = SectionNumberUtil.extractSectionNumber(text.trim());
            
            // 找到同级或更高级的章节
            if (paraNumber != null && SectionNumberUtil.isHeadingStyle(style)) {
                int paraLevel = SectionNumberUtil.getLevel(paraNumber);
                int sectionLevel = SectionNumberUtil.getLevel(section.number);
                if (paraLevel <= sectionLevel) {
                    return i;
                }
            }
        }
        
        return -1;
    }
    
    /**
     * 查找段落后的表格
     */
    private CTTbl findTableAfterParagraph(CTBody body, CTP paragraph) {
        org.apache.xmlbeans.XmlCursor cursor = paragraph.newCursor();
        cursor.toEndToken();
        
        int emptyParaCount = 0;
        while (cursor.toNextSibling()) {
            org.apache.xmlbeans.XmlObject obj = cursor.getObject();
            
            if (obj instanceof CTTbl) {
                cursor.close();
                return (CTTbl) obj;
            }
            
            if (obj instanceof CTP nextPara) {
                String text = "";
                try {
                    text = new XWPFParagraph(nextPara, null).getText();
                } catch (Exception ignored) {}
                
                if (text != null && !text.trim().isEmpty()) break;
                if (++emptyParaCount > 2) break;
            } else {
                break;
            }
        }
        cursor.close();
        
        return null;
    }
    
    /**
     * 查找段落后的Caption
     */
    private XWPFParagraph findCaptionAfterParagraph(XWPFDocument document, XWPFParagraph para) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        int index = paragraphs.indexOf(para);
        if (index == -1) return null;
        
        for (int i = index + 1; i < paragraphs.size() && i < index + 5; i++) {
            XWPFParagraph p = paragraphs.get(i);
            if (SectionNumberUtil.isCaptionStyle(p.getStyle()) || 
                (p.getText() != null && p.getText().trim().startsWith("表"))) {
                return p;
            }
        }
        return null;
    }
    
    /**
     * 查找表格后的插入点
     */
    private XWPFParagraph findInsertPointAfterTable(XWPFDocument document, XWPFParagraph beforePara) {
        CTBody body = document.getDocument().getBody();
        CTTbl table = findTableAfterParagraph(body, beforePara.getCTP());
        
        if (table == null) return beforePara;
        
        org.apache.xmlbeans.XmlCursor cursor = table.newCursor();
        cursor.toEndToken();
        cursor.toNextToken();
        
        // 找到表格后的段落
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            CTP checkPara = body.getPArray(i);
            org.apache.xmlbeans.XmlCursor checkCursor = checkPara.newCursor();
            if (checkCursor.comparePosition(cursor) > 0) {
                checkCursor.close();
                cursor.close();
                
                XWPFParagraph p = new XWPFParagraph(checkPara, document);
                String text = p.getText();
                if (text == null || text.trim().isEmpty()) {
                    return p;
                } else {
                    CTP newPara = body.insertNewP(i);
                    return new XWPFParagraph(newPara, document);
                }
            }
            checkCursor.close();
        }
        cursor.close();
        
        CTP newPara = body.addNewP();
        return new XWPFParagraph(newPara, document);
    }

    // ==================== 创建方法 ====================
    
    /**
     * 创建子章节段落
     */
    private XWPFParagraph createSubSection(XWPFDocument document, XWPFParagraph afterPara,
                                            String number, String name, int boundary) {
        CTBody body = document.getDocument().getBody();
        int afterIndex = findParagraphIndex(body, afterPara.getCTP());
        
        if (afterIndex == -1) afterIndex = body.sizeOfPArray() - 1;
        
        int insertIndex = afterIndex + 1;
        if (boundary > 0 && insertIndex >= boundary) {
            insertIndex = boundary;
        }
        
        CTP ctp = body.insertNewP(insertIndex);
        XWPFParagraph para = new XWPFParagraph(ctp, document);
        
        // 设置样式
        int level = SectionNumberUtil.getLevel(number);
        String styleId = subSectionFormat != null ? subSectionFormat.styleId : SectionNumberUtil.getHeadingStyleId(level);
        try {
            para.setStyle(styleId);
        } catch (Exception ignored) {}
        
        // 禁用自动编号
        disableNumbering(ctp);
        
        // 设置内容
        String title = number + " " + name + "测试";
        TextFormat fmt = subSectionFormat != null ? subSectionFormat : new TextFormat("黑体", 12, false, styleId);
        
        XWPFRun run = para.createRun();
        run.setText(title);
        run.setFontFamily(fmt.fontFamily);
        run.setFontSize(fmt.fontSize);
        if (fmt.bold != null) run.setBold(fmt.bold);
        
        para.setAlignment(ParagraphAlignment.LEFT);
        
        System.out.println("创建子章节: " + title);
        return para;
    }
    
    /**
     * 创建Caption
     */
    private XWPFParagraph createCaption(XWPFDocument document, XWPFParagraph afterPara,
                                         String number, String name) {
        CTBody body = document.getDocument().getBody();
        int afterIndex = findParagraphIndex(body, afterPara.getCTP());
        
        if (afterIndex == -1) afterIndex = body.sizeOfPArray() - 1;
        
        CTP ctp = body.insertNewP(afterIndex + 1);
        XWPFParagraph para = new XWPFParagraph(ctp, document);
        
        // 设置样式
        String styleId = captionFormat != null ? captionFormat.styleId : "11";
        try {
            para.setStyle(styleId);
        } catch (Exception ignored) {}
        
        // 设置内容
        String text = "表" + number + " " + name + "测试";
        TextFormat fmt = captionFormat != null ? captionFormat : new TextFormat("黑体", 12, false, styleId);
        
        XWPFRun run = para.createRun();
        run.setText(text);
        run.setFontFamily(fmt.fontFamily);
        run.setFontSize(fmt.fontSize);
        if (fmt.bold != null) run.setBold(fmt.bold);
        
        para.setAlignment(ParagraphAlignment.CENTER);
        
        return para;
    }
    
    /**
     * 更新Caption
     */
    private void updateCaption(XWPFDocument document, XWPFParagraph subPara, String number, String name) {
        XWPFParagraph caption = findCaptionAfterParagraph(document, subPara);
        if (caption != null) {
            updateParagraphText(caption, "表" + number + " " + name + "测试");
        }
    }

    // ==================== 表格方法 ====================
    
    /**
     * 插入新表格
     */
    private void insertNewTable(XWPFDocument document, CTP paragraph, TestCase testCase) {
        org.apache.xmlbeans.XmlCursor cursor = paragraph.newCursor();
        cursor.toEndToken();
        cursor.toNextToken();
        
        cursor.beginElement(
            new javax.xml.namespace.QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "tbl", "w"));
        cursor.toParent();
        
        CTTbl cttbl = null;
        if (cursor.getObject() instanceof CTTbl) {
            cttbl = (CTTbl) cursor.getObject();
        }
        cursor.close();
        
        if (cttbl == null) return;
        
        XWPFTable table = new XWPFTable(cttbl, document);
        
        // 创建表格结构
        createTableStructure(table, testCase);
        fillTableData(table, testCase);
        styleTable(cttbl);
    }
    
    /**
     * 复制表格
     */
    private CTTbl copyTable(XWPFDocument document, CTP afterParagraph, CTTbl source) {
        try {
            org.apache.xmlbeans.XmlCursor cursor = afterParagraph.newCursor();
            cursor.toEndToken();
            cursor.toNextToken();
            
            cursor.beginElement(
                new javax.xml.namespace.QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "tbl", "w"));
            cursor.toParent();
            
            CTTbl newTable = null;
            if (cursor.getObject() instanceof CTTbl) {
                newTable = (CTTbl) cursor.getObject();
            }
            cursor.close();
            
            if (newTable == null) return null;
            
            // 复制行
            for (int i = 0; i < source.sizeOfTrArray(); i++) {
                CTRow sourceRow = source.getTrArray(i);
                CTRow newRow = newTable.addNewTr();
                
                if (sourceRow.isSetTrPr()) {
                    newRow.setTrPr((CTTrPr) sourceRow.getTrPr().copy());
                }
                
                for (int j = 0; j < sourceRow.sizeOfTcArray(); j++) {
                    CTTc sourceCell = sourceRow.getTcArray(j);
                    CTTc newCell = newRow.addNewTc();
                    
                    if (sourceCell.getTcPr() != null) {
                        newCell.setTcPr((CTTcPr) sourceCell.getTcPr().copy());
                    }
                    
                    // 只复制标签列
                    boolean copyContent = (i == 0 && (j == 0 || j == 2)) || (i > 0 && j == 0);
                    
                    if (copyContent) {
                        for (int k = 0; k < sourceCell.sizeOfPArray(); k++) {
                            CTP sourceP = sourceCell.getPArray(k);
                            CTP newP = newCell.addNewP();
                            if (sourceP.getPPr() != null) {
                                newP.setPPr((CTPPr) sourceP.getPPr().copy());
                            }
                            for (int m = 0; m < sourceP.sizeOfRArray(); m++) {
                                CTR sourceR = sourceP.getRArray(m);
                                CTR newR = newP.addNewR();
                                if (sourceR.getRPr() != null) {
                                    newR.setRPr((CTRPr) sourceR.getRPr().copy());
                                }
                                for (int n = 0; n < sourceR.sizeOfTArray(); n++) {
                                    CTText sourceT = sourceR.getTArray(n);
                                    CTText newT = newR.addNewT();
                                    newT.setStringValue(sourceT.getStringValue());
                                }
                            }
                        }
                    } else {
                        newCell.addNewP();
                    }
                }
            }
            
            if (source.getTblPr() != null) {
                newTable.setTblPr((CTTblPr) source.getTblPr().copy());
            }
            if (source.getTblGrid() != null) {
                newTable.setTblGrid((CTTblGrid) source.getTblGrid().copy());
            }
            
            return newTable;
        } catch (Exception e) {
            System.err.println("复制表格失败: " + e.getMessage());
            return null;
        }
    }
    
    /**
     * 创建表格结构
     */
    private void createTableStructure(XWPFTable table, TestCase testCase) {
        // 第1行：4列
        XWPFTableRow firstRow = table.createRow();
        for (int i = 0; i < 4; i++) firstRow.createCell();
        setCellText(firstRow.getCell(0), "测试项名称");
        setCellText(firstRow.getCell(2), "标识");
        
        // 后续行：2列（带合并）
        String[] labels = {"测试内容", "测试策略与方法", "判定准则", "测试终止条件", "追踪关系"};
        for (String label : labels) {
            XWPFTableRow row = table.createRow();
            XWPFTableCell labelCell = row.createCell();
            setCellText(labelCell, label);
            
            XWPFTableCell dataCell = row.createCell();
            CTTcPr tcPr = dataCell.getCTTc().isSetTcPr() ? 
                dataCell.getCTTc().getTcPr() : dataCell.getCTTc().addNewTcPr();
            if (!tcPr.isSetGridSpan()) {
                tcPr.addNewGridSpan().setVal(BigInteger.valueOf(3));
            }
        }
    }
    
    /**
     * 清空表格数据
     */
    private void clearTableData(XWPFTable table) {
        for (int i = 0; i < table.getNumberOfRows(); i++) {
            XWPFTableRow row = table.getRow(i);
            if (row == null) continue;
            
            int cellCount = row.getTableCells().size();
            if (i == 0 && cellCount >= 4) {
                clearCell(row.getCell(1));
                clearCell(row.getCell(3));
            } else if (cellCount >= 2) {
                for (int j = 1; j < cellCount; j++) {
                    clearCell(row.getCell(j));
                }
            }
        }
    }
    
    /**
     * 填充表格数据
     */
    private void fillTableData(XWPFTable table, TestCase testCase) {
        Map<String, String> columnData = testCase.getColumnData();
        int rowCount = table.getNumberOfRows();
        if (rowCount == 0) return;
        
        XWPFTableRow firstRow = table.getRow(0);
        int firstRowCellCount = firstRow != null ? firstRow.getTableCells().size() : 0;
        
        if (firstRowCellCount >= 4) {
            // 4列格式
            String label0 = getCellText(firstRow.getCell(0));
            String label2 = getCellText(firstRow.getCell(2));
            
            String match0 = findMatchingColumn(label0, columnData.keySet());
            if (match0 != null) setCellText(firstRow.getCell(1), testCase.getColumnValue(match0));
            
            String match2 = findMatchingColumn(label2, columnData.keySet());
            if (match2 != null) setCellText(firstRow.getCell(3), testCase.getColumnValue(match2));
            
            // 后续行
            for (int i = 1; i < rowCount; i++) {
                XWPFTableRow row = table.getRow(i);
                if (row == null || row.getTableCells().size() < 2) continue;
                
                String label = getCellText(row.getCell(0));
                String match = findMatchingColumn(label, columnData.keySet());
                if (match != null) {
                    int dataIdx = row.getTableCells().size() >= 3 ? 2 : 1;
                    setCellText(row.getCell(dataIdx), testCase.getColumnValue(match));
                }
            }
        } else {
            // 2列格式
            for (int i = 0; i < rowCount; i++) {
                XWPFTableRow row = table.getRow(i);
                if (row == null || row.getTableCells().size() < 2) continue;
                
                String label = getCellText(row.getCell(0));
                String match = findMatchingColumn(label, columnData.keySet());
                if (match != null) {
                    int dataIdx = row.getTableCells().size() >= 3 ? 2 : 1;
                    setCellText(row.getCell(dataIdx), testCase.getColumnValue(match));
                }
            }
        }
    }
    
    /**
     * 设置表格样式
     */
    private void styleTable(CTTbl cttbl) {
        CTTblPr tblPr = cttbl.getTblPr();
        if (tblPr == null) tblPr = cttbl.addNewTblPr();
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
        
        CTBorder border = CTBorder.Factory.newInstance();
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(16));
        border.setColor("000000");

        borders.setTop(border);
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
    }

    // ==================== 工具方法 ====================
    
    /**
     * 验证文件格式
     */
    private void validateFormat(String path) throws Exception {
        if (path.toLowerCase().endsWith(".doc") && !path.toLowerCase().endsWith(".docx")) {
            throw new Exception("不支持旧版Word格式(.doc)，请转换为.docx格式");
        }
    }
    
    /**
     * 查找段落索引
     */
    private int findParagraphIndex(CTBody body, CTP paragraph) {
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i) == paragraph) return i;
        }
        return -1;
    }
    
    /**
     * 禁用段落编号
     */
    private void disableNumbering(CTP ctp) {
        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTNumPr numPr = ppr.isSetNumPr() ? ppr.getNumPr() : ppr.addNewNumPr();
        CTDecimalNumber numId = numPr.isSetNumId() ? numPr.getNumId() : numPr.addNewNumId();
        numId.setVal(BigInteger.ZERO);
    }
    
    /**
     * 更新段落文本
     */
    private void updateParagraphText(XWPFParagraph para, String text) {
        while (!para.getRuns().isEmpty()) {
            para.removeRun(0);
        }
        
        disableNumbering(para.getCTP());
        
        TextFormat fmt = subSectionFormat != null ? subSectionFormat : new TextFormat("黑体", 12, false, "4");
        
        XWPFRun run = para.createRun();
        run.setText(text);
        run.setFontFamily(fmt.fontFamily);
        run.setFontSize(fmt.fontSize);
        if (fmt.bold != null) run.setBold(fmt.bold);
    }
    
    /**
     * 删除段落
     */
    private void removeParagraph(XWPFDocument document, XWPFParagraph para) {
        CTBody body = document.getDocument().getBody();
        int index = findParagraphIndex(body, para.getCTP());
        if (index >= 0) {
            body.removeP(index);
        }
    }
    
    /**
     * 提取文本格式
     */
    private TextFormat extractTextFormat(XWPFParagraph para) {
        String styleId = para.getStyle();
        List<XWPFRun> runs = para.getRuns();
        
        if (runs == null || runs.isEmpty()) {
            return new TextFormat("黑体", 12, false, styleId != null ? styleId : "4");
        }
        
        XWPFRun run = runs.get(0);
        String fontFamily = run.getFontFamily();
        Integer fontSize = null;
        try {
            int size = run.getFontSize();
            if (size > 0) fontSize = size;
        } catch (Exception ignored) {}
        
        if (fontFamily == null || fontFamily.isEmpty()) fontFamily = "黑体";
        if (fontSize == null) fontSize = 12;
        
        return new TextFormat(fontFamily, fontSize, run.isBold(), styleId != null ? styleId : "4");
    }
    
    /**
     * 查找匹配的列名
     */
    private String findMatchingColumn(String label, Set<String> columns) {
        if (label == null || label.isEmpty()) return null;
        
        // 完全匹配
        for (String col : columns) {
            if (col.equals(label)) return col;
        }
        
        // 包含匹配
        for (String col : columns) {
            if (label.contains(col) || col.contains(label)) return col;
        }
        
        // 去空格匹配
        String labelNoSpace = label.replaceAll("\\s+", "");
        for (String col : columns) {
            if (col.replaceAll("\\s+", "").equals(labelNoSpace)) return col;
        }
        
        return null;
    }
    
    private String getCellText(XWPFTableCell cell) {
        if (cell == null) return "";
        StringBuilder sb = new StringBuilder();
        for (XWPFParagraph p : cell.getParagraphs()) {
            sb.append(p.getText());
        }
        return sb.toString().trim();
    }
    
    private void setCellText(XWPFTableCell cell, String text) {
        if (cell == null) return;
        while (!cell.getParagraphs().isEmpty()) {
            cell.removeParagraph(0);
        }
        XWPFParagraph para = cell.addParagraph();
        XWPFRun run = para.createRun();
        run.setText(text != null ? text : "");
    }
    
    private void clearCell(XWPFTableCell cell) {
        if (cell == null) return;
        while (!cell.getParagraphs().isEmpty()) {
            cell.removeParagraph(0);
        }
        cell.addParagraph();
    }
    
    // ==================== 内部类 ====================
    
    private record SectionInfo(String number, String name, XWPFParagraph paragraph) {
        @Override
        public String toString() {
            return number + " " + name;
        }
    }
    
    private record PlaceholderInfo(String parentNumber, String text, XWPFParagraph paragraph) {
        @Override
        public String toString() {
            return text;
        }
    }
    
    private record TextFormat(String fontFamily, Integer fontSize, Boolean bold, String styleId) {}
}
