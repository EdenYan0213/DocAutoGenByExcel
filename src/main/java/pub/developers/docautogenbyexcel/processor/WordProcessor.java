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
    
    // 通用章节编号模式：匹配任意层级的编号，如 1, 1.1, 1.1.1, 1.1.1.1 等
    private static final Pattern SECTION_NUMBER_PATTERN = Pattern.compile("^(\\d+(?:\\.\\d+)*)\\s+(.+)$");
    // 章节标题匹配模式：X.X 模块名称（保留向后兼容，但推荐使用 SECTION_NUMBER_PATTERN）
    private static final Pattern SECTION_PATTERN = SECTION_NUMBER_PATTERN;
    // 占位符匹配模式：X.x 或 X.X.x 等，表示该章节下有多个子章节
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("^(\\d+(?:\\.\\d+)*)\\.x\\s*(.+)?$", Pattern.CASE_INSENSITIVE);
    
    // 目录样式 ID: 22=toc1, 25=toc2, 16=toc3
    private static final java.util.Set<String> TOC_STYLES = java.util.Set.of("22", "25", "16");
    // 正文标题样式 ID: 2=Heading1, 3=Heading2, 4=Heading3, 5=Heading4
    private static final java.util.Set<String> HEADING_STYLES = java.util.Set.of("2", "3", "4", "5");
    
    /**
     * 内部类：保存Run的格式信息
     */
    private static class RunFormat {
        String fontFamily;
        Integer fontSize;
        Boolean bold;
        
        RunFormat(String fontFamily, Integer fontSize, Boolean bold) {
            this.fontFamily = fontFamily;
            this.fontSize = fontSize;
            this.bold = bold;
        }
    }
    
    /**
     * 内部类：保存子章节标题的格式信息（编号部分和内容部分可能不同）
     */
    private static class SubSectionFormat {
        RunFormat numberFormat;  // 编号部分格式（如 5.2.1）
        RunFormat contentFormat; // 内容部分格式（如 登录功能测试）
        String styleId;          // 段落样式ID
        
        SubSectionFormat() {
            // 默认格式
            this.numberFormat = new RunFormat("黑体", 12, false);
            this.contentFormat = new RunFormat("黑体", 12, false);
            this.styleId = "4";
        }
    }
    
    /**
     * 内部类：保存Caption的格式信息
     */
    private static class CaptionFormat {
        RunFormat format;
        String styleId;
        
        CaptionFormat() {
            // 默认格式
            this.format = new RunFormat("黑体", 12, false);
            this.styleId = "11";
        }
    }
    
    // 模板格式缓存
    private SubSectionFormat templateSubSectionFormat = null;
    private CaptionFormat templateCaptionFormat = null;
    
    // ==================== 目录编号辅助方法 ====================
    
    /** 获取章节编号的层级数（如 "5.3.1" 返回 3） */
    private static int getSectionLevel(String sectionNumber) {
        if (sectionNumber == null || sectionNumber.isEmpty()) return 0;
        return sectionNumber.split("\\.").length;
    }
    
    /** 判断 childNumber 是否是 parentNumber 的直接子节点 */
    private static boolean isDirectChild(String parentNumber, String childNumber) {
        if (parentNumber == null || childNumber == null || !childNumber.startsWith(parentNumber + ".")) {
            return false;
        }
        return childNumber.substring(parentNumber.length() + 1).matches("^\\d+$");
    }
    
    /** 比较两个章节编号的大小（支持任意层级），如 "1.2" < "1.10" < "2.1" */
    private static int compareSectionNumbers(String a, String b) {
        if (a == null || b == null) return (a == null) ? (b == null ? 0 : -1) : 1;
        String[] partsA = a.split("\\."), partsB = b.split("\\.");
        for (int i = 0; i < Math.min(partsA.length, partsB.length); i++) {
            int cmp = Integer.compare(Integer.parseInt(partsA[i]), Integer.parseInt(partsB[i]));
            if (cmp != 0) return cmp;
        }
        return Integer.compare(partsA.length, partsB.length);
    }
    
    // ==================== 样式判断辅助方法 ====================
    
    /** 判断是否为目录样式 */
    private static boolean isTocStyle(String styleName) {
        return styleName != null && 
            (TOC_STYLES.contains(styleName) || styleName.toLowerCase().startsWith("toc"));
    }
    
    /** 判断是否为标题样式（包括Heading 1-4） */
    private static boolean isHeadingStyle(String styleName) {
        return styleName != null && 
            (HEADING_STYLES.contains(styleName) || 
             styleName.toLowerCase().contains("heading") ||
             styleName.contains("标题") ||
             styleName.contains("程序标题"));
    }
    
    /** 判断是否为 Heading 2（主章节样式） */
    private static boolean isHeading2Style(String styleId) {
        return styleId != null && 
            (styleId.equals("3") || styleId.toLowerCase().contains("heading 2"));
    }
    
    /** 判断是否为主章节样式（Heading 1 或 Heading 2） */
    private static boolean isMainSectionStyle(String styleName) {
        return styleName != null && 
            (styleName.equals("2") || styleName.equals("3") || 
             styleName.toLowerCase().contains("heading 1") ||
             styleName.toLowerCase().contains("heading 2") ||
             styleName.contains("标题 1") || styleName.contains("标题 2"));
    }
    
    /** 判断是否为子章节样式（Heading 3 或 Heading 4） */
    private static boolean isSubSectionStyle(String styleName) {
        return styleName != null && 
            (styleName.equals("4") || styleName.equals("5") ||
             styleName.toLowerCase().contains("heading 3") || 
             styleName.toLowerCase().contains("heading 4") ||
             styleName.contains("标题 3") || styleName.contains("程序标题"));
    }
    
    // ==================== 通用辅助方法 ====================
    
    /** 在段落列表中查找段落的索引 */
    private static int findParagraphIndex(List<XWPFParagraph> paragraphs, XWPFParagraph target) {
        for (int i = 0; i < paragraphs.size(); i++) {
            if (paragraphs.get(i).getCTP() == target.getCTP()) {
                return i;
            }
        }
        return -1;
    }
    
    /** 在CTBody中查找段落的索引 */
    private static int findParagraphIndexInBody(CTBody body, CTP paragraph) {
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i) == paragraph) {
                return i;
            }
        }
        return -1;
    }
    
    /** 禁用段落的编号（设置 numId=0） */
    private static void disableParagraphNumbering(CTP ctp) {
        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTNumPr numPr = ppr.isSetNumPr() ? ppr.getNumPr() : ppr.addNewNumPr();
        CTDecimalNumber numId = numPr.isSetNumId() ? numPr.getNumId() : numPr.addNewNumId();
        numId.setVal(BigInteger.ZERO);
    }
    
    /** 设置Run的格式 */
    private static void applyRunFormat(XWPFRun run, RunFormat format) {
        run.setFontFamily(format.fontFamily);
        run.setFontSize(format.fontSize);
        if (format.bold != null) {
            run.setBold(format.bold);
        }
    }

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
            throw new Exception("""
                    不支持旧版Word格式(.doc文件)。请将文件转换为.docx格式后再使用。
                    转换方法：
                    1. 使用Microsoft Word打开.doc文件
                    2. 选择'文件' -> '另存为'
                    3. 在'文件类型'中选择'Word文档(*.docx)'
                    4. 保存后使用新的.docx文件""");
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
            if (!moduleDataMap.isEmpty()) {
                ModuleData firstModule = moduleDataMap.values().iterator().next();
                if (!firstModule.getTestCases().isEmpty()) {
                    TestCase firstCase = firstModule.getTestCases().get(0);
                    List<String> excelColumnNames = new java.util.ArrayList<>(firstCase.getColumnData().keySet());
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
                String parentNumber = placeholder.parentNumber; // 如 "5" 或 "5.3"
                String placeholderText = placeholder.text; // 如 "5.x" 或 "5.3.x 功能测试"
                
                // 在Excel中查找所有匹配的直接子模块（使用通用方法）
                List<String> matchedModules = new java.util.ArrayList<>();
                for (String moduleNumber : moduleDataMap.keySet()) {
                    // 使用通用方法判断是否是直接子节点
                    if (isDirectChild(parentNumber, moduleNumber)) {
                        matchedModules.add(moduleNumber);
                    }
                }
                
                // 按编号排序（支持任意层级）
                matchedModules.sort(WordProcessor::compareSectionNumbers);
                
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
                
                // 找到下一个主章节的位置作为边界（不在其后面插入）
                XWPFParagraph nextSectionPara = findNextMainSection(document, sectionPara, sectionNumber);
                int sectionBoundaryIndex = -1;
                if (nextSectionPara != null) {
                    for (int idx = 0; idx < body.sizeOfPArray(); idx++) {
                        if (body.getPArray(idx) == nextSectionPara.getCTP()) {
                            sectionBoundaryIndex = idx;
                            break;
                        }
                    }
                    System.out.println("模块" + sectionNumber + "的边界（下一个主章节）位置: " + sectionBoundaryIndex + " (" + nextSectionPara.getText() + ")");
                }
                
                // 从第一个已存在的子章节提取格式（只提取一次）
                if (templateSubSectionFormat == null && !existingSubSectionParas.isEmpty()) {
                    XWPFParagraph firstSubSection = existingSubSectionParas.get(0);
                    templateSubSectionFormat = extractSubSectionFormat(firstSubSection);
                    System.out.println("已从模板子章节提取格式");
                    
                    // 同时提取Caption格式
                    XWPFParagraph captionPara = findCaptionAfterSubSection(document, firstSubSection);
                    if (captionPara != null) {
                        templateCaptionFormat = extractCaptionFormat(captionPara);
                        System.out.println("已从模板Caption提取格式");
                    }
                }
                
                XWPFParagraph currentInsertPoint = lastElementInSection != null ? lastElementInSection : sectionPara;
                System.out.println("初始插入点: " + currentInsertPoint.getText());
                
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
                        System.out.println("当前插入点: " + currentInsertPoint.getText());
                        
                        // 检查当前插入点是否超过了章节边界
                        XWPFParagraph actualInsertPoint = currentInsertPoint;
                        if (sectionBoundaryIndex > 0) {
                            int currentIndex = -1;
                            for (int idx = 0; idx < body.sizeOfPArray(); idx++) {
                                if (body.getPArray(idx) == currentInsertPoint.getCTP()) {
                                    currentIndex = idx;
                                    break;
                                }
                            }
                            if (currentIndex >= sectionBoundaryIndex) {
                                // 当前插入点已经超过边界，使用边界前一个位置
                                System.out.println("警告: 插入点(" + currentIndex + ")超过边界(" + sectionBoundaryIndex + ")，调整到边界前");
                                // 在边界之前插入，使用下一个主章节的前一个段落
                                CTP prevPara = body.getPArray(sectionBoundaryIndex - 1);
                                actualInsertPoint = new XWPFParagraph(prevPara, document);
                            }
                        }
                        
                        XWPFParagraph subSectionPara = createSubSectionParagraphBeforeBoundary(
                            document, actualInsertPoint, subSectionNumber, testCase.getTestName(), 
                            sectionBoundaryIndex);
                        
                        // 创建表格标题（Caption）
                        XWPFParagraph captionPara = createTableCaption(document, subSectionPara, 
                            subSectionNumber, testCase.getTestName());
                        
                        // 始终为新创建的子章节创建新表格（不复用已存在的表格）
                        if (templateTable != null) {
                            // 如果有模板，复制模板表格（在Caption后面）
                            System.out.println("复制模板表格到子章节" + subSectionNumber + "后");
                            CTTbl newCttbl = copyTable(document, captionPara.getCTP(), templateTable.getCTTbl());
                            if (newCttbl != null) {
                                XWPFTable newTable = new XWPFTable(newCttbl, document);
                                System.out.println("模板表格复制成功，行数: " + newTable.getNumberOfRows());
                                clearTableDataColumns(newTable);
                                fillTableData(newTable, testCase);
                                System.out.println("表格数据填充完成");
                            } else {
                                System.err.println("复制模板表格失败");
                            }
                        } else {
                            // 如果没有模板，创建新表格
                            insertNewTableAfterParagraph(document, subSectionPara.getCTP(), testCase);
                        }
                        
                        // 更新插入点为表格后面（确保下一个子章节在表格后面创建）
                        currentInsertPoint = findInsertPointAfterTable(document, subSectionPara);
                        System.out.println("更新插入点到表格后面: " + currentInsertPoint.getText());
                        
                        // 每次插入新内容后，边界位置应该相应增加
                        // （子章节标题 + Caption + 表格 = 大约3-4个元素）
                        if (sectionBoundaryIndex > 0) {
                            sectionBoundaryIndex += 4;  // 预估每个子章节增加4个元素
                            System.out.println("更新边界位置到: " + sectionBoundaryIndex);
                        }
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
            throw new Exception("""
                    不支持旧版Word格式(.doc文件)。请将文件转换为.docx格式后再使用。
                    转换方法：
                    1. 使用Microsoft Word打开.doc文件
                    2. 选择'文件' -> '另存为'
                    3. 在'文件类型'中选择'Word文档(*.docx)'
                    4. 保存后使用新的.docx文件""", e);
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
     * 扫描Word文档中已存在的子章节（任意层级，如5.3.1、5.3.2、5.3.1.1等）
     * 从目录读取子章节编号，然后在正文中查找对应段落
     *
     * @param document Word文档
     * @return 子章节信息列表
     */
    private List<SubSectionInfo> scanExistingSubSections(XWPFDocument document) {
        List<SubSectionInfo> subSections = new java.util.ArrayList<>();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        // 存储目录中的子章节: 编号 -> 名称
        java.util.Map<String, String> tocSubSections = new java.util.LinkedHashMap<>();
        
        // 第一遍：从目录中读取子章节编号和名称（使用通用模式匹配任意层级）
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }

                String styleName = para.getStyle();
                if (!isTocStyle(styleName)) {
                    continue;
                }

                // 使用通用模式匹配任意层级的章节编号
                Matcher matcher = SECTION_NUMBER_PATTERN.matcher(text.trim());
                if (matcher.matches()) {
                    String sectionNumber = matcher.group(1);
                    // 只处理层级大于2的子章节（如 X.X.X 或更深层级）
                    if (getSectionLevel(sectionNumber) <= 2) {
                        continue;
                    }
                    String sectionName = matcher.group(2).trim();
                    // 去除页码
                    int tabIndex = sectionName.indexOf('\t');
                    if (tabIndex > 0) {
                        sectionName = sectionName.substring(0, tabIndex).trim();
                    }
                    tocSubSections.put(sectionNumber, sectionName);
                }
            } catch (Exception ignored) {
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
                    if (isTocStyle(styleName)) {
                        continue;
                    }
                    
                    String trimmed = text.trim();
                    // 匹配带编号或仅名称
                    if ((trimmed.startsWith(subNumber + " ") || trimmed.equals(subName)) && isHeadingStyle(styleName)) {
                        subSections.add(new SubSectionInfo(subNumber, para));
                        System.out.println("找到正文中的子章节: " + subNumber + " -> " + trimmed);
                        break;
                    }
                } catch (Exception ignored) {
                }
            }
        }
        
        return subSections;
    }
    
    /**
     * 更新段落文本（去除自动编号，使用模板格式）
     */
    private void updateParagraphText(XWPFParagraph paragraph, String newText) {
        // 清除原有内容
        while (!paragraph.getRuns().isEmpty()) {
            paragraph.removeRun(0);
        }
        
        // 禁用从样式继承的编号
        disableParagraphNumbering(paragraph.getCTP());
        
        // 使用模板格式或默认格式
        SubSectionFormat subFmt = templateSubSectionFormat != null ? templateSubSectionFormat : new SubSectionFormat();
        
        // 解析编号和内容（格式如 "5.2.1 登录功能测试"）
        int spaceIndex = newText.indexOf(' ');
        if (spaceIndex > 0) {
            String number = newText.substring(0, spaceIndex + 1);  // 包含空格
            String content = newText.substring(spaceIndex + 1);
            
            // 编号部分
            XWPFRun numRun = paragraph.createRun();
            numRun.setText(number);
            applyRunFormat(numRun, subFmt.numberFormat);
            
            // 内容部分
            XWPFRun contentRun = paragraph.createRun();
            contentRun.setText(content);
            applyRunFormat(contentRun, subFmt.contentFormat);
        } else {
            // 没有空格，整体使用内容格式
            XWPFRun run = paragraph.createRun();
            run.setText(newText);
            applyRunFormat(run, subFmt.contentFormat);
        }
    }
    
    /**
     * 更新Caption文本（使用模板格式）
     */
    private void updateCaptionText(XWPFParagraph paragraph, String newText) {
        // 清除原有内容
        while (!paragraph.getRuns().isEmpty()) {
            paragraph.removeRun(0);
        }
        
        // 使用模板格式或默认格式
        CaptionFormat captionFmt = templateCaptionFormat != null ? templateCaptionFormat : new CaptionFormat();
        
        // 添加新文本
        XWPFRun run = paragraph.createRun();
        run.setText(newText);
        applyRunFormat(run, captionFmt.format);
    }
    
    /**
     * 更新表格标题（Caption）
     * 查找子章节后面的Caption样式段落并更新
     */
    private void updateTableCaption(XWPFDocument document, XWPFParagraph subSectionPara, 
                                     String subSectionNumber, String testName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        int subSectionIndex = findParagraphIndex(paragraphs, subSectionPara);
        if (subSectionIndex == -1) {
            return;
        }
        
        // 查找子章节后的第一个Caption样式段落
        for (int i = subSectionIndex + 1; i < paragraphs.size() && i < subSectionIndex + 5; i++) {
            XWPFParagraph para = paragraphs.get(i);
            String styleName = para.getStyle();
            
            if (isCaptionStyle(styleName)) {
                String newCaption = "表" + subSectionNumber + " " + testName + "测试";
                updateCaptionText(para, newCaption);
                System.out.println("更新表格标题为: " + newCaption);
                break;
            }
            
            // 如果遇到Heading样式，停止查找
            String text = para.getText();
            if (text != null && !text.trim().isEmpty() && isHeadingStyle(styleName)) {
                break;
            }
        }
    }
    
    /**
     * 从Run中提取格式信息（如果Run没有直接格式，则从段落样式推断）
     */
    private RunFormat extractRunFormat(XWPFRun run) {
        if (run == null) {
            return new RunFormat("黑体", 12, false);
        }
        
        String fontFamily = run.getFontFamily();
        Integer fontSize = null;
        Boolean bold = run.isBold();
        
        try {
            int size = run.getFontSize();
            if (size > 0) {
                fontSize = size;
            }
        } catch (Exception e) {
            // 忽略
        }
        
        // 如果Run没有直接格式，从段落样式推断默认格式
        if (fontFamily == null || fontFamily.isEmpty() || fontSize == null) {
            XWPFParagraph para = (XWPFParagraph) run.getParent();
            if (para != null) {
                String styleId = para.getStyle();
                // 根据样式ID推断格式
                // Heading 3 (ID=4): 黑体
                // Caption (ID=11): 黑体
                if (styleId != null) {
                    if (styleId.equals("4") || styleId.toLowerCase().contains("heading")) {
                        if (fontFamily == null || fontFamily.isEmpty()) {
                            fontFamily = "黑体";
                        }
                    } else if (styleId.equals("11") || styleId.toLowerCase().contains("caption")) {
                        if (fontFamily == null || fontFamily.isEmpty()) {
                            fontFamily = "黑体";
                        }
                    }
                }
            }
        }
        
        // 最终默认值
        if (fontFamily == null || fontFamily.isEmpty()) {
            fontFamily = "黑体";
        }
        if (fontSize == null) {
            fontSize = 12;  // 默认小四号
        }
        
        return new RunFormat(fontFamily, fontSize, bold);
    }
    
    /**
     * 从模板子章节段落中提取格式
     * 分析编号部分和内容部分的格式（可能不同）
     * 对于 Heading 3 样式，Word 模板通常使用自动编号：
     * - 编号部分：黑体 10号字体
     * - 内容部分：黑体 12pt (小四号)
     */
    private SubSectionFormat extractSubSectionFormat(XWPFParagraph templatePara) {
        SubSectionFormat format = new SubSectionFormat();
        
        if (templatePara == null) {
            return format;
        }
        
        // 获取样式ID
        String styleId = templatePara.getStyle();
        if (styleId != null) {
            format.styleId = styleId;
        }
        
        List<XWPFRun> runs = templatePara.getRuns();
        
        // 对于 Heading 3 样式，通常使用自动编号，内容文本没有编号
        // 根据模板的样式，默认格式为：
        // - 编号：黑体 10号
        // - 内容：黑体 12pt
        boolean isHeading3 = styleId != null && 
            (styleId.equals("4") || styleId.toLowerCase().contains("heading"));
        
        if (isHeading3) {
            // Heading 3 的默认格式
            format.numberFormat = new RunFormat("黑体", 10, false);  // 编号用10号字体
            format.contentFormat = new RunFormat("黑体", 12, false); // 内容用12pt (小四号)
        } else if (runs != null && !runs.isEmpty()) {
            // 其他样式，从 Run 中提取
            if (runs.size() >= 2) {
                format.numberFormat = extractRunFormat(runs.get(0));
                format.contentFormat = extractRunFormat(runs.get(1));
            } else {
                RunFormat singleFormat = extractRunFormat(runs.get(0));
                format.numberFormat = singleFormat;
                format.contentFormat = singleFormat;
            }
        }
        
        System.out.println("提取子章节格式: 编号[" + format.numberFormat.fontFamily + ", " + 
            format.numberFormat.fontSize + "pt], 内容[" + format.contentFormat.fontFamily + ", " + 
            format.contentFormat.fontSize + "pt]");
        
        return format;
    }
    
    /**
     * 从模板Caption段落中提取格式
     * Caption 样式默认格式：
     * - 黑体 12pt (小四号)
     */
    private CaptionFormat extractCaptionFormat(XWPFParagraph templateCaption) {
        CaptionFormat format = new CaptionFormat();
        
        if (templateCaption == null) {
            return format;
        }
        
        // 获取样式ID
        String styleId = templateCaption.getStyle();
        if (styleId != null) {
            format.styleId = styleId;
        }
        
        // 对于 Caption 样式，使用默认格式：黑体 12pt
        boolean isCaption = styleId != null && 
            (styleId.equals("11") || styleId.toLowerCase().contains("caption"));
        
        if (isCaption) {
            format.format = new RunFormat("黑体", 12, false);
        } else {
            List<XWPFRun> runs = templateCaption.getRuns();
            if (runs != null && !runs.isEmpty()) {
                format.format = extractRunFormat(runs.get(0));
            }
        }
        
        System.out.println("提取Caption格式: [" + format.format.fontFamily + ", " + 
            format.format.fontSize + "pt, bold=" + format.format.bold + "]");
        
        return format;
    }
    
    /**
     * 在子章节段落后查找Caption段落（用于提取格式）
     */
    private XWPFParagraph findCaptionAfterSubSection(XWPFDocument document, XWPFParagraph subSectionPara) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        int subIndex = findParagraphIndex(paragraphs, subSectionPara);
        if (subIndex == -1) {
            return null;
        }
        
        // 在子章节后找Caption（通常在2-3个段落内）
        for (int i = subIndex + 1; i < paragraphs.size() && i < subIndex + 5; i++) {
            XWPFParagraph para = paragraphs.get(i);
            String styleName = para.getStyle();
            String text = para.getText();
            
            // Caption样式或以"表"开头的文本
            if (isCaptionStyle(styleName) || (text != null && text.trim().startsWith("表"))) {
                return para;
            }
        }
        
        return null;
    }
    
    /** 判断是否为Caption样式 */
    private static boolean isCaptionStyle(String styleName) {
        return styleName != null && (styleName.equalsIgnoreCase("Caption") ||
            styleName.equals("11") || styleName.contains("题注"));
    }

    /**
     * 创建表格标题（Caption）
     */
    private XWPFParagraph createTableCaption(XWPFDocument document, XWPFParagraph afterPara,
                                              String subSectionNumber, String testName) {
        CTBody body = document.getDocument().getBody();
        
        int afterIndex = findParagraphIndexInBody(body, afterPara.getCTP());
        if (afterIndex == -1) {
            afterIndex = body.sizeOfPArray() - 1;
        }
        
        // 在afterIndex+1位置插入Caption段落
        CTP ctp = body.insertNewP(afterIndex + 1);
        XWPFParagraph para = new XWPFParagraph(ctp, document);
        
        // 使用模板格式或默认格式
        CaptionFormat captionFmt = templateCaptionFormat != null ? templateCaptionFormat : new CaptionFormat();
        
        // 设置Caption样式
        try {
            para.setStyle(captionFmt.styleId);
        } catch (Exception e) {
            // 忽略样式设置失败
        }
        
        // 设置Caption内容
        String captionText = "表" + subSectionNumber + " " + testName + "测试";
        XWPFRun run = para.createRun();
        run.setText(captionText);
        applyRunFormat(run, captionFmt.format);
        
        // 设置居中对齐
        para.setAlignment(ParagraphAlignment.CENTER);
        
        System.out.println("创建表格标题: " + captionText);
        
        return para;
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
        return createSubSectionParagraphBeforeBoundary(document, afterPara, moduleNumber, testName, -1);
    }
    
    /**
     * 创建子章节标题段落（带边界检查）
     * 确保插入位置不超过指定的边界
     * 
     * @param boundaryIndex 边界索引，-1表示不检查边界
     */
    private XWPFParagraph createSubSectionParagraphBeforeBoundary(XWPFDocument document, XWPFParagraph afterPara, 
                                                     String moduleNumber, String testName, int boundaryIndex) {
        CTBody body = document.getDocument().getBody();
        
        int afterIndex = findParagraphIndexInBody(body, afterPara.getCTP());
        if (afterIndex == -1) {
            System.err.println("找不到插入点段落，使用末尾插入");
            afterIndex = body.sizeOfPArray() - 1;
        }
        
        // 计算实际插入位置
        int insertIndex = afterIndex + 1;
        
        // 如果有边界限制，确保不超过边界
        if (boundaryIndex > 0 && insertIndex >= boundaryIndex) {
            insertIndex = boundaryIndex;
            System.out.println("调整插入位置到边界: " + insertIndex);
        }
        
        // 在insertIndex位置插入新段落（子章节标题）
        CTP ctp = body.insertNewP(insertIndex);
        XWPFParagraph para = new XWPFParagraph(ctp, document);

        // 使用模板格式或默认格式
        SubSectionFormat subFmt = templateSubSectionFormat != null ? templateSubSectionFormat : new SubSectionFormat();
        
        // 设置样式
        try {
            para.setStyle(subFmt.styleId);
        } catch (Exception e) {
            // 忽略样式设置失败
        }
        
        // 禁用从样式继承的编号
        disableParagraphNumbering(ctp);
        
        // 设置标题内容（使用模板格式）
        String subTitle = moduleNumber + " " + testName + "测试";
        
        // 编号部分
        XWPFRun numRun = para.createRun();
        numRun.setText(moduleNumber + " ");
        applyRunFormat(numRun, subFmt.numberFormat);
        
        // 内容部分
        XWPFRun contentRun = para.createRun();
        contentRun.setText(testName + "测试");
        applyRunFormat(contentRun, subFmt.contentFormat);

        // 设置段落格式
        para.setAlignment(ParagraphAlignment.LEFT);
        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(120));
        
        System.out.println("在位置 " + insertIndex + " 插入子章节: " + subTitle);
        
        return para;
    }

    /** 扫描Word文档中的所有章节编号（从目录TOC读取） */
    private List<String> scanWordSections(XWPFDocument document) {
        List<String> sections = new java.util.ArrayList<>();
        for (XWPFParagraph para : document.getParagraphs()) {
            try {
                String text = para.getText();
                if (text == null || text.trim().isEmpty() || !isTocStyle(para.getStyle())) continue;
                Matcher m = SECTION_NUMBER_PATTERN.matcher(text.trim());
                if (m.matches() && !sections.contains(m.group(1))) {
                    sections.add(m.group(1));
                }
            } catch (Exception e) { /* ignore */ }
        }
        return sections;
    }

    /** 查找章节标题段落（先从目录获取名称，再在正文中查找） */
    private XWPFParagraph findSectionParagraph(XWPFDocument document, String moduleNumber) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        String tocName = null;
        
        // 第一遍：从目录中提取章节名称
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                if (text == null || !isTocStyle(para.getStyle())) continue;
                Matcher m = SECTION_PATTERN.matcher(text.trim());
                if (m.matches() && moduleNumber.equals(m.group(1))) {
                    tocName = m.group(2).split("\t")[0].trim();
                    System.out.println("从目录提取章节名称: " + moduleNumber + " -> " + tocName);
                    break;
                }
            } catch (Exception e) { /* ignore */ }
        }
        
        // 第二遍：在正文中查找
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                String style = para.getStyle();
                if (text == null || isTocStyle(style)) continue;
                String trimmed = text.trim();
                Matcher m = SECTION_PATTERN.matcher(trimmed);
                // 通过编号匹配
                if (m.matches() && moduleNumber.equals(m.group(1))) {
                    System.out.println("在正文中找到章节(通过编号): " + trimmed + " [" + style + "]");
                    return para;
                }
                // 通过名称+样式匹配
                if (trimmed.equals(tocName) && isHeadingStyle(style)) {
                    System.out.println("在正文中找到章节(通过名称+样式): " + trimmed + " [" + style + "]");
                    return para;
                }
            } catch (Exception e) { /* ignore */ }
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
     * 在段落后强制创建新表格（不查找已存在的表格）
     */
    private void insertNewTableAfterParagraph(XWPFDocument document, CTP paragraph, TestCase testCase) {
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
        System.out.println("创建新表格，段落: " + "非空");
        
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
                clearTableDataColumns(newTable);
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
//            e.printStackTrace();
            return null;
        }
    }
    
    /**
     * 查找段落后表格的CTTbl对象
     * 支持跳过空行段落查找表格
     * 只查找紧邻段落的表格，不查找距离较远的表格
     */
    private CTTbl findTableCttblAfterParagraph(CTBody body, CTP paragraph) {
        int pIndex = findParagraphIndexInBody(body, paragraph);
        if (pIndex == -1) {
            return null;
        }
        
        // 调试信息：检查段落内容
        try {
            XWPFParagraph para = new XWPFParagraph(paragraph, null);
            String paraText = para.getText();
            // 如果段落是子章节标题（任意层级，如"4.3.2 注册功能测试"），且是新创建的，后面不应该有表格
            if (paraText != null) {
                Matcher matcher = SECTION_NUMBER_PATTERN.matcher(paraText.trim());
                if (matcher.matches() && getSectionLevel(matcher.group(1)) > 2) {
                    // 这是子章节（层级>2），检查是否是刚创建的
                    // 如果是新创建的子章节，后面不应该有表格，应该返回null
                }
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
            if (obj instanceof CTP nextPara) {
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

        return closestTable;
    }
    
    /** 找到下一个主章节（用于确定当前章节的边界） */
    private XWPFParagraph findNextMainSection(XWPFDocument document, XWPFParagraph currentPara, String currentNumber) {
        boolean found = false;
        for (XWPFParagraph para : document.getParagraphs()) {
            if (para.getCTP() == currentPara.getCTP()) { found = true; continue; }
            if (!found) continue;
            String text = para.getText();
            String style = para.getStyle();
            if (text != null && !text.trim().isEmpty() && !isTocStyle(style) && isHeading2Style(style)) {
                System.out.println("找到下一个主章节: " + text.trim());
                return para;
            }
        }
        return null;
    }
    
    /** 找到章节的最后一个子章节（用于确定插入点） */
    private XWPFParagraph findLastElementInSection(XWPFDocument document, XWPFParagraph sectionPara, String sectionNumber) {
        List<XWPFParagraph> subs = findExistingSubSectionsInSection(document, sectionPara, sectionNumber);
        return subs.isEmpty() ? sectionPara : subs.get(subs.size() - 1);
    }
    
    /** 找到章节下所有已存在的子章节段落 */
    private List<XWPFParagraph> findExistingSubSectionsInSection(XWPFDocument document, 
                                                                   XWPFParagraph sectionPara, String sectionNumber) {
        List<XWPFParagraph> subSections = new java.util.ArrayList<>();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        int sectionIndex = findParagraphIndex(paragraphs, sectionPara);
        if (sectionIndex == -1) return subSections;
        
        for (int i = sectionIndex + 1; i < paragraphs.size(); i++) {
            XWPFParagraph para = paragraphs.get(i);
            String text = para.getText();
            String style = para.getStyle();
            if (text == null || text.trim().isEmpty() || isTocStyle(style)) continue;
            if (isMainSectionStyle(style)) break;
            if (isSubSectionStyle(style)) subSections.add(para);
        }
        return subSections;
    }
    
    /** 找到表格后的插入点 */
    private XWPFParagraph findInsertPointAfterTable(XWPFDocument document, XWPFParagraph beforeTablePara) {
        CTBody body = document.getDocument().getBody();
        CTTbl table = findTableCttblAfterParagraph(body, beforeTablePara.getCTP());
        if (table == null) return beforeTablePara;
        
        org.apache.xmlbeans.XmlCursor tblCursor = table.newCursor();
        tblCursor.toEndToken();
        tblCursor.toNextToken();
        
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            org.apache.xmlbeans.XmlCursor pCursor = body.getPArray(i).newCursor();
            boolean isAfter = pCursor.comparePosition(tblCursor) > 0;
            pCursor.close();
            if (isAfter) {
                tblCursor.close();
                XWPFParagraph p = new XWPFParagraph(body.getPArray(i), document);
                String text = p.getText();
                return (text == null || text.trim().isEmpty()) ? p : new XWPFParagraph(body.insertNewP(i), document);
            }
        }
        tblCursor.close();
        return new XWPFParagraph(body.addNewP(), document);
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
        Map<String, String> columnData = testCase.getColumnData();
        
        int rowCount = table.getNumberOfRows();
        if (rowCount == 0) {
            return;
        }
        
        XWPFTableRow firstRow = table.getRow(0);
        int firstRowCellCount = firstRow != null ? firstRow.getTableCells().size() : 0;
        int startRow = 0;
        
        // 如果第一行是4列格式，特殊处理第一行
        if (firstRowCellCount == 4) {
            fillFourColumnFirstRow(firstRow, testCase, columnData.keySet());
            startRow = 1;
        }
        
        // 处理数据行
        for (int i = startRow; i < rowCount; i++) {
            XWPFTableRow row = table.getRow(i);
            if (row == null || row.getTableCells().size() < 2) {
                continue;
            }
            fillRowByLabel(row, testCase, columnData.keySet());
        }
    }
    
    /** 填充4列格式的第一行 */
    private void fillFourColumnFirstRow(XWPFTableRow firstRow, TestCase testCase, java.util.Set<String> columnNames) {
        String firstCol = getCellText(firstRow.getCell(0)).trim();
        String thirdCol = getCellText(firstRow.getCell(2)).trim();
        
        String firstColMatch = findMatchingColumn(firstCol, columnNames);
        if (firstColMatch != null) {
            setCellValue(firstRow.getCell(1), testCase.getColumnValue(firstColMatch));
        }
        
        String thirdColMatch = findMatchingColumn(thirdCol, columnNames);
        if (thirdColMatch != null) {
            setCellValue(firstRow.getCell(3), testCase.getColumnValue(thirdColMatch));
        }
    }
    
    /** 根据标签列填充数据到数据列 */
    private void fillRowByLabel(XWPFTableRow row, TestCase testCase, java.util.Set<String> columnNames) {
        String label = getCellText(row.getCell(0)).trim();
        String matchedColumn = findMatchingColumn(label, columnNames);
        
        if (matchedColumn != null) {
            String value = testCase.getColumnValue(matchedColumn);
            // 由于单元格合并，优先填充单元格2（如果存在），否则填充单元格1
            int dataCell = row.getTableCells().size() >= 3 ? 2 : 1;
            setCellValue(row.getCell(dataCell), value);
        }
    }
    
    /** 清空表格的数据列（只保留标签列） */
    private void clearTableDataColumns(XWPFTable table) {
        for (int i = 0; i < table.getNumberOfRows(); i++) {
            XWPFTableRow row = table.getRow(i);
            if (row == null) continue;
            int cellCount = row.getTableCells().size();
            if (i == 0 && cellCount >= 4) {
                clearCellContent(row.getCell(1));
                clearCellContent(row.getCell(3));
            } else if (cellCount >= 2) {
                for (int j = 1; j < cellCount; j++) clearCellContent(row.getCell(j));
            }
        }
    }
    
    /** 清空单元格内容 */
    private void clearCellContent(XWPFTableCell cell) {
        if (cell == null) return;
        while (!cell.getParagraphs().isEmpty()) cell.removeParagraph(0);
        cell.addParagraph();
    }
    
    /** 在Excel列名集合中查找匹配的列名（支持完全匹配和模糊匹配） */
    private String findMatchingColumn(String label, java.util.Set<String> cols) {
        if (label == null || label.isEmpty()) return null;
        if (cols.contains(label)) return label;
        for (String col : cols) {
            if (label.contains(col) || col.contains(label)) return col;
        }
        String clean = label.replace(" ", "").replace("　", "");
        for (String col : cols) {
            if (clean.equals(col.replace(" ", "").replace("　", ""))) return col;
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
        
        // 填充表头行（第1行，如果是4列格式）
        XWPFTableRow headerRow = table.getRow(0);
        if (headerRow != null && headerRow.getTableCells().size() >= 4) {
            fillFourColumnFirstRow(headerRow, testCase, columnData.keySet());
        }
        
        // 填充数据行（第2行及以后）
        for (int i = 1; i < currentRowCount; i++) {
            XWPFTableRow row = table.getRow(i);
            if (row == null || row.getTableCells().size() < 2) {
                continue;
            }
            fillRowByLabel(row, testCase, columnData.keySet());
        }
    }
    
    /** 获取单元格文本内容 */
    private String getCellText(XWPFTableCell cell) {
        if (cell == null) return "";
        StringBuilder sb = new StringBuilder();
        cell.getParagraphs().forEach(p -> sb.append(p.getText()));
        return sb.toString().trim();
    }
    
    /** 设置单元格值 */
    private void setCellValue(XWPFTableCell cell, String value) {
        if (cell == null) return;
        while (!cell.getParagraphs().isEmpty()) cell.removeParagraph(0);
        cell.addParagraph().createRun().setText(value != null ? value : "");
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