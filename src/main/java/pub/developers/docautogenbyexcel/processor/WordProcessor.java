package pub.developers.docautogenbyexcel.processor;

import org.apache.poi.xwpf.usermodel.*;
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
                if (existingTable != null) {
                    // 如果已有表格，先为第一个测试用例创建子章节（4.3.1），然后填充表格
                    TestCase firstTestCase = moduleData.getTestCases().get(0);
                    String firstSubSectionNumber = sectionNumber + ".1";
                    
                    // 检查是否已经存在4.3.1子章节
                    XWPFParagraph firstSubSectionPara = findSectionParagraph(document, firstSubSectionNumber);
                    if (firstSubSectionPara == null) {
                        // 如果不存在，创建第一个子章节（在表格前插入）
                        // 找到表格在body中的位置
                        int tableIndex = -1;
                        for (int i = 0; i < body.sizeOfTblArray(); i++) {
                            if (body.getTblArray(i) == existingTable) {
                                tableIndex = i;
                                break;
                            }
                        }
                        
                        // 找到表格前的段落位置（应该是sectionPara）
                        int sectionParaIndex = -1;
                        for (int i = 0; i < body.sizeOfPArray(); i++) {
                            if (body.getPArray(i) == sectionPara.getCTP()) {
                                sectionParaIndex = i;
                                break;
                            }
                        }
                        
                        if (sectionParaIndex >= 0) {
                            // 在章节段落后插入第一个子章节
                            firstSubSectionPara = createSubSectionParagraph(
                                document, sectionPara, firstSubSectionNumber, firstTestCase.getTestName());
                        }
                    }
                    
                    // 填充第一个表格
                    if (firstSubSectionPara != null) {
                        // 查找子章节后的表格（应该是existingTable）
                        CTTbl tableAfterSubSection = findTableCttblAfterParagraph(body, firstSubSectionPara.getCTP());
                        if (tableAfterSubSection != null) {
                            XWPFTable table = new XWPFTable(tableAfterSubSection, document);
                            fillTableData(table, firstTestCase);
                        } else {
                            // 如果找不到，使用existingTable
                            XWPFTable table = new XWPFTable(existingTable, document);
                            fillTableData(table, firstTestCase);
                        }
                    } else {
                        // 如果创建失败，直接填充表格
                        XWPFTable table = new XWPFTable(existingTable, document);
                        fillTableData(table, firstTestCase);
                    }
                    System.out.println("模块" + sectionNumber + "第一个表格填充完成");
                    
                    // 保存第一个表格作为模板，用于复制
                    XWPFTable templateTable = new XWPFTable(existingTable, document);
                    
                    // 如果这是第一个模板表格，保存为全局模板
                    if (globalTemplateTable == null) {
                        globalTemplateTable = templateTable;
                        System.out.println("保存模块" + sectionNumber + "的表格作为全局模板");
                    }
                    
                    // 如果还有更多测试用例，为剩余的测试用例创建子章节和表格
                    if (moduleData.getTestCaseCount() > 1) {
                        // 找到第一个表格后的位置，作为插入点
                        // 使用更可靠的方法：找到表格在body中的位置，然后找到表格后的第一个段落
                        int tableIndexInBody = -1;
                        for (int i = 0; i < body.sizeOfTblArray(); i++) {
                            if (body.getTblArray(i) == existingTable) {
                                tableIndexInBody = i;
                                break;
                            }
                        }
                        
                        // 找到第一个表格后的位置，作为插入点
                        // 方法：直接在第一个表格后创建空段落作为插入点，不依赖查找逻辑
                        CTBody body2 = document.getDocument().getBody();
                        XWPFParagraph insertAfterPara = null;
                        
                        // 找到第一个子章节后的表格
                        CTTbl firstTable = findTableCttblAfterParagraph(body, firstSubSectionPara.getCTP());
                        if (firstTable == null) {
                            firstTable = existingTable; // 如果找不到，使用existingTable
                        }
                        
                        if (firstTable != null) {
                            // 在表格后创建空段落作为插入点
                            org.apache.xmlbeans.XmlCursor tableCursor = firstTable.newCursor();
                            tableCursor.toEndToken();
                            tableCursor.toNextToken();
                            
                            // 找到表格后的段落位置
                            int insertPos = -1;
                            for (int k = 0; k < body2.sizeOfPArray(); k++) {
                                CTP checkPara = body2.getPArray(k);
                                org.apache.xmlbeans.XmlCursor checkCursor = checkPara.newCursor();
                                if (checkCursor.comparePosition(tableCursor) > 0) {
                                    insertPos = k;
                                    checkCursor.close();
                                    break;
                                }
                                checkCursor.close();
                            }
                            tableCursor.close();
                            
                            if (insertPos >= 0) {
                                // 在表格后插入空段落
                                CTP newPara = body2.insertNewP(insertPos);
                                insertAfterPara = new XWPFParagraph(newPara, document);
                                System.out.println("在第一个表格后创建空段落作为插入点（位置: " + insertPos + "）");
                            } else {
                                // 如果找不到，使用firstSubSectionPara
                                insertAfterPara = firstSubSectionPara;
                                System.out.println("使用第一个子章节作为插入点");
                            }
                        } else {
                            // 如果找不到表格，使用firstSubSectionPara
                            insertAfterPara = firstSubSectionPara;
                            System.out.println("找不到第一个表格，使用第一个子章节作为插入点");
                        }
                        
                        // 为剩余的测试用例创建子章节和表格
                        List<TestCase> remainingTestCases = moduleData.getTestCases().subList(1, moduleData.getTestCases().size());
                        System.out.println("模块" + sectionNumber + "需要为" + remainingTestCases.size() + "个剩余测试用例创建子章节和表格");
                        
                        // 如果insertAfterPara还是null，在第一个表格后创建一个空段落作为插入点
                        if (insertAfterPara == null && existingTable != null) {
                            System.out.println("在第一个表格后创建空段落作为插入点");
                            // 找到表格在body中的位置
                            org.apache.xmlbeans.XmlCursor existingTableCursor = existingTable.newCursor();
                            existingTableCursor.toEndToken();
                            existingTableCursor.toNextToken();
                            
                            // 找到表格后的段落位置，但只查找属于同一主章节的段落
                            int insertPos = -1;
                            for (int k = 0; k < body2.sizeOfPArray(); k++) {
                                CTP checkPara7 = body2.getPArray(k);
                                org.apache.xmlbeans.XmlCursor checkCursor7 = checkPara7.newCursor();
                                if (checkCursor7.comparePosition(existingTableCursor) > 0) {
                                    // 检查这个段落是否属于同一主章节
                                    XWPFParagraph xwpfPara7 = new XWPFParagraph(checkPara7, document);
                                    String paraText7 = xwpfPara7.getText();
                                    
                                    boolean belongsToSameSection = false;
                                    if (paraText7 == null || paraText7.trim().isEmpty()) {
                                        // 空段落，属于同一章节
                                        belongsToSameSection = true;
                                    } else if (paraText7.trim().matches("^\\d+\\.\\d+\\s+.*")) {
                                        // 主章节标题，检查是否属于同一主章节
                                        String[] parts = paraText7.trim().split("\\s+", 2);
                                        if (parts.length > 0) {
                                            String mainSectionNum = parts[0];
                                            // 如果遇到其他主章节，停止查找
                                            if (!mainSectionNum.equals(sectionNumber)) {
                                                checkCursor7.close();
                                                break;
                                            }
                                            belongsToSameSection = true;
                                        }
                                    } else if (paraText7.trim().matches("^\\d+\\.\\d+\\.\\d+\\s+.*")) {
                                        // 子章节标题，检查是否属于同一主章节
                                        String[] parts = paraText7.trim().split("\\s+", 2);
                                        if (parts.length > 0) {
                                            String subSectionNum = parts[0];
                                            belongsToSameSection = subSectionNum.startsWith(sectionNumber + ".");
                                        }
                                    } else {
                                        // 普通段落，属于同一章节
                                        belongsToSameSection = true;
                                    }
                                    
                                    if (belongsToSameSection) {
                                        insertPos = k;
                                        checkCursor7.close();
                                        break;
                                    } else {
                                        // 遇到其他主章节，停止查找
                                        checkCursor7.close();
                                        break;
                                    }
                                }
                                checkCursor7.close();
                            }
                            existingTableCursor.close();
                            
                            if (insertPos >= 0) {
                                CTP newPara3 = body2.insertNewP(insertPos);
                                insertAfterPara = new XWPFParagraph(newPara3, document);
                                System.out.println("在第一个表格后创建空段落作为插入点（位置: " + insertPos + "）");
                            } else {
                                // 如果找不到，在第一个子章节后创建空段落
                                if (firstSubSectionPara != null) {
                                    int firstSubSectionIndex = -1;
                                    for (int k = 0; k < body2.sizeOfPArray(); k++) {
                                        if (body2.getPArray(k) == firstSubSectionPara.getCTP()) {
                                            firstSubSectionIndex = k;
                                            break;
                                        }
                                    }
                                    if (firstSubSectionIndex >= 0) {
                                        // 找到第一个子章节后的表格位置
                                        CTTbl firstTable2 = findTableCttblAfterParagraph(body, firstSubSectionPara.getCTP());
                                        if (firstTable2 != null) {
                                            // 找到表格后的位置
                                            org.apache.xmlbeans.XmlCursor firstTableCursor2 = firstTable2.newCursor();
                                            firstTableCursor2.toEndToken();
                                            firstTableCursor2.toNextToken();
                                            
                                            for (int k = firstSubSectionIndex + 1; k < body2.sizeOfPArray(); k++) {
                                                CTP checkPara8 = body2.getPArray(k);
                                                org.apache.xmlbeans.XmlCursor checkCursor8 = checkPara8.newCursor();
                                                if (checkCursor8.comparePosition(firstTableCursor2) > 0) {
                                                    CTP newPara3 = body2.insertNewP(k);
                                                    insertAfterPara = new XWPFParagraph(newPara3, document);
                                                    checkCursor8.close();
                                                    firstTableCursor2.close();
                                                    System.out.println("在第一个表格后创建空段落作为插入点（位置: " + k + "）");
                                                    break;
                                                }
                                                checkCursor8.close();
                                            }
                                            firstTableCursor2.close();
                                        }
                                    }
                                }
                                
                                if (insertAfterPara == null) {
                                    CTP newPara3 = body2.addNewP();
                                    insertAfterPara = new XWPFParagraph(newPara3, document);
                                    System.out.println("在body末尾创建空段落作为插入点");
                                }
                            }
                        }
                        
                        // 如果还是找不到，使用第一个子章节（作为最后的备选，但这样会导致顺序问题）
                        if (insertAfterPara == null) {
                            System.err.println("警告：找不到合适的插入点，使用第一个子章节，可能导致顺序问题");
                            insertAfterPara = firstSubSectionPara != null ? firstSubSectionPara : sectionPara;
                        }
                        
                        System.out.println("模块" + sectionNumber + "插入点: " + (insertAfterPara != null ? (insertAfterPara.getText() != null && !insertAfterPara.getText().trim().isEmpty() ? insertAfterPara.getText() : "空段落") : "null"));
                        
                        // 为剩余的测试用例创建子章节和表格（按顺序）
                        for (int i = 0; i < remainingTestCases.size(); i++) {
                            TestCase testCase = remainingTestCases.get(i);
                            int sequenceNumber = i + 2; // 从2开始编号
                            
                            // 创建子章节标题段落
                            String subSectionNumber = sectionNumber + "." + sequenceNumber;
                            System.out.println("创建子章节: " + subSectionNumber + " " + testCase.getTestName() + "测试");
                            
                            XWPFParagraph subSectionPara = createSubSectionParagraph(
                                document, insertAfterPara, subSectionNumber, testCase.getTestName());
                            
                            // 在子章节后查找并填充表格（优先使用已存在的表格，如果不存在则复制模板）
                            System.out.println("为子章节" + subSectionNumber + "查找并填充表格");
                            insertTestCaseTableAfterParagraph(document, subSectionPara.getCTP(), testCase, templateTable);
                            
                            // 创建完表格后，立即在表格后创建一个空段落作为下次的插入点
                            // 这样可以确保下一个子章节插入在当前子章节的表格之后
                            org.apache.xmlbeans.XmlCursor subSectionCursor = subSectionPara.getCTP().newCursor();
                            subSectionCursor.toEndToken();
                            
                            // 找到子章节后的表格
                            CTBody bodyAfterTable = document.getDocument().getBody();
                            CTTbl tableAfterSub = findTableCttblAfterParagraph(bodyAfterTable, subSectionPara.getCTP());
                            if (tableAfterSub != null) {
                                // 在表格后创建空段落
                                org.apache.xmlbeans.XmlCursor tableCursor = tableAfterSub.newCursor();
                                tableCursor.toEndToken();
                                tableCursor.toNextToken();
                                
                                tableCursor.beginElement(
                                    new javax.xml.namespace.QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p", "w"));
                                tableCursor.toParent();
                                
                                // 获取刚插入的CTP对象
                                CTP newEmptyPara = null;
                                if (tableCursor.getObject() instanceof CTP) {
                                    newEmptyPara = (CTP) tableCursor.getObject();
                                }
                                
                                tableCursor.close();
                                
                                if (newEmptyPara != null) {
                                    insertAfterPara = new XWPFParagraph(newEmptyPara, document);
                                    System.out.println("在子章节" + subSectionNumber + "的表格后创建空段落作为下次插入点");
                                } else {
                                    System.err.println("创建空段落失败");
                                }
                            } else {
                                System.err.println("找不到子章节" + subSectionNumber + "后的表格");
                            }
                            
                            subSectionCursor.close();
                        }
                        System.out.println("模块" + sectionNumber + "剩余" + remainingTestCases.size() + "个表格创建完成");
                    }
                } else {
                    // 如果没有表格，在章节后插入内容（为所有测试用例创建子章节和表格）
                    insertModuleContent(document, sectionPara, sectionNumber, moduleData, globalTemplateTable);
                    System.out.println("模块" + sectionNumber + "处理完成（生成" + 
                                     moduleData.getTestCaseCount() + "个表格）");
                }
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
     *
     * @param document Word文档
     * @return 子章节信息列表
     */
    private List<SubSectionInfo> scanExistingSubSections(XWPFDocument document) {
        List<SubSectionInfo> subSections = new java.util.ArrayList<>();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }

                // 匹配子章节格式：X.X.X 模块名称（如 "5.3.1 登录功能测试"）
                Matcher matcher = SECTION_PATTERN.matcher(text.trim());
                if (matcher.matches()) {
                    String sectionNumber = matcher.group(1);
                    // 只添加子章节（X.X.X格式），不添加主章节（X.X格式）
                    if (sectionNumber.matches("^\\d+\\.\\d+\\.\\d+$")) {
                        subSections.add(new SubSectionInfo(sectionNumber, para));
                    }
                }
            } catch (org.apache.xmlbeans.impl.values.XmlValueDisconnectedException e) {
                // 段落已被删除，跳过
                continue;
            } catch (Exception e) {
                // 其他异常，跳过
                continue;
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
     *
     * @param document Word文档
     * @param afterPara 在哪个段落后插入
     * @param moduleNumber 模块编号，如 "5.1"
     * @param testName 测试名称，如 "登录功能"
     * @return 创建的子章节段落
     */
    private XWPFParagraph createSubSectionParagraph(XWPFDocument document, XWPFParagraph afterPara, 
                                                     String moduleNumber, String testName) {
        CTBody body = document.getDocument().getBody();
        CTP afterCTP = afterPara.getCTP();
        
        // 使用XmlCursor在段落后直接插入新段落元素
        org.apache.xmlbeans.XmlCursor paraCursor = afterCTP.newCursor();
        paraCursor.toEndToken();
        paraCursor.toNextToken();
        
        // 在cursor位置插入新的段落元素
        paraCursor.beginElement(
            new javax.xml.namespace.QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p", "w"));
        paraCursor.toParent();
        
        // 获取刚插入的CTP对象
        CTP ctp = null;
        if (paraCursor.getObject() instanceof CTP) {
            ctp = (CTP) paraCursor.getObject();
        }
        
        paraCursor.close();
        
        if (ctp == null) {
            System.err.println("创建子章节段落失败");
            // 降级处理：使用老方法
            int pIndex = -1;
            for (int i = 0; i < body.sizeOfPArray(); i++) {
                if (body.getPArray(i) == afterCTP) {
                    pIndex = i;
                    break;
                }
            }
            if (pIndex == -1) {
                pIndex = body.sizeOfPArray();
            }
            ctp = body.insertNewP(pIndex + 1);
        }
        
        XWPFParagraph para = new XWPFParagraph(ctp, document);

        // 设置标题样式：小四、加粗
        String subTitle = moduleNumber + " " + testName + "测试";
        XWPFRun run = para.createRun();
        run.setText(subTitle);
        run.setBold(true);
        run.setFontSize(12); // 小四号字体

        // 设置段落格式
        para.setAlignment(ParagraphAlignment.LEFT);
        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(120)); // 段后间距
        
        return para;
    }

    /**
     * 扫描Word文档中的所有章节编号
     *
     * @param document Word文档
     * @return 章节编号列表，如 ["4.3", "5.3", "6.2"]
     */
    private List<String> scanWordSections(XWPFDocument document) {
        List<String> sectionNumbers = new java.util.ArrayList<>();
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }

                // 匹配章节标题格式：X.X 模块名称
                Matcher matcher = SECTION_PATTERN.matcher(text.trim());
                if (matcher.matches()) {
                    String sectionNumber = matcher.group(1);
                    // 只添加主章节（X.X格式），不添加子章节（X.X.X格式）
                    if (sectionNumber.matches("^\\d+\\.\\d+$") && !sectionNumber.matches("^\\d+\\.\\d+\\..+$")) {
                        if (!sectionNumbers.contains(sectionNumber)) {
                            sectionNumbers.add(sectionNumber);
                        }
                    }
                }
            } catch (org.apache.xmlbeans.impl.values.XmlValueDisconnectedException e) {
                // 段落已被删除，跳过
                continue;
            } catch (Exception e) {
                // 其他异常，跳过
                continue;
            }
        }
        
        return sectionNumbers;
    }

    /**
     * 查找章节标题段落
     *
     * @param document     Word文档
     * @param moduleNumber 模块编号，如"5.3"或"4.3"
     * @return 段落对象，未找到返回null
     */
    private XWPFParagraph findSectionParagraph(XWPFDocument document, String moduleNumber) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph para : paragraphs) {
            try {
                String text = para.getText();
                
                if (text == null || text.trim().isEmpty()) {
                    continue;
                }

                // 匹配章节标题格式：支持多种格式
                // 1. "5.3 功能测试" 或 "4.3 功能测试"
                // 2. "5.3.1 登录功能测试" (子章节，但也会匹配)
                // 3. 去掉所有空格后匹配
                String trimmedText = text.trim();
                
                // 先尝试标准格式：X.X 模块名称
                Matcher matcher = SECTION_PATTERN.matcher(trimmedText);
                if (matcher.matches()) {
                    String foundModuleNumber = matcher.group(1);
                    if (moduleNumber.equals(foundModuleNumber)) {
                        return para;
                    }
                }
                
                // 如果标准格式不匹配，尝试更宽松的匹配
                // 检查是否以模块编号开头（允许后面有任意字符）
                if (trimmedText.startsWith(moduleNumber + " ") || 
                    trimmedText.startsWith(moduleNumber + ".") ||
                    trimmedText.equals(moduleNumber)) {
                    return para;
                }
            } catch (org.apache.xmlbeans.impl.values.XmlValueDisconnectedException e) {
                // 段落已被删除，跳过
                continue;
            } catch (Exception e) {
                // 其他异常，跳过
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
