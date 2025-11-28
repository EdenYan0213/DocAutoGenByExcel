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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Word模板处理模块
 * 负责读取Word模板，定位章节，插入表格和子标题
 */
public class WordProcessor {
    
    // 章节标题匹配模式：X.X 模块名称（X为数字）
    private static final Pattern SECTION_PATTERN = Pattern.compile("^(\\d+\\.\\d+)\\s+(.+)$");

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

            // 查找并处理每个模块
            int successCount = 0;
            for (java.util.Map.Entry<String, ModuleData> entry : moduleDataMap.entrySet()) {
                String moduleNumber = entry.getKey();
                ModuleData moduleData = entry.getValue();

                // 查找章节标题位置
                XWPFParagraph sectionPara = findSectionParagraph(document, moduleNumber);
                if (sectionPara == null) {
                    System.out.println("Word模板中未找到模块: " + moduleNumber);
                    continue;
                }

                // 在章节后插入内容
                insertModuleContent(document, sectionPara, moduleNumber, moduleData);
                successCount++;
                System.out.println("模块" + moduleNumber + "处理完成（生成" + 
                                 moduleData.getTestCaseCount() + "个表格）");
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
     * 查找章节标题段落
     *
     * @param document     Word文档
     * @param moduleNumber 模块编号，如"5.3"
     * @return 段落对象，未找到返回null
     */
    private XWPFParagraph findSectionParagraph(XWPFDocument document, String moduleNumber) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph para : paragraphs) {
            String text = para.getText();
            
            if (text == null || text.trim().isEmpty()) {
                continue;
            }

            // 匹配章节标题格式
            Matcher matcher = SECTION_PATTERN.matcher(text.trim());
            if (matcher.matches()) {
                String foundModuleNumber = matcher.group(1);
                if (moduleNumber.equals(foundModuleNumber)) {
                    return para;
                }
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
                                     String moduleNumber, ModuleData moduleData) {
        List<TestCase> testCases = moduleData.getTestCases();

        if (testCases.isEmpty()) {
            return;
        }

        // 获取body和sectionPara的CTP
        CTBody body = document.getDocument().getBody();
        CTP sectionCTP = sectionPara.getCTP();
        
        // 找到sectionCTP在body的P数组中的索引
        int pIndex = -1;
        for (int i = 0; i < body.sizeOfPArray(); i++) {
            if (body.getPArray(i) == sectionCTP) {
                pIndex = i;
                break;
            }
        }

        if (pIndex == -1) {
            return;
        }

        // 倒序插入内容，确保位置正确（每个测试用例：子标题 + 表格 + 空行）
        for (int i = testCases.size() - 1; i >= 0; i--) {
            TestCase testCase = testCases.get(i);
            int sequenceNumber = i + 1;
            int insertPos = pIndex + 1; // 在段落后插入

            // 1. 先插入子标题
            String subTitle = moduleNumber + "." + sequenceNumber + " " + testCase.getTestName() + "测试";
            CTP subTitlePara = insertSubTitleAt(document, body, insertPos, subTitle);

            // 2. 然后在子标题后插入表格
            insertTestCaseTableAfterParagraph(document, subTitlePara, testCase);

            // 3. 如果不是最后一个，插入空行
            if (i < testCases.size() - 1) {
                body.insertNewP(insertPos);
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
     * 在段落后插入测试用例表格
     */
    private void insertTestCaseTableAfterParagraph(XWPFDocument document, CTP paragraph, TestCase testCase) {
        // 在 body 末尾创建表格
        CTBody body = document.getDocument().getBody();
        CTTbl cttbl = body.addNewTbl();
        
        // 创建XWPFTable对象
        XWPFTable table = new XWPFTable(cttbl, document);
        
        // 先创建6行4列的表格结构
        for (int i = 0; i < 6; i++) {
            XWPFTableRow row = table.createRow();
            // 确保每行有4列
            while (row.getTableCells().size() < 4) {
                row.createCell();
            }
        }
        
        // 填充表格数据
        fillTableData(table, testCase);
        
        // 设置表格样式
        styleTable(cttbl);
        
        // 将表格移动到段落后
        org.apache.xmlbeans.XmlCursor tableCursor = cttbl.newCursor();
        org.apache.xmlbeans.XmlCursor paraCursor = paragraph.newCursor();
        
        // 移动到段落结束后
        paraCursor.toEndToken();
        paraCursor.toNextToken();
        
        // 将表格移动到段落后
        tableCursor.moveXml(paraCursor);
        
        paraCursor.close();
        tableCursor.close();
    }
    
    /**
     * 填充表格数据
     */
    private void fillTableData(XWPFTable table, TestCase testCase) {
        // 设置表格宽度为100%
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr == null) {
            tblPr = table.getCTTbl().addNewTblPr();
        }
        CTTblWidth tblWidth = tblPr.getTblW();
        if (tblWidth == null) {
            tblWidth = tblPr.addNewTblW();
        }
        tblWidth.setType(STTblWidth.DXA);
        tblWidth.setW(BigInteger.valueOf(9072)); // 页面宽度

        // 填充表格数据
        // 第1行：测试项名称、testName、标识、id
        setCellValue(table, 0, 0, "测试项名称");
        setCellValue(table, 0, 1, testCase.getTestName());
        setCellValue(table, 0, 2, "标识");
        setCellValue(table, 0, 3, testCase.getId());

        // 第2行：测试内容、content
        setCellValue(table, 1, 0, "测试内容");
        setCellValue(table, 1, 1, testCase.getContent());
        setCellValue(table, 1, 2, "");
        setCellValue(table, 1, 3, "");

        // 第3行：测试策略与方法、strategy
        setCellValue(table, 2, 0, "测试策略与方法");
        setCellValue(table, 2, 1, testCase.getStrategy());
        setCellValue(table, 2, 2, "");
        setCellValue(table, 2, 3, "");

        // 第4行：判定准则、criteria
        setCellValue(table, 3, 0, "判定准则");
        setCellValue(table, 3, 1, testCase.getCriteria());
        setCellValue(table, 3, 2, "");
        setCellValue(table, 3, 3, "");

        // 第5行：测试终止条件、stopCondition
        setCellValue(table, 4, 0, "测试终止条件");
        setCellValue(table, 4, 1, testCase.getStopCondition());
        setCellValue(table, 4, 2, "");
        setCellValue(table, 4, 3, "");

        // 第6行：追踪关系、trace
        setCellValue(table, 5, 0, "追踪关系");
        setCellValue(table, 5, 1, testCase.getTrace());
        setCellValue(table, 5, 2, "");
        setCellValue(table, 5, 3, "");
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
            if (cell.getParagraphs().size() > 0) {
                cell.removeParagraph(0);
            }
        } catch (Exception e) {
            // 忽略错误，继续添加新段落
        }
        
        XWPFParagraph para = cell.addParagraph();
        XWPFRun run = para.createRun();
        run.setText(value != null ? value : "");
        
        // 设置单元格对齐方式：垂直居中、左对齐
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        para.setAlignment(ParagraphAlignment.LEFT);
        
        // 设置字体大小：五号
        run.setFontSize(10);
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
