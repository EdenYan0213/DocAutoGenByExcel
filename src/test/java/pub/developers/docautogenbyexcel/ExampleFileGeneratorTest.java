package pub.developers.docautogenbyexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Path;

public class ExampleFileGeneratorTest {

    @Test
    public void generateExampleFilesAndRunTool() throws Exception {
        String baseDir = System.getProperty("user.dir");
        Path excelPath = Path.of(baseDir, "2 号文档完整填充模板.xlsx");
        Path wordPath = Path.of(baseDir, "2-XX 软件配置项测试报告 (公开）.docx");

        // 1) 生成 Excel 文件（测试用例 + 测试步骤 + 基本信息示例）
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("测试用例");
            Row header = sheet.createRow(0);
            String[] headers = new String[]{"模块编号", "测试用例标识", "测试项名称", "标识", "需求ID", "测试内容", "测试策略与方法", "判定准则"};
            for (int i = 0; i < headers.length; i++) {
                Cell c = header.createCell(i);
                c.setCellValue(headers[i]);
            }

            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("5");
            r1.createCell(1).setCellValue("TC-5-001");
            r1.createCell(2).setCellValue("配置项校验");
            r1.createCell(3).setCellValue("ID-5-001");
            r1.createCell(4).setCellValue("REQ-5");
            r1.createCell(5).setCellValue("校验配置项是否存在");
            r1.createCell(6).setCellValue("手动");
            r1.createCell(7).setCellValue("通过/失败");

            // 测试步骤Sheet
            XSSFSheet steps = wb.createSheet("测试步骤");
            Row sh = steps.createRow(0);
            sh.createCell(0).setCellValue("测试用例标识");
            sh.createCell(1).setCellValue("步骤序号");
            sh.createCell(2).setCellValue("测试步骤");
            sh.createCell(3).setCellValue("预期结果");

            Row s1 = steps.createRow(1);
            s1.createCell(0).setCellValue("TC-5-001");
            s1.createCell(1).setCellValue(1);
            s1.createCell(2).setCellValue("检查配置项是否存在");
            s1.createCell(3).setCellValue("配置项存在");

            // 保存 Excel
            try (FileOutputStream fos = new FileOutputStream(excelPath.toFile())) {
                wb.write(fos);
            }
        }

        // 2) 生成 Word 模板（包含主章节、子章节和模板表格）
        try (XWPFDocument doc = new XWPFDocument()) {
            // 为了让 WordProcessor 能识别章节编号，先在文档中加入目录样式（TOC）条目
            XWPFParagraph tocParaMain = doc.createParagraph();
            try { tocParaMain.setStyle("22"); } catch (Exception ignored) {}
            XWPFRun tocRunMain = tocParaMain.createRun();
            tocRunMain.setText("5 模块配置");

            XWPFParagraph tocParaSub = doc.createParagraph();
            try { tocParaSub.setStyle("16"); } catch (Exception ignored) {}
            XWPFRun tocRunSub = tocParaSub.createRun();
            tocRunSub.setText("5.1 配置项测试");

            // 主章节（设为Heading2样式）
            XWPFParagraph pMain = doc.createParagraph();
            pMain.setStyle("3");
            XWPFRun rMain = pMain.createRun();
            rMain.setText("5 模块配置");
            rMain.setBold(true);

            // 子章节（模板子章节，Heading3/4）
            XWPFParagraph pSub = doc.createParagraph();
            pSub.setStyle("4");
            XWPFRun rSub = pSub.createRun();
            rSub.setText("5.1 配置项测试");
            rSub.setBold(true);

            // 模板表格：包含标签/数据格式与测试步骤子表格
            XWPFTable table = doc.createTable();
            // 第1行：测试用例标识
            XWPFTableRow row0 = table.getRow(0);
            row0.getCell(0).setText("测试用例标识");
            row0.addNewTableCell().setText("");
            row0.addNewTableCell().setText("");

            // 行：测试项名称
            XWPFTableRow row1 = table.createRow();
            row1.getCell(0).setText("测试项名称");
            row1.getCell(1).setText("");
            row1.getCell(2).setText("");

            // 行：测试内容
            XWPFTableRow row2 = table.createRow();
            row2.getCell(0).setText("测试内容");
            row2.getCell(1).setText("");
            row2.getCell(2).setText("");

            // 行：测试策略与方法
            XWPFTableRow row3 = table.createRow();
            row3.getCell(0).setText("测试策略与方法");
            row3.getCell(1).setText("");
            row3.getCell(2).setText("");

            // 行：判定准则
            XWPFTableRow row4 = table.createRow();
            row4.getCell(0).setText("判定准则");
            row4.getCell(1).setText("");
            row4.getCell(2).setText("");

            // 子表格标题行：测试步骤
            XWPFTableRow row5 = table.createRow();
            row5.getCell(0).setText("测试步骤");
            row5.getCell(1).setText("");
            row5.getCell(2).setText("");

            // 子表格列头
            XWPFTableRow row6 = table.createRow();
            row6.getCell(0).setText("序号");
            row6.getCell(1).setText("操作");
            row6.getCell(2).setText("预期结果");

            // 样例步骤行
            XWPFTableRow row7 = table.createRow();
            row7.getCell(0).setText("1");
            row7.getCell(1).setText("");
            row7.getCell(2).setText("");

            // 居中Caption样式示例
            XWPFParagraph caption = doc.createParagraph();
            caption.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun capRun = caption.createRun();
            capRun.setText("表5.1 配置项测试");

            // 保存 Word
            try (FileOutputStream fos = new FileOutputStream(wordPath.toFile())) {
                doc.write(fos);
            }
        }

        // 3) 调用 TestRunner 运行工具流程
        System.out.println("Generated files: " + excelPath + " , " + wordPath);
        long runStartTime = System.currentTimeMillis();
        // 调用 test runner main
        TestRunner.main(new String[0]);

        // 验证输出文件存在（兼容 test_result_时间戳.docx 和 test_result.docx）
        File outDir = Path.of(baseDir, "test_output").toFile();
        File[] generatedFiles = outDir.listFiles((dir, name) ->
                (name.startsWith("test_result_") && name.endsWith(".docx")) || "test_result.docx".equals(name));

        File latestFile = null;
        if (generatedFiles != null) {
            for (File file : generatedFiles) {
                if (file.lastModified() >= runStartTime && (latestFile == null || file.lastModified() > latestFile.lastModified())) {
                    latestFile = file;
                }
            }
        }

        if (latestFile == null || !latestFile.exists() || latestFile.length() == 0) {
            throw new Exception("工具未生成输出文件或文件为空: " + Path.of(baseDir, "test_output").toAbsolutePath());
        }
    }
}
