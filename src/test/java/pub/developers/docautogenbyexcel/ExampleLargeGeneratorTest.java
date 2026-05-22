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

import java.io.FileOutputStream;
import java.nio.file.Path;

/**
 * 生成大型示例 Excel/Word 并运行 TestRunner，供压力/集成测试使用
 */
public class ExampleLargeGeneratorTest {

    @Test
    public void generateLargeExampleAndRun() throws Exception {
        String baseDir = System.getProperty("user.dir");
        Path excelPath = Path.of(baseDir, "2 号文档完整填充模板.xlsx");
        Path wordPath = Path.of(baseDir, "2-XX 软件配置项测试报告 (公开）.docx");

        // 生成大型 Excel：模块 1..15，每模块 8 个测试用例
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("测试用例");
            Row header = sheet.createRow(0);
            String[] headers = new String[]{"模块编号", "测试用例标识", "测试项名称", "标识", "需求ID", "测试内容", "测试策略与方法", "判定准则"};
            for (int i = 0; i < headers.length; i++) header.createCell(i).setCellValue(headers[i]);

            int rowIdx = 1;
            for (int m = 1; m <= 15; m++) {
                String module = String.valueOf(m);
                for (int t = 1; t <= 8; t++) {
                    Row r = sheet.createRow(rowIdx++);
                    r.createCell(0).setCellValue(module);
                    r.createCell(1).setCellValue("TC-" + module + String.format("-%03d", t));
                    r.createCell(2).setCellValue("功能测试-模块" + module + "-用例" + t);
                    r.createCell(3).setCellValue("ID-" + module + "-" + t);
                    r.createCell(4).setCellValue("REQ-" + module);
                    r.createCell(5).setCellValue("步骤描述：验证模块" + module + "的第" + t + "项行为");
                    r.createCell(6).setCellValue("自动化");
                    r.createCell(7).setCellValue("通过/失败");
                }
            }

            // 测试步骤Sheet：为每个用例添加3个步骤
            XSSFSheet steps = wb.createSheet("测试步骤");
            Row sh = steps.createRow(0);
            sh.createCell(0).setCellValue("测试用例标识");
            sh.createCell(1).setCellValue("步骤序号");
            sh.createCell(2).setCellValue("测试步骤");
            sh.createCell(3).setCellValue("预期结果");

            int sIdx = 1;
            for (int m = 1; m <= 15; m++) {
                for (int t = 1; t <= 8; t++) {
                    String caseId = "TC-" + m + String.format("-%03d", t);
                    for (int s = 1; s <= 3; s++) {
                        Row r = steps.createRow(sIdx++);
                        r.createCell(0).setCellValue(caseId);
                        r.createCell(1).setCellValue(s);
                        r.createCell(2).setCellValue("执行操作 " + s + " (模块" + m + ")");
                        r.createCell(3).setCellValue("期望结果 " + s);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(excelPath.toFile())) {
                wb.write(fos);
            }
        }

        // 生成大型 Word 模板：包含 TOC 条目 + 正文标题，每个主模块至少一个子章节模板
        try (XWPFDocument doc = new XWPFDocument()) {
            // TOC 风格条目（让 WordProcessor 能扫描到章节编号）
            for (int m = 1; m <= 15; m++) {
                XWPFParagraph toc = doc.createParagraph();
                try { toc.setStyle(m <= 3 ? "22" : "16"); } catch (Exception ignored) {}
                XWPFRun run = toc.createRun();
                run.setText(m + " 模块说明");
            }

            // 正文：每个模块一个主标题和一个子标题（子标题可作为模板）
            for (int m = 1; m <= 15; m++) {
                XWPFParagraph main = doc.createParagraph();
                main.setStyle("3");
                XWPFRun rm = main.createRun();
                rm.setText(m + " 模块说明");

                // 子章节模板
                XWPFParagraph sub = doc.createParagraph();
                sub.setStyle("4");
                XWPFRun rs = sub.createRun();
                rs.setText(m + ".1 示例子章节");

                // 在第1个模块下添加一个模板表格，供后续模块复制
                if (m == 1) {
                    XWPFTable table = doc.createTable();
                    XWPFTableRow r0 = table.getRow(0);
                    r0.getCell(0).setText("测试用例标识");
                    r0.addNewTableCell().setText("");
                    r0.addNewTableCell().setText("");

                    XWPFTableRow r1 = table.createRow();
                    r1.getCell(0).setText("测试项名称"); r1.getCell(1).setText(""); r1.getCell(2).setText("");
                    XWPFTableRow r2 = table.createRow();
                    r2.getCell(0).setText("测试内容"); r2.getCell(1).setText(""); r2.getCell(2).setText("");
                    XWPFTableRow r3 = table.createRow();
                    r3.getCell(0).setText("测试步骤"); r3.getCell(1).setText(""); r3.getCell(2).setText("");
                    XWPFTableRow r4 = table.createRow();
                    r4.getCell(0).setText("序号"); r4.getCell(1).setText("操作"); r4.getCell(2).setText("预期结果");
                    XWPFTableRow r5 = table.createRow();
                    r5.getCell(0).setText("1"); r5.getCell(1).setText(""); r5.getCell(2).setText("");
                }
            }

            // 保存 Word
            try (FileOutputStream fos = new FileOutputStream(wordPath.toFile())) {
                doc.write(fos);
            }
        }

        // 运行 TestRunner
        System.out.println("生成大型示例文件，调用 TestRunner 进行填充...");
        TestRunner.main(new String[0]);
    }
}
