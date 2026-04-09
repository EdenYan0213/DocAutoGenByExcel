package pub.developers.docautogenbyexcel;

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
 * 生成多变体模板（不同表格结构）和对应 Excel 数据，并运行 TestRunner
 */
public class ExampleVariantsGeneratorTest {

    @Test
    public void generateVariantsAndRun() throws Exception {
        String baseDir = System.getProperty("user.dir");
        String ts = String.valueOf(System.currentTimeMillis());
        Path excelPath = Path.of(baseDir, "2 号文档完整填充模板_" + ts + ".xlsx");
        Path wordPath = Path.of(baseDir, "2-XX 软件配置项测试报告 (公开）_" + ts + ".docx");

        // 1. 生成 Excel：测试用例、基本信息、列表型表格（多种表结构）
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            // 测试用例Sheet（模块2..4，每模块3用例）
            XSSFSheet cases = wb.createSheet("测试用例");
            Row h = cases.createRow(0);
            String[] headers = new String[]{"模块编号", "测试用例标识", "测试项名称", "标识", "需求ID", "测试内容", "测试策略与方法", "判定准则"};
            for (int i = 0; i < headers.length; i++) h.createCell(i).setCellValue(headers[i]);
            int rid = 1;
            for (int m = 2; m <= 4; m++) {
                for (int t = 1; t <= 3; t++) {
                    Row r = cases.createRow(rid++);
                    r.createCell(0).setCellValue(String.valueOf(m));
                    r.createCell(1).setCellValue("TC-" + m + "-" + t);
                    r.createCell(2).setCellValue("模块" + m + " 功能测试" + t);
                    r.createCell(3).setCellValue("ID-" + m + "-" + t);
                    r.createCell(4).setCellValue("REQ-" + m);
                    r.createCell(5).setCellValue("步骤描述：模块" + m + " 的用例" + t);
                    r.createCell(6).setCellValue(t % 2 == 0 ? "自动化" : "手动");
                    r.createCell(7).setCellValue("通过/失败");
                }
            }

            // 测试步骤Sheet
            XSSFSheet steps = wb.createSheet("测试步骤");
            Row sh = steps.createRow(0);
            sh.createCell(0).setCellValue("测试用例标识");
            sh.createCell(1).setCellValue("步骤序号");
            sh.createCell(2).setCellValue("测试步骤");
            sh.createCell(3).setCellValue("预期结果");
            int sid = 1;
            for (int m = 2; m <= 4; m++) {
                for (int t = 1; t <= 3; t++) {
                    String cid = "TC-" + m + "-" + t;
                    for (int s = 1; s <= 2; s++) {
                        Row r = steps.createRow(sid++);
                        r.createCell(0).setCellValue(cid);
                        r.createCell(1).setCellValue(s);
                        r.createCell(2).setCellValue("操作步骤 " + s);
                        r.createCell(3).setCellValue("期望 " + s);
                    }
                }
            }

            // 基本信息Sheet（表格名称、字段名、字段值）
            XSSFSheet basic = wb.createSheet("基本信息");
            Row bh = basic.createRow(0);
            bh.createCell(0).setCellValue("表格名称");
            bh.createCell(1).setCellValue("字段名");
            bh.createCell(2).setCellValue("字段值");
            Row b1 = basic.createRow(1);
            b1.createCell(0).setCellValue("项目基本信息");
            b1.createCell(1).setCellValue("项目名称");
            b1.createCell(2).setCellValue("示例项目A");
            Row b2 = basic.createRow(2);
            b2.createCell(0).setCellValue("项目基本信息");
            b2.createCell(1).setCellValue("版本");
            b2.createCell(2).setCellValue("v1.2.3");

            // 列表型表格（第一列为表格名称）
            XSSFSheet list = wb.createSheet("列表型表格");
            Row lh = list.createRow(0);
            lh.createCell(0).setCellValue("表格名称");
            lh.createCell(1).setCellValue("项名");
            lh.createCell(2).setCellValue("说明");
            Row l1 = list.createRow(1);
            l1.createCell(0).setCellValue("配置清单"); l1.createCell(1).setCellValue("cfgA"); l1.createCell(2).setCellValue("配置A说明");
            Row l2 = list.createRow(2);
            l2.createCell(0).setCellValue("配置清单"); l2.createCell(1).setCellValue("cfgB"); l2.createCell(2).setCellValue("配置B说明");

            try (FileOutputStream fos = new FileOutputStream(excelPath.toFile())) {
                wb.write(fos);
            }
        }

        // 2. 生成 Word 模板：包含多种表格结构（KV表、4列标签表、列表型表格、测试用例模板表）
        try (XWPFDocument doc = new XWPFDocument()) {
            // TOC 条目（为多个模块创建，使后续能够识别并填充多个模块）
            XWPFParagraph toc1 = doc.createParagraph();
            try { toc1.setStyle("22"); } catch (Exception ignored) {}
            XWPFRun tr1 = toc1.createRun(); tr1.setText("2 模块总览");
            XWPFParagraph toc1b = doc.createParagraph();
            try { toc1b.setStyle("22"); } catch (Exception ignored) {}
            XWPFRun tr1b = toc1b.createRun(); tr1b.setText("3 模块总览");
            XWPFParagraph toc1c = doc.createParagraph();
            try { toc1c.setStyle("22"); } catch (Exception ignored) {}
            XWPFRun tr1c = toc1c.createRun(); tr1c.setText("4 模块总览");

            XWPFParagraph toc2 = doc.createParagraph();
            try { toc2.setStyle("16"); } catch (Exception ignored) {}
            XWPFRun tr2 = toc2.createRun(); tr2.setText("2.1 子章节示例");

            // 主标题（模块2/3/4）
            XWPFParagraph hMain = doc.createParagraph(); hMain.setStyle("3"); hMain.createRun().setText("2 模块总览");
            XWPFParagraph hMain3 = doc.createParagraph(); hMain3.setStyle("3"); hMain3.createRun().setText("3 模块总览");
            XWPFParagraph hMain4 = doc.createParagraph(); hMain4.setStyle("3"); hMain4.createRun().setText("4 模块总览");

            // 2列KV表（基本信息）
            XWPFParagraph pKv = doc.createParagraph();
            XWPFRun runKv = pKv.createRun();
            runKv.setBold(true);
            runKv.setText("基本信息表（KV）");
            XWPFTable kv = doc.createTable();
            XWPFTableRow kvr0 = kv.getRow(0); kvr0.getCell(0).setText("项目名称"); kvr0.addNewTableCell().setText("");
            XWPFTableRow kvr1 = kv.createRow(); kvr1.getCell(0).setText("版本"); kvr1.getCell(1).setText("");

            // 在KV表前插入Caption，便于 TableFillProcessor 匹配
            XWPFParagraph kvCaption = doc.createParagraph();
            try { kvCaption.setStyle("11"); } catch (Exception ignored) {}
            kvCaption.createRun().setText("表1 项目基本信息");

            // 4列标签-数据表
            XWPFParagraph p4 = doc.createParagraph();
            XWPFRun run4 = p4.createRun();
            run4.setBold(true);
            run4.setText("标签-数据表（2对）");
            XWPFTable t4 = doc.createTable();
            XWPFTableRow r0 = t4.getRow(0); r0.getCell(0).setText("标签1"); r0.addNewTableCell().setText(""); r0.addNewTableCell().setText("标签2"); r0.addNewTableCell().setText("");
            XWPFTableRow r1 = t4.createRow(); r1.getCell(0).setText("作者"); r1.getCell(1).setText(""); r1.getCell(2).setText("日期"); r1.getCell(3).setText("");

            // 列表型表格示例（表格名称在第一列）
            XWPFParagraph pList = doc.createParagraph();
            XWPFRun runList = pList.createRun();
            runList.setBold(true);
            runList.setText("列表型表格示例");
            XWPFTable listTbl = doc.createTable();
            XWPFTableRow lst0 = listTbl.getRow(0); lst0.getCell(0).setText("表格名称"); lst0.addNewTableCell().setText("项名"); lst0.addNewTableCell().setText("说明");
            XWPFTableRow lst1 = listTbl.createRow(); lst1.getCell(0).setText("配置清单"); lst1.getCell(1).setText("cfgA"); lst1.getCell(2).setText("");

            // 在列表表前插入Caption，便于匹配
            XWPFParagraph listCaption = doc.createParagraph();
            try { listCaption.setStyle("11"); } catch (Exception ignored) {}
            listCaption.createRun().setText("表2 配置清单");

            // 测试用例模板表（用于复制并填充测试用例和测试步骤）
            XWPFParagraph pCase = doc.createParagraph();
            XWPFRun runCase = pCase.createRun();
            runCase.setBold(true);
            runCase.setText("测试用例模板表");
            XWPFTable caseTbl = doc.createTable();
            XWPFTableRow c0 = caseTbl.getRow(0); c0.getCell(0).setText("测试用例标识"); c0.addNewTableCell().setText(""); c0.addNewTableCell().setText("");
            XWPFTableRow c1 = caseTbl.createRow(); c1.getCell(0).setText("测试项名称"); c1.getCell(1).setText(""); c1.getCell(2).setText("");
            XWPFTableRow c2 = caseTbl.createRow(); c2.getCell(0).setText("测试步骤"); c2.getCell(1).setText(""); c2.getCell(2).setText("");
            XWPFTableRow c3 = caseTbl.createRow(); c3.getCell(0).setText("序号"); c3.getCell(1).setText("操作"); c3.getCell(2).setText("预期结果");
            XWPFTableRow c4 = caseTbl.createRow(); c4.getCell(0).setText("1"); c4.getCell(1).setText(""); c4.getCell(2).setText("");

            // 保存 Word
            try (FileOutputStream fos = new FileOutputStream(wordPath.toFile())) {
                doc.write(fos);
            }
        }

        // 3. 运行 TestRunner
        System.out.println("生成多变体模板并运行 TestRunner...");
        TestRunner.main(new String[0]);
    }
}
