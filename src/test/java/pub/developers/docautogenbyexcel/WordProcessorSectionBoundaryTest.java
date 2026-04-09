package pub.developers.docautogenbyexcel;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.jupiter.api.Test;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.processor.WordProcessor;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class WordProcessorSectionBoundaryTest {

    @Test
    void generatedSubSectionsStayInsideParentSectionBoundary() throws Exception {
        Path tempDir = Files.createTempDirectory("word-boundary-test-");
        Path templatePath = tempDir.resolve("template.docx");
        Path outputPath = tempDir.resolve("output.docx");

        createTemplateWith52And61(templatePath);

        Map<String, ModuleData> moduleDataMap = new LinkedHashMap<>();
        ModuleData module52 = new ModuleData("5.2");
        module52.addTestCase(createTestCase("5.2", "文档上传和下载测试", "GNCS_001"));
        module52.addTestCase(createTestCase("5.2", "异常路径测试", "GNCS_002"));
        module52.addTestCase(createTestCase("5.2", "边界值测试", "GNCS_003"));
        module52.addTestCase(createTestCase("5.2", "回归测试", "GNCS_004"));
        moduleDataMap.put("5.2", module52);

        WordProcessor processor = new WordProcessor();
        processor.processWord(templatePath.toString(), outputPath.toString(), moduleDataMap);

        try (FileInputStream fis = new FileInputStream(outputPath.toFile());
             XWPFDocument outDoc = new XWPFDocument(fis)) {

            int indexOf61 = -1;
            int generated52Count = 0;

            for (int i = 0; i < outDoc.getParagraphs().size(); i++) {
                XWPFParagraph para = outDoc.getParagraphs().get(i);
                String text = para.getText() == null ? "" : para.getText().trim();
                String style = para.getStyle() == null ? "" : para.getStyle();

                if ("3".equals(style) && text.startsWith("6.1 ")) {
                    indexOf61 = i;
                }
                if ("4".equals(style) && text.matches("^5\\.2\\.\\d+\\s+.+")) {
                    generated52Count++;
                    if (indexOf61 != -1) {
                        assertTrue(i < indexOf61,
                            "5.2.x 子章节出现在 6.1 之后: " + text + " at index " + i + ", 6.1 at " + indexOf61);
                    }
                }
            }

            assertTrue(indexOf61 != -1, "未找到 6.1 主章节，测试模板构造失败");
            assertEquals(4, generated52Count, "应生成 4 个 5.2.x 子章节");
        }
    }

    @Test
    void generatedSubSectionsStayInsideAllModuleBoundaries() throws Exception {
        Path tempDir = Files.createTempDirectory("word-boundary-multi-");
        Path templatePath = tempDir.resolve("template-multi.docx");
        Path outputPath = tempDir.resolve("output-multi.docx");

        createTemplateWith523And54(templatePath);

        Map<String, ModuleData> moduleDataMap = new LinkedHashMap<>();

        ModuleData module52 = new ModuleData("5.2");
        module52.addTestCase(createTestCase("5.2", "模块52用例1", "M52_001"));
        module52.addTestCase(createTestCase("5.2", "模块52用例2", "M52_002"));
        module52.addTestCase(createTestCase("5.2", "模块52用例3", "M52_003"));
        moduleDataMap.put("5.2", module52);

        ModuleData module53 = new ModuleData("5.3");
        module53.addTestCase(createTestCase("5.3", "模块53用例1", "M53_001"));
        module53.addTestCase(createTestCase("5.3", "模块53用例2", "M53_002"));
        module53.addTestCase(createTestCase("5.3", "模块53用例3", "M53_003"));
        moduleDataMap.put("5.3", module53);

        WordProcessor processor = new WordProcessor();
        processor.processWord(templatePath.toString(), outputPath.toString(), moduleDataMap);

        try (FileInputStream fis = new FileInputStream(outputPath.toFile());
             XWPFDocument outDoc = new XWPFDocument(fis)) {

            int indexOf53 = -1;
            int indexOf54 = -1;
            int generated52Count = 0;
            int generated53Count = 0;

            for (int i = 0; i < outDoc.getParagraphs().size(); i++) {
                XWPFParagraph para = outDoc.getParagraphs().get(i);
                String text = para.getText() == null ? "" : para.getText().trim();
                String style = para.getStyle() == null ? "" : para.getStyle();

                if ("3".equals(style) && text.startsWith("5.3 ")) {
                    indexOf53 = i;
                }
                if ("3".equals(style) && text.startsWith("5.4 ")) {
                    indexOf54 = i;
                }

                if ("4".equals(style) && text.matches("^5\\.2\\.\\d+\\s+.+")) {
                    generated52Count++;
                    if (indexOf53 != -1) {
                        assertTrue(i < indexOf53,
                            "5.2.x 子章节越过 5.3 边界: " + text + " at index " + i + ", 5.3 at " + indexOf53);
                    }
                }

                if ("4".equals(style) && text.matches("^5\\.3\\.\\d+\\s+.+")) {
                    generated53Count++;
                    if (indexOf54 != -1) {
                        assertTrue(i < indexOf54,
                            "5.3.x 子章节越过 5.4 边界: " + text + " at index " + i + ", 5.4 at " + indexOf54);
                    }
                }
            }

            assertTrue(indexOf53 != -1, "未找到 5.3 主章节，测试模板构造失败");
            assertTrue(indexOf54 != -1, "未找到 5.4 主章节，测试模板构造失败");
            assertEquals(3, generated52Count, "应生成 3 个 5.2.x 子章节");
            assertEquals(3, generated53Count, "应生成 3 个 5.3.x 子章节");
        }
    }

    private static TestCase createTestCase(String moduleNumber, String testName, String id) {
        TestCase testCase = new TestCase(moduleNumber);
        testCase.addColumnData("模块编号", moduleNumber);
        testCase.addColumnData("测试项名称", testName);
        testCase.addColumnData("标识", id);
        testCase.addColumnData("测试内容", "内容");
        testCase.addColumnData("测试策略与方法", "策略");
        testCase.addColumnData("判定准则", "准则");
        testCase.addColumnData("测试终止条件", "终止条件");
        testCase.addColumnData("追踪关系", "追踪");
        return testCase;
    }

    private static void createTemplateWith52And61(Path templatePath) throws Exception {
        try (XWPFDocument doc = new XWPFDocument()) {
            // TOC-like entries so WordProcessor can scan sections from TOC paragraphs.
            XWPFParagraph toc52 = doc.createParagraph();
            toc52.setStyle("22");
            toc52.createRun().setText("5.2 功能测试");

            XWPFParagraph toc521 = doc.createParagraph();
            toc521.setStyle("16");
            toc521.createRun().setText("5.2.1 占位子章节");

            XWPFParagraph toc61 = doc.createParagraph();
            toc61.setStyle("22");
            toc61.createRun().setText("6.1 动态测试环境");

            XWPFParagraph section52 = doc.createParagraph();
            section52.setStyle("3");
            section52.createRun().setText("5.2 功能测试");

            XWPFParagraph subsection521 = doc.createParagraph();
            subsection521.setStyle("4");
            subsection521.createRun().setText("5.2.1 占位子章节");

            XWPFTable table = doc.createTable(6, 4);
            XWPFTableRow row0 = table.getRow(0);
            row0.getCell(0).setText("测试项名称");
            row0.getCell(2).setText("标识");

            table.getRow(1).getCell(0).setText("测试内容");
            table.getRow(2).getCell(0).setText("测试策略与方法");
            table.getRow(3).getCell(0).setText("判定准则");
            table.getRow(4).getCell(0).setText("测试终止条件");
            table.getRow(5).getCell(0).setText("追踪关系");

            XWPFParagraph section61 = doc.createParagraph();
            section61.setStyle("3");
            section61.createRun().setText("6.1 动态测试环境");

            try (FileOutputStream fos = new FileOutputStream(templatePath.toFile())) {
                doc.write(fos);
            }
        }
    }

    private static void createTemplateWith523And54(Path templatePath) throws Exception {
        try (XWPFDocument doc = new XWPFDocument()) {
            XWPFParagraph toc52 = doc.createParagraph();
            toc52.setStyle("22");
            toc52.createRun().setText("5.2 功能测试");

            XWPFParagraph toc521 = doc.createParagraph();
            toc521.setStyle("16");
            toc521.createRun().setText("5.2.1 占位子章节");

            XWPFParagraph toc53 = doc.createParagraph();
            toc53.setStyle("22");
            toc53.createRun().setText("5.3 性能测试");

            XWPFParagraph toc531 = doc.createParagraph();
            toc531.setStyle("16");
            toc531.createRun().setText("5.3.1 占位子章节");

            XWPFParagraph toc54 = doc.createParagraph();
            toc54.setStyle("22");
            toc54.createRun().setText("5.4 流程测试");

            XWPFParagraph section52 = doc.createParagraph();
            section52.setStyle("3");
            section52.createRun().setText("5.2 功能测试");

            XWPFParagraph subsection521 = doc.createParagraph();
            subsection521.setStyle("4");
            subsection521.createRun().setText("5.2.1 占位子章节");

            XWPFTable table52 = doc.createTable(6, 4);
            table52.getRow(0).getCell(0).setText("测试项名称");
            table52.getRow(0).getCell(2).setText("标识");
            table52.getRow(1).getCell(0).setText("测试内容");
            table52.getRow(2).getCell(0).setText("测试策略与方法");
            table52.getRow(3).getCell(0).setText("判定准则");
            table52.getRow(4).getCell(0).setText("测试终止条件");
            table52.getRow(5).getCell(0).setText("追踪关系");

            XWPFParagraph section53 = doc.createParagraph();
            section53.setStyle("3");
            section53.createRun().setText("5.3 性能测试");

            XWPFParagraph subsection531 = doc.createParagraph();
            subsection531.setStyle("4");
            subsection531.createRun().setText("5.3.1 占位子章节");

            XWPFTable table53 = doc.createTable(6, 4);
            table53.getRow(0).getCell(0).setText("测试项名称");
            table53.getRow(0).getCell(2).setText("标识");
            table53.getRow(1).getCell(0).setText("测试内容");
            table53.getRow(2).getCell(0).setText("测试策略与方法");
            table53.getRow(3).getCell(0).setText("判定准则");
            table53.getRow(4).getCell(0).setText("测试终止条件");
            table53.getRow(5).getCell(0).setText("追踪关系");

            XWPFParagraph section54 = doc.createParagraph();
            section54.setStyle("3");
            section54.createRun().setText("5.4 流程测试");

            try (FileOutputStream fos = new FileOutputStream(templatePath.toFile())) {
                doc.write(fos);
            }
        }
    }
}
