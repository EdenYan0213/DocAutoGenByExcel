package pub.developers.docautogenbyexcel;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;

public class InspectOutputTest {

    @Test
    public void printTestCaseTables() throws Exception {
        File outDir = new File("test_output");
        if (!outDir.exists() || !outDir.isDirectory()) {
            System.out.println("test_output 目录不存在");
            return;
        }

        File[] files = outDir.listFiles((d, name) -> name.startsWith("test_result_") && name.endsWith(".docx"));
        if (files == null || files.length == 0) {
            System.out.println("未找到输出的 docx 文件");
            return;
        }

        Arrays.sort(files, (a, b) -> Long.compare(b.lastModified(), a.lastModified()));
        File latest = files[0];
        System.out.println("检查文件: " + latest.getAbsolutePath());

        try (FileInputStream fis = new FileInputStream(latest);
             XWPFDocument doc = new XWPFDocument(fis)) {

            int idx = 0;
            for (XWPFTable table : doc.getTables()) {
                idx++;
                String txt = table.getText();
                if (txt.contains("测试用例标识") || txt.contains("测试项名称") || txt.contains("测试步骤")) {
                    System.out.println("--- 表格 #" + idx + " ---");
                    System.out.println(txt.replaceAll("\r?\n", " | "));
                }
            }
        }
    }
}
