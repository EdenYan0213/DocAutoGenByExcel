package pub.developers.docautogenbyexcel;

import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.nio.file.Path;

public class InspectTemplateTest {

    @Test
    public void inspectTemplate() throws Exception {
        String base = System.getProperty("user.dir");
        Path p = Path.of(base, "output", "2-XX 软件配置项测试报告 (公开）_202603091646.docx");
        if (!p.toFile().exists()) {
            p = Path.of(base, "2-XX 软件配置项测试报告 (公开）.docx");
            if (!p.toFile().exists()) {
                p = Path.of(base, "2-XX软件配置项测试报告(公开）.docx");
            }
        }
        System.out.println("Inspecting: " + p);

        try (FileInputStream fis = new FileInputStream(p.toFile()); XWPFDocument doc = new XWPFDocument(fis)) {
            System.out.println("Paragraphs:");
            int idx = 0;
            for (XWPFParagraph para : doc.getParagraphs()) {
                String style = para.getStyle();
                String text = para.getText();
                if (text != null && !text.trim().isEmpty()) {
                    System.out.println("  [" + idx + "] (style=" + style + ") -> " + text.trim());
                }
                idx++;
            }

            System.out.println("Tables: " + doc.getTables().size());
            int tno = 1;
            for (XWPFTable tbl : doc.getTables()) {
                System.out.println("--- Table #" + tno + " ---");
                // try to find caption immediately before table
                int firstRow = tbl.getNumberOfRows();
                if (firstRow > 0) {
                    XWPFTableRow r = tbl.getRow(0);
                    System.out.print("Header cells: ");
                    for (int i = 0; i < r.getTableCells().size(); i++) {
                        System.out.print("[" + i + ":" + r.getCell(i).getText().replaceAll("\\s+", " ") + "] ");
                    }
                    System.out.println();
                }
                tno++;
            }
        }
    }
}
