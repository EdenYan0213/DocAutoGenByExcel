package pub.developers.docautogenbyexcel;

import org.junit.jupiter.api.Test;
import pub.developers.docautogenbyexcel.reader.ExcelReader;
import pub.developers.docautogenbyexcel.model.ModuleData;

import java.io.File;
import java.util.Map;

public class InspectExcelTest {

    @Test
    public void printParsedExcelData() throws Exception {
        File dir = new File(".");
        File[] files = dir.listFiles((d, name) -> name.startsWith("2 号文档完整填充模板_") && name.endsWith(".xlsx"));
        if (files == null || files.length == 0) {
            System.out.println("未找到生成的 Excel 文件");
            return;
        }
        // pick latest
        File latest = files[0];
        for (File f : files) if (f.lastModified() > latest.lastModified()) latest = f;
        System.out.println("读取 Excel: " + latest.getAbsolutePath());

        ExcelReader reader = new ExcelReader();
        Map<String, ModuleData> map = reader.readExcel(latest.getAbsolutePath());
        System.out.println("模块数量: " + map.size());
        for (Map.Entry<String, ModuleData> e : map.entrySet()) {
            System.out.println("模块: " + e.getKey() + " -> 用例数: " + e.getValue().getTestCaseCount());
            e.getValue().getTestCases().forEach(tc -> System.out.println("  用例: " + tc.getColumnData()));
        }
    }
}
