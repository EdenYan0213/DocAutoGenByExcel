package pub.developers.docautogenbyexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;

public class GenerateTemplateExcelTest {

    @Test
    public void generateExcelMatchingTemplate() throws Exception {
        String out = "generated_for_template.xlsx";
        try (Workbook wb = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(out)) {
            Sheet sheet = wb.createSheet("测试用例");
            Row header = sheet.createRow(0);
            String[] headers = new String[]{"模块编号", "测试项名称", "标识", "测试内容", "测试策略与方法", "判定准则", "测试终止条件", "追踪关系"};
            for (int i = 0; i < headers.length; i++) {
                Cell c = header.createCell(i);
                c.setCellValue(headers[i]);
            }

            // sample rows
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("4.3");
            r1.createCell(1).setCellValue("登录功能");
            r1.createCell(2).setCellValue("F001");
            r1.createCell(3).setCellValue("验证用户登录功能");
            r1.createCell(4).setCellValue("1) 输入正确的用户名和密码");
            r1.createCell(5).setCellValue("正常登录应跳转到首页");
            r1.createCell(6).setCellValue("测试用例执行完成");
            r1.createCell(7).setCellValue("需求文档V1.0");

            Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue("6.2");
            r2.createCell(1).setCellValue("用户登录接口");
            r2.createCell(2).setCellValue("I001");
            r2.createCell(3).setCellValue("验证用户登录接口");
            r2.createCell(4).setCellValue("发送POST请求到/login接口");
            r2.createCell(5).setCellValue("正常登录应返回token和用户信息");
            r2.createCell(6).setCellValue("接口测试执行完成");
            r2.createCell(7).setCellValue("接口文档A");

            wb.write(fos);
        }
        System.out.println("生成Excel文件: " + out);
    }
}
