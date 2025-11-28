package pub.developers.docautogenbyexcel.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

/**
 * 测试数据生成工具
 * 用于生成测试用的Word模板和Excel数据文件
 */
public class TestDataGenerator {

    /**
     * 生成Word模板文件
     */
    public static void generateWordTemplate(String outputPath) throws IOException {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(outputPath)) {

            // 添加标题
            XWPFParagraph titlePara = document.createParagraph();
            titlePara.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = titlePara.createRun();
            titleRun.setText("测试报告");
            titleRun.setBold(true);
            titleRun.setFontSize(18);

            // 添加空行
            document.createParagraph();

            // 添加章节 5.3 功能测试
            XWPFParagraph section1 = document.createParagraph();
            XWPFRun run1 = section1.createRun();
            run1.setText("5.3 功能测试");
            run1.setBold(true);
            run1.setFontSize(14);

            // 添加空行
            document.createParagraph();

            // 添加章节 6.2 接口测试
            XWPFParagraph section2 = document.createParagraph();
            XWPFRun run2 = section2.createRun();
            run2.setText("6.2 接口测试");
            run2.setBold(true);
            run2.setFontSize(14);

            // 添加空行
            document.createParagraph();

            // 添加其他内容示例
            XWPFParagraph otherPara = document.createParagraph();
            XWPFRun otherRun = otherPara.createRun();
            otherRun.setText("7.1 其他测试内容");
            otherRun.setBold(true);
            otherRun.setFontSize(14);

            document.write(out);
            System.out.println("Word模板已生成: " + outputPath);
        }
    }

    /**
     * 生成Excel测试数据文件
     */
    public static void generateExcelData(String outputPath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(outputPath)) {

            Sheet sheet = workbook.createSheet("测试用例");

            // 创建表头
            Row headerRow = sheet.createRow(0);
            String[] headers = {"模块编号", "testName", "id", "content", "strategy", "criteria", "stopCondition", "trace"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                CellStyle headerStyle = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                headerStyle.setFont(font);
                cell.setCellStyle(headerStyle);
            }

            // 创建测试数据 - 模块5.3的数据
            createTestDataRow(sheet, 1, "5.3", "登录功能", "F001", 
                "验证用户登录功能，包含正常登录、密码错误、账户锁定等场景",
                "1) 输入正确的用户名和密码；2) 输入错误的密码；3) 输入已锁定的账户",
                "1) 正常登录应跳转到首页；2) 密码错误应提示\"密码不正确\"；3) 账户锁定应提示\"账户已被锁定\"",
                "测试用例执行完成或出现阻塞性缺陷",
                "需求文档V1.0");

            createTestDataRow(sheet, 2, "5.3", "注册功能", "F002",
                "验证用户注册功能，包含正常注册、邮箱重复、密码强度验证等场景",
                "1) 输入有效的邮箱和密码；2) 输入已存在的邮箱；3) 输入弱密码",
                "1) 正常注册应创建账户并发送验证邮件；2) 邮箱重复应提示\"该邮箱已被注册\"；3) 弱密码应提示\"密码强度不足\"",
                "测试用例执行完成或出现阻塞性缺陷",
                "需求文档V1.0");

            createTestDataRow(sheet, 3, "5.3", "密码重置", "F003",
                "验证用户密码重置功能，包含发送重置邮件、验证重置链接、设置新密码等场景",
                "1) 输入已注册的邮箱；2) 点击重置链接；3) 设置新密码",
                "1) 应发送包含重置链接的邮件；2) 重置链接应有效；3) 新密码应成功设置",
                "测试用例执行完成或出现阻塞性缺陷",
                "需求文档V1.0");

            // 创建测试数据 - 模块6.2的数据
            createTestDataRow(sheet, 4, "6.2", "用户登录接口", "I001",
                "验证用户登录接口，测试正常登录、参数校验、异常处理等场景",
                "1) 发送POST请求到/login接口；2) 测试各种参数组合；3) 验证返回结果",
                "1) 正常登录应返回token和用户信息；2) 参数缺失应返回400错误；3) 密码错误应返回401错误",
                "接口测试执行完成或出现阻塞性缺陷",
                "接口文档A");

            createTestDataRow(sheet, 5, "6.2", "用户注册接口", "I002",
                "验证用户注册接口，测试正常注册、参数校验、数据验证等场景",
                "1) 发送POST请求到/register接口；2) 测试各种参数组合；3) 验证返回结果",
                "1) 正常注册应返回201状态码和用户ID；2) 邮箱格式错误应返回400错误；3) 邮箱重复应返回409错误",
                "接口测试执行完成或出现阻塞性缺陷",
                "接口文档A");

            createTestDataRow(sheet, 6, "6.2", "获取用户信息接口", "I003",
                "验证获取用户信息接口，测试正常获取、权限校验、异常处理等场景",
                "1) 发送GET请求到/user/{id}接口；2) 测试不同权限的用户；3) 验证返回结果",
                "1) 正常获取应返回用户详细信息；2) 无权限应返回403错误；3) 用户不存在应返回404错误",
                "接口测试执行完成或出现阻塞性缺陷",
                "接口文档A");

            // 自动调整列宽
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(out);
            System.out.println("Excel测试数据已生成: " + outputPath);
        }
    }

    /**
     * 创建测试数据行
     */
    private static void createTestDataRow(Sheet sheet, int rowNum, String moduleNumber,
                                         String testName, String id, String content,
                                         String strategy, String criteria,
                                         String stopCondition, String trace) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(moduleNumber);
        row.createCell(1).setCellValue(testName);
        row.createCell(2).setCellValue(id);
        row.createCell(3).setCellValue(content);
        row.createCell(4).setCellValue(strategy);
        row.createCell(5).setCellValue(criteria);
        row.createCell(6).setCellValue(stopCondition);
        row.createCell(7).setCellValue(trace);
    }

    /**
     * 主方法：生成测试文件
     */
    public static void main(String[] args) {
        try {
            // 生成Word模板
            String wordTemplatePath = "test_template.docx";
            generateWordTemplate(wordTemplatePath);
            System.out.println("✓ Word模板生成成功: " + wordTemplatePath);

            // 生成Excel数据
            String excelDataPath = "test_data.xlsx";
            generateExcelData(excelDataPath);
            System.out.println("✓ Excel数据生成成功: " + excelDataPath);

            System.out.println("\n测试文件生成完成！");
            System.out.println("可以使用以下命令运行工具：");
            System.out.println("java -jar target/DocAutoGenByExcel-0.0.1-SNAPSHOT.jar -excel " + excelDataPath + " -word " + wordTemplatePath + " -out ./output");

        } catch (Exception e) {
            System.err.println("生成测试文件失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}


