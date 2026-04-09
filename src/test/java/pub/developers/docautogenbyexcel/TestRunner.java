package pub.developers.docautogenbyexcel;

/**
 * 测试运行器 - 用于测试改进后的代码
 */
public class TestRunner {

    public static void main(String[] args) {
        System.out.println("========================================");
        System.out.println("DocAutoGenByExcel 测试运行器");
        System.out.println("========================================\n");

        try {
            // 使用绝对路径并确保目录存在
            String baseDir = System.getProperty("user.dir");
            String excelPath = baseDir + "/2 号文档完整填充模板.xlsx";
            String wordPath = baseDir + "/2-XX 软件配置项测试报告 (公开）.docx";
            String outputPath = baseDir + "/test_output/test_result_" + System.currentTimeMillis() + ".docx";

            System.out.println("当前工作目录：" + baseDir);
            System.out.println("Excel 文件路径：" + excelPath);
            System.out.println("Excel 文件是否存在：" + new java.io.File(excelPath).exists());
            System.out.println("Word 文件是否存在：" + new java.io.File(wordPath).exists());

            // 如果文件不存在，尝试查找最接近的文件
            if (!new java.io.File(excelPath).exists()) {
                System.out.println("\n查找 Excel 文件...");
                java.io.File dir = new java.io.File(baseDir);
                java.io.File[] xlsxFiles = dir.listFiles((d, name) -> name.endsWith(".xlsx"));
                if (xlsxFiles != null) {
                    for (java.io.File f : xlsxFiles) {
                        System.out.println("  找到：" + f.getName());
                    }
                }
                throw new Exception("未找到 Excel 文件：" + excelPath);
            }

            // 确保输出目录存在
            java.io.File outputDir = new java.io.File(outputPath).getParentFile();
            if (!outputDir.exists()) {
                outputDir.mkdirs();
            }

            // 创建配置文件
            createConfig(excelPath, wordPath, outputPath);

            // 运行工具
            System.out.println("[1/4] 开始读取 Excel 数据...");
            testExcelReading(excelPath);

            System.out.println("\n[2/4] 开始处理 Word 模板...");
            testWordProcessing(excelPath, wordPath, outputPath);

            System.out.println("\n[3/4] 验证输出文件...");
            verifyOutput(outputPath);

            System.out.println("\n========================================");
            System.out.println("✓ 测试成功完成！");
            System.out.println("输出文件：" + outputPath);
            System.out.println("========================================");

        } catch (Exception e) {
            System.err.println("\n========================================");
            System.err.println("✗ 测试失败！");
            System.err.println("错误信息：" + e.getMessage());
            System.err.println("========================================");
            e.printStackTrace();
            System.exit(1);
        }
    }

    private static void createConfig(String excelPath, String wordPath, String outputPath) throws Exception {
        System.out.println("创创建配置文件...");
        // 直接写入配置文件
        java.io.PrintWriter writer = new java.io.PrintWriter("config.properties", "UTF-8");
        writer.println("# Excel 数据驱动 Word 多模块动态表格生成工具 - 配置文件");
        writer.println("excel.path=" + excelPath.replace("\\", "/"));
        writer.println("word.path=" + wordPath.replace("\\", "/"));
        writer.println("output.path=" + outputPath.replace("\\", "/").substring(0, outputPath.lastIndexOf("/")));
        writer.close();
        System.out.println("✓ 配置文件已保存");
    }

    private static void testExcelReading(String excelPath) throws Exception {
        pub.developers.docautogenbyexcel.reader.ExcelReader excelReader = new pub.developers.docautogenbyexcel.reader.ExcelReader();
        java.util.Map<String, pub.developers.docautogenbyexcel.model.ModuleData> moduleDataMap = excelReader
                .readExcel(excelPath);

        System.out.println("  ✓ Excel 文件成功读取");
        System.out.println("  找到 " + moduleDataMap.size() + " 个测试模块");

        // 打印模块信息
        for (java.util.Map.Entry<String, pub.developers.docautogenbyexcel.model.ModuleData> entry : moduleDataMap
                .entrySet()) {
            System.out.println("    - " + entry.getKey() + ": " + entry.getValue().getTestCases().size() + " 条测试用例");
        }
    }

    private static void testWordProcessing(String excelPath, String wordPath, String outputPath) throws Exception {
        // 读取 Excel 数据
        pub.developers.docautogenbyexcel.reader.ExcelReader excelReader = new pub.developers.docautogenbyexcel.reader.ExcelReader();
        java.util.Map<String, pub.developers.docautogenbyexcel.model.ModuleData> moduleDataMap = excelReader
                .readExcel(excelPath);

        // 处理 Word 模板
        if (!moduleDataMap.isEmpty()) {
            pub.developers.docautogenbyexcel.processor.WordProcessor wordProcessor = new pub.developers.docautogenbyexcel.processor.WordProcessor();
            int successCount = wordProcessor.processWord(wordPath, outputPath, moduleDataMap);
            System.out.println("  ✓ Word 模板处理完成");
            System.out.println("  成功处理 " + successCount + " 个测试用例模块");
        } else {
            System.out.println("  ⚠ 未找到测试用例数据，跳过测试用例处理");
            // 复制 Word 模板到输出路径
            java.nio.file.Files.copy(
                    java.nio.file.Paths.get(wordPath),
                    java.nio.file.Paths.get(outputPath),
                    java.nio.file.StandardCopyOption.REPLACE_EXISTING);
        }

        // 处理其他表格
        System.out.println("\n  处理其他表格（基本信息、列表型表格）...");
        processAdditionalTables(excelPath, outputPath, moduleDataMap);
    }

    private static void processAdditionalTables(String excelPath, String outputPath,
            java.util.Map<String, pub.developers.docautogenbyexcel.model.ModuleData> moduleDataMap) throws Exception {
        pub.developers.docautogenbyexcel.reader.TableDataReader tableReader = new pub.developers.docautogenbyexcel.reader.TableDataReader();
        pub.developers.docautogenbyexcel.processor.TableFillProcessor tableFillProcessor = new pub.developers.docautogenbyexcel.processor.TableFillProcessor();

        // 读取基本信息
        System.out.println("  读取基本信息...");
        java.util.Map<String, pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData> basicInfoMap = tableReader
                .readBasicInfo(excelPath);
        System.out.println("    找到 " + basicInfoMap.size() + " 个基本信息表格");

        // 读取所有列表型表格数据
        System.out.println("  读取列表型表格...");
        java.util.Map<String, pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData> allListData = tableReader
                .readAllListTableData(excelPath);
        System.out.println("    找到 " + allListData.size() + " 个列表型表格");

        boolean hasTraceabilityData = moduleDataMap != null && !moduleDataMap.isEmpty();

        // 如果有数据需要填充
        if (!basicInfoMap.isEmpty() || !allListData.isEmpty() || hasTraceabilityData) {
            // 打开文档进行二次处理
            try (java.io.FileInputStream fis = new java.io.FileInputStream(outputPath);
                    org.apache.poi.xwpf.usermodel.XWPFDocument document = new org.apache.poi.xwpf.usermodel.XWPFDocument(
                            fis)) {

                int basicInfoCount = tableFillProcessor.fillBasicInfoTables(document, basicInfoMap);
                int listCount = tableFillProcessor.fillListTables(document, allListData);
                int traceabilityCount = tableFillProcessor.fillTestTraceabilityTables(document, moduleDataMap);

                // 保存文档
                try (java.io.FileOutputStream fos = new java.io.FileOutputStream(outputPath)) {
                    document.write(fos);
                }

                if (basicInfoCount > 0) {
                    System.out.println("    ✓ 填充基本信息表格：" + basicInfoCount + " 个");
                }
                if (listCount > 0) {
                    System.out.println("    ✓ 填充列表型表格：" + listCount + " 个");
                }
                if (traceabilityCount > 0) {
                    System.out.println("    ✓ 填充测试项追踪表：" + traceabilityCount + " 个");
                }
            }
        } else {
            System.out.println("    ⚠ 没有需要填充的其他表格");
        }
    }

    private static void verifyOutput(String outputPath) throws Exception {
        java.io.File outputFile = new java.io.File(outputPath);

        if (!outputFile.exists()) {
            throw new Exception("输出文件不存在：" + outputPath);
        }

        long fileSize = outputFile.length();
        System.out.println("  ✓ 输出文件已生成");
        System.out.println("  文件大小：" + fileSize + " 字节 (" + (fileSize / 1024) + " KB)");

        // 验证 Word 文档可以打开
        try (java.io.FileInputStream fis = new java.io.FileInputStream(outputPath);
                org.apache.poi.xwpf.usermodel.XWPFDocument document = new org.apache.poi.xwpf.usermodel.XWPFDocument(
                        fis)) {
            int tableCount = document.getTables().size();
            int paragraphCount = document.getParagraphs().size();
            System.out.println("  ✓ Word 文档验证通过");
            System.out.println("    表格数：" + tableCount);
            System.out.println("    段落数：" + paragraphCount);
        }
    }
}
