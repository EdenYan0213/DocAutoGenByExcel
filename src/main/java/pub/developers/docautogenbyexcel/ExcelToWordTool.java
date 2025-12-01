package pub.developers.docautogenbyexcel;

import org.apache.commons.cli.*;
import pub.developers.docautogenbyexcel.config.ConfigLoader;
import pub.developers.docautogenbyexcel.processor.WordProcessor;
import pub.developers.docautogenbyexcel.reader.ExcelReader;
import pub.developers.docautogenbyexcel.util.FileUtil;

import java.io.File;
import java.util.Map;

/**
 * Excel数据驱动Word多模块动态表格生成工具
 * 主程序入口
 */
public class ExcelToWordTool {

    public static void main(String[] args) {
        try {
            // 解析命令行参数或加载配置文件
            ConfigLoader config = parseArguments(args);
            
            // 验证文件路径
            validatePaths(config);
            
            // 读取Excel数据
            System.out.println("开始读取Excel数据");
            ExcelReader excelReader = new ExcelReader();
            Map<String, pub.developers.docautogenbyexcel.model.ModuleData> moduleDataMap = 
                excelReader.readExcel(config.getExcelPath());
            
            // 生成输出文件路径
            String outputPath = FileUtil.generateOutputFileName(
                config.getWordPath(), 
                config.getOutputPath()
            );
            
            // 确保输出目录存在
            File outputFile = new File(outputPath);
            FileUtil.ensureDirectoryExists(outputFile.getParent());
            
            // 处理Word模板
            System.out.println("开始处理Word模板");
            WordProcessor wordProcessor = new WordProcessor();
            int successCount = wordProcessor.processWord(
                config.getWordPath(), 
                outputPath, 
                moduleDataMap
            );
            
            System.out.println("生成成功！输出文件: " + outputPath);
            System.out.println("成功处理 " + successCount + " 个模块");
            
        } catch (Exception e) {
            System.err.println("错误: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }

    /**
     * 解析命令行参数或加载配置文件
     */
    private static ConfigLoader parseArguments(String[] args) throws Exception {
        ConfigLoader config = new ConfigLoader();
        
        // 创建命令行选项
        Options options = new Options();
        options.addOption("excel", true, "Excel文件路径");
        options.addOption("word", true, "Word模板文件路径");
        options.addOption("out", true, "输出目录路径");
        options.addOption("config", false, "使用配置文件");
        options.addOption("h", "help", false, "显示帮助信息");

        CommandLineParser parser = new DefaultParser();
        CommandLine cmd = parser.parse(options, args);

        // 显示帮助信息
        if (cmd.hasOption("h")) {
            printHelp(options);
            System.exit(0);
        }

        // 如果指定使用配置文件
        if (cmd.hasOption("config") || (args.length == 0)) {
            config.loadFromFile();
        } else {
            // 从命令行参数读取
            if (!cmd.hasOption("excel") || !cmd.hasOption("word")) {
                throw new Exception("缺少必填参数：-excel 和 -word");
            }
            
            config.setExcelPath(cmd.getOptionValue("excel"));
            config.setWordPath(cmd.getOptionValue("word"));
            
            if (cmd.hasOption("out")) {
                config.setOutputPath(cmd.getOptionValue("out"));
            } else {
                // 默认输出路径为Excel文件同目录
                File excelFile = new File(config.getExcelPath());
                String parentPath = excelFile.getParent();
                // 如果没有父目录（例如只是文件名），使用当前目录
                config.setOutputPath(parentPath != null ? parentPath : ".");
            }
        }

        return config;
    }

    /**
     * 验证文件路径
     */
    private static void validatePaths(ConfigLoader config) throws Exception {
        // 验证配置不为空
        if (config.getExcelPath() == null || config.getExcelPath().trim().isEmpty()) {
            throw new Exception("Excel文件路径不能为空");
        }
        if (config.getWordPath() == null || config.getWordPath().trim().isEmpty()) {
            throw new Exception("Word模板文件路径不能为空");
        }
        if (config.getOutputPath() == null || config.getOutputPath().trim().isEmpty()) {
            throw new Exception("输出路径不能为空");
        }
        
        // 验证Excel文件
        File excelFile = new File(config.getExcelPath());
        if (!excelFile.exists()) {
            throw new Exception("Excel文件不存在: " + config.getExcelPath());
        }
        if (!excelFile.canRead()) {
            throw new Exception("无法读取Excel文件: " + config.getExcelPath());
        }

        // 验证Word模板文件
        File wordFile = new File(config.getWordPath());
        if (!wordFile.exists()) {
            throw new Exception("Word模板文件不存在: " + config.getWordPath());
        }
        if (!wordFile.canRead()) {
            throw new Exception("无法读取Word模板文件: " + config.getWordPath());
        }

        // 验证输出目录
        File outputDir = new File(config.getOutputPath());
        if (outputDir.exists() && !outputDir.isDirectory()) {
            throw new Exception("输出路径不是目录: " + config.getOutputPath());
        }
    }

    /**
     * 打印帮助信息
     */
    private static void printHelp(Options options) {
        HelpFormatter formatter = new HelpFormatter();
        formatter.printHelp("ExcelToWordTool", options);
        System.out.println("\n使用示例:");
        System.out.println("  java -jar ExcelToWordTool.jar -excel \"D:\\data.xlsx\" -word \"D:\\template.docx\" -out \"D:\\output\"");
        System.out.println("  java -jar ExcelToWordTool.jar -config  # 使用config.properties配置文件");
    }
}

