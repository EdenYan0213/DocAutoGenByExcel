package pub.developers.docautogenbyexcel.config;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

/**
 * 配置文件加载器
 */
public class ConfigLoader {
    private static final String CONFIG_FILE = "config.properties";
    
    private String excelPath;
    private String wordPath;
    private String outputPath;

    /**
     * 从配置文件加载配置
     */
    public void loadFromFile() throws Exception {
        Properties props = new Properties();
        try (FileInputStream fis = new FileInputStream(CONFIG_FILE)) {
            props.load(fis);
            
            excelPath = props.getProperty("excel.path");
            wordPath = props.getProperty("word.path");
            outputPath = props.getProperty("output.path");
            
            if (excelPath == null || wordPath == null) {
                throw new Exception("配置文件缺少必填项：excel.path 或 word.path");
            }
            
            if (outputPath == null || outputPath.trim().isEmpty()) {
                // 默认输出路径为Excel文件同目录
                if (excelPath != null) {
                    int lastIndex = excelPath.lastIndexOf('/');
                    if (lastIndex == -1) {
                        lastIndex = excelPath.lastIndexOf('\\');
                    }
                    if (lastIndex != -1) {
                        outputPath = excelPath.substring(0, lastIndex + 1);
                    } else {
                        outputPath = "./";
                    }
                } else {
                    outputPath = "./";
                }
            }
        } catch (IOException e) {
            throw new Exception("读取配置文件失败: " + e.getMessage(), e);
        }
    }

    public String getExcelPath() {
        return excelPath;
    }

    public void setExcelPath(String excelPath) {
        this.excelPath = excelPath;
    }

    public String getWordPath() {
        return wordPath;
    }

    public void setWordPath(String wordPath) {
        this.wordPath = wordPath;
    }

    public String getOutputPath() {
        return outputPath;
    }

    public void setOutputPath(String outputPath) {
        this.outputPath = outputPath;
    }
}

