package pub.developers.docautogenbyexcel.util;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 文件工具类
 */
public class FileUtil {
    
    /**
     * 生成输出文件名
     * 格式：原始模板名_生成时间.docx
     */
    public static String generateOutputFileName(String templatePath, String outputDir) {
        File templateFile = new File(templatePath);
        String templateName = templateFile.getName();
        
        // 移除扩展名
        int lastDot = templateName.lastIndexOf('.');
        if (lastDot > 0) {
            templateName = templateName.substring(0, lastDot);
        }
        
        // 生成时间戳
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");
        String timestamp = sdf.format(new Date());
        
        // 处理输出目录路径
        File outputDirFile = new File(outputDir);
        // 如果是相对路径，转换为绝对路径
        if (!outputDirFile.isAbsolute()) {
            outputDirFile = new File(System.getProperty("user.dir"), outputDir);
        }
        
        // 确保输出目录存在
        if (!outputDirFile.exists()) {
            outputDirFile.mkdirs();
        }
        
        // 构建输出文件路径
        File outputFile = new File(outputDirFile, templateName + "_" + timestamp + ".docx");
        
        // 返回绝对路径，确保路径正确
        try {
            return outputFile.getCanonicalPath();
        } catch (Exception e) {
            return outputFile.getAbsolutePath();
        }
    }
    
    /**
     * 确保目录存在
     */
    public static void ensureDirectoryExists(String dirPath) throws Exception {
        File dir = new File(dirPath);
        if (!dir.exists()) {
            if (!dir.mkdirs()) {
                throw new Exception("无法创建输出目录: " + dirPath);
            }
        }
    }
}

