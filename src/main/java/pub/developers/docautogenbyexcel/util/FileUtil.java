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
        
        // 确保输出目录以/结尾
        if (!outputDir.endsWith("/") && !outputDir.endsWith("\\")) {
            outputDir += File.separator;
        }
        
        return outputDir + templateName + "_" + timestamp + ".docx";
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

