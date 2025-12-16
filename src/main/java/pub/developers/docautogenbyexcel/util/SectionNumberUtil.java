package pub.developers.docautogenbyexcel.util;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 章节编号工具类
 * 支持任意级别的章节编号，如：1、1.1、1.1.1、1.1.1.1 等
 */
public class SectionNumberUtil {
    
    // 通用章节编号匹配模式：支持任意级别（1、1.1、1.1.1、1.1.1.1...）
    private static final Pattern SECTION_NUMBER_PATTERN = Pattern.compile("^(\\d+(?:\\.\\d+)*)$");
    
    // 章节标题匹配模式：编号 + 空格 + 名称
    private static final Pattern SECTION_TITLE_PATTERN = Pattern.compile("^(\\d+(?:\\.\\d+)*)\\s+(.+)$");
    
    // 占位符匹配模式：X.x 或 X.X.x 等
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("^(\\d+(?:\\.\\d+)*)\\.x\\s*(.+)?$", Pattern.CASE_INSENSITIVE);
    
    /**
     * 判断字符串是否是有效的章节编号
     */
    public static boolean isValidSectionNumber(String text) {
        if (text == null || text.isEmpty()) {
            return false;
        }
        return SECTION_NUMBER_PATTERN.matcher(text.trim()).matches();
    }
    
    /**
     * 从章节标题中提取编号
     * @param title 章节标题，如 "1.2.3 功能测试"
     * @return 编号，如 "1.2.3"；如果不匹配返回null
     */
    public static String extractSectionNumber(String title) {
        if (title == null || title.isEmpty()) {
            return null;
        }
        Matcher matcher = SECTION_TITLE_PATTERN.matcher(title.trim());
        if (matcher.matches()) {
            return matcher.group(1);
        }
        return null;
    }
    
    /**
     * 从章节标题中提取名称
     * @param title 章节标题，如 "1.2.3 功能测试"
     * @return 名称，如 "功能测试"；如果不匹配返回null
     */
    public static String extractSectionName(String title) {
        if (title == null || title.isEmpty()) {
            return null;
        }
        Matcher matcher = SECTION_TITLE_PATTERN.matcher(title.trim());
        if (matcher.matches()) {
            String name = matcher.group(2).trim();
            // 去除可能的页码（以制表符分隔）
            int tabIndex = name.indexOf('\t');
            if (tabIndex > 0) {
                name = name.substring(0, tabIndex).trim();
            }
            return name;
        }
        return null;
    }
    
    /**
     * 获取章节级别
     * @param sectionNumber 章节编号，如 "1.2.3"
     * @return 级别数，如 3（1是一级，1.1是二级，1.1.1是三级）
     */
    public static int getLevel(String sectionNumber) {
        if (sectionNumber == null || sectionNumber.isEmpty()) {
            return 0;
        }
        return sectionNumber.split("\\.").length;
    }
    
    /**
     * 获取父级章节编号
     * @param sectionNumber 章节编号，如 "1.2.3"
     * @return 父级编号，如 "1.2"；如果是顶级返回null
     */
    public static String getParentNumber(String sectionNumber) {
        if (sectionNumber == null || sectionNumber.isEmpty()) {
            return null;
        }
        int lastDot = sectionNumber.lastIndexOf('.');
        if (lastDot > 0) {
            return sectionNumber.substring(0, lastDot);
        }
        return null;
    }
    
    /**
     * 判断是否是指定父级的子章节
     * @param childNumber 子章节编号
     * @param parentNumber 父章节编号
     * @return 是否是子章节
     */
    public static boolean isChildOf(String childNumber, String parentNumber) {
        if (childNumber == null || parentNumber == null) {
            return false;
        }
        return childNumber.startsWith(parentNumber + ".") && 
               getLevel(childNumber) == getLevel(parentNumber) + 1;
    }
    
    /**
     * 判断是否是指定父级的后代章节（包括子、孙等）
     */
    public static boolean isDescendantOf(String descendantNumber, String ancestorNumber) {
        if (descendantNumber == null || ancestorNumber == null) {
            return false;
        }
        return descendantNumber.startsWith(ancestorNumber + ".");
    }
    
    /**
     * 生成子章节编号
     * @param parentNumber 父章节编号，如 "1.2"
     * @param index 子章节索引（从1开始）
     * @return 子章节编号，如 "1.2.1"
     */
    public static String generateChildNumber(String parentNumber, int index) {
        if (parentNumber == null || parentNumber.isEmpty()) {
            return String.valueOf(index);
        }
        return parentNumber + "." + index;
    }
    
    /**
     * 判断字符串是否是占位符
     * @param text 文本，如 "5.x" 或 "5.2.x 功能测试"
     * @return 是否是占位符
     */
    public static boolean isPlaceholder(String text) {
        if (text == null || text.isEmpty()) {
            return false;
        }
        return PLACEHOLDER_PATTERN.matcher(text.trim()).matches();
    }
    
    /**
     * 从占位符中提取父级编号
     * @param placeholder 占位符文本，如 "5.x" 或 "5.2.x 功能测试"
     * @return 父级编号，如 "5" 或 "5.2"
     */
    public static String extractPlaceholderParent(String placeholder) {
        if (placeholder == null || placeholder.isEmpty()) {
            return null;
        }
        Matcher matcher = PLACEHOLDER_PATTERN.matcher(placeholder.trim());
        if (matcher.matches()) {
            return matcher.group(1);
        }
        return null;
    }
    
    /**
     * 比较两个章节编号的顺序
     * @return 负数表示a在b前面，0表示相等，正数表示a在b后面
     */
    public static int compare(String a, String b) {
        if (a == null && b == null) return 0;
        if (a == null) return -1;
        if (b == null) return 1;
        
        String[] partsA = a.split("\\.");
        String[] partsB = b.split("\\.");
        
        int minLength = Math.min(partsA.length, partsB.length);
        for (int i = 0; i < minLength; i++) {
            int numA = Integer.parseInt(partsA[i]);
            int numB = Integer.parseInt(partsB[i]);
            if (numA != numB) {
                return numA - numB;
            }
        }
        
        return partsA.length - partsB.length;
    }
    
    /**
     * 获取Heading样式ID（根据章节级别）
     * 级别1 -> Heading 1 (样式ID "2")
     * 级别2 -> Heading 2 (样式ID "3")
     * 级别3 -> Heading 3 (样式ID "4")
     * 级别4+ -> Heading 4 (样式ID "5")
     */
    public static String getHeadingStyleId(int level) {
        return switch (level) {
            case 1 -> "2";
            case 2 -> "3";
            case 3 -> "4";
            default -> "5";
        };
    }
    
    /**
     * 判断样式ID是否是Heading样式
     */
    public static boolean isHeadingStyle(String styleId) {
        if (styleId == null) return false;
        return styleId.equals("2") || styleId.equals("3") || 
               styleId.equals("4") || styleId.equals("5") ||
               styleId.toLowerCase().contains("heading") ||
               styleId.contains("标题");
    }
    
    /**
     * 判断样式ID是否是TOC（目录）样式
     */
    public static boolean isTocStyle(String styleId) {
        if (styleId == null) return false;
        return styleId.equals("22") || styleId.equals("25") || 
               styleId.equals("16") || styleId.toLowerCase().startsWith("toc");
    }
    
    /**
     * 判断样式ID是否是Caption样式
     */
    public static boolean isCaptionStyle(String styleId) {
        if (styleId == null) return false;
        return styleId.equals("11") || 
               styleId.equalsIgnoreCase("caption") || 
               styleId.contains("题注");
    }
}

