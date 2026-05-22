package pub.developers.docautogenbyexcel.config;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

/**
 * 表格配置类
 * 从配置文件读取表格识别和填充的配置参数
 */
public class TableConfig {
    
    private static final String CONFIG_FILE = "table-config.properties";
    private static TableConfig instance;
    private final Properties properties;
    
    // 默认值
    private static final String DEFAULT_TESTCASE_REQUIRED_COLUMN = "模块编号";
    private static final String DEFAULT_BASICINFO_COLUMN_TABLENAME = "表格名称";
    private static final String DEFAULT_BASICINFO_COLUMN_FIELDNAME = "字段名";
    private static final String DEFAULT_BASICINFO_COLUMN_FIELDVALUE = "字段值";
    private static final String DEFAULT_LISTDATA_COLUMN_TABLENAME = "表格名称";
    
    private TableConfig() {
        properties = new Properties();
        loadConfig();
    }
    
    /**
     * 获取单例实例
     */
    public static synchronized TableConfig getInstance() {
        if (instance == null) {
            instance = new TableConfig();
        }
        return instance;
    }
    
    /**
     * 重新加载配置（用于测试或动态更新）
     */
    public void reload() {
        loadConfig();
    }
    
    /**
     * 加载配置文件（UTF-8编码）
     */
    private void loadConfig() {
        try (InputStream is = getClass().getClassLoader().getResourceAsStream(CONFIG_FILE)) {
            if (is != null) {
                // 使用UTF-8编码读取配置文件
                try (InputStreamReader reader = new InputStreamReader(is, StandardCharsets.UTF_8)) {
                    properties.load(reader);
                    System.out.println("加载配置文件: " + CONFIG_FILE);
                }
            } else {
                System.out.println("未找到配置文件 " + CONFIG_FILE + "，使用默认配置");
            }
        } catch (IOException e) {
            System.out.println("读取配置文件失败: " + e.getMessage() + "，使用默认配置");
        }
    }
    
    /**
     * 获取配置值
     */
    private String getProperty(String key, String defaultValue) {
        return properties.getProperty(key, defaultValue);
    }
    
    // ==================== 测试用例配置 ====================
    
    /**
     * 获取测试用例Sheet的必填列名
     * 只要Sheet包含此列，即认为是测试用例Sheet
     */
    public String getTestCaseRequiredColumn() {
        return getProperty("testcase.required.column", DEFAULT_TESTCASE_REQUIRED_COLUMN);
    }
    
    // ==================== 基本信息配置 ====================
    
    /**
     * 获取基本信息Sheet的表格名称列名
     */
    public String getBasicInfoTableNameColumn() {
        return getProperty("basicinfo.column.tablename", DEFAULT_BASICINFO_COLUMN_TABLENAME);
    }
    
    /**
     * 获取基本信息Sheet的字段名列名
     */
    public String getBasicInfoFieldNameColumn() {
        return getProperty("basicinfo.column.fieldname", DEFAULT_BASICINFO_COLUMN_FIELDNAME);
    }
    
    /**
     * 获取基本信息Sheet的字段值列名
     */
    public String getBasicInfoFieldValueColumn() {
        return getProperty("basicinfo.column.fieldvalue", DEFAULT_BASICINFO_COLUMN_FIELDVALUE);
    }
    
    // ==================== 列表型数据配置 ====================
    
    /**
     * 获取列表型Sheet的表格名称列名
     */
    public String getListDataTableNameColumn() {
        return getProperty("listdata.column.tablename", DEFAULT_LISTDATA_COLUMN_TABLENAME);
    }
    
    // ==================== 调试配置 ====================
    
    /**
     * 是否启用调试日志
     */
    public boolean isDebugEnabled() {
        return Boolean.parseBoolean(getProperty("debug.enabled", "false"));
    }
    
    /**
     * 判断Sheet是否是测试用例Sheet
     * @param headerColumns Sheet的表头列名列表
     * @return true如果包含测试用例必填列
     */
    public boolean isTestCaseSheet(java.util.List<String> headerColumns) {
        String requiredColumn = getTestCaseRequiredColumn();
        return headerColumns.contains(requiredColumn);
    }
    
    /**
     * 判断Sheet是否是基本信息Sheet
     * @param headerColumns Sheet的表头列名列表
     * @return true如果符合基本信息Sheet的列结构
     */
    public boolean isBasicInfoSheet(java.util.List<String> headerColumns) {
        if (headerColumns.size() < 3) {
            return false;
        }
        return headerColumns.get(0).equals(getBasicInfoTableNameColumn()) &&
               headerColumns.get(1).equals(getBasicInfoFieldNameColumn()) &&
               headerColumns.get(2).equals(getBasicInfoFieldValueColumn());
    }
    
    /**
     * 判断Sheet是否是列表型数据Sheet
     * @param headerColumns Sheet的表头列名列表
     * @return true如果第一列是表格名称列
     */
    public boolean isListDataSheet(java.util.List<String> headerColumns) {
        if (headerColumns.isEmpty()) {
            return false;
        }
        return headerColumns.get(0).equals(getListDataTableNameColumn());
    }
    
    /**
     * 获取Properties对象（用于扩展配置加载）
     */
    public static Properties getProperties() {
        return getInstance().properties;
    }
}

