package pub.developers.docautogenbyexcel;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import pub.developers.docautogenbyexcel.util.SectionNumberUtil;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 章节编号工具类测试
 * 测试多级目录支持
 */
@DisplayName("章节编号工具类测试")
public class SectionNumberUtilTest {

    @Test
    @DisplayName("测试章节级别识别")
    void testGetLevel() {
        // 一级目录
        assertEquals(1, SectionNumberUtil.getLevel("1"));
        assertEquals(1, SectionNumberUtil.getLevel("5"));
        
        // 二级目录
        assertEquals(2, SectionNumberUtil.getLevel("1.1"));
        assertEquals(2, SectionNumberUtil.getLevel("5.3"));
        
        // 三级目录
        assertEquals(3, SectionNumberUtil.getLevel("1.1.1"));
        assertEquals(3, SectionNumberUtil.getLevel("5.3.2"));
        
        // 四级目录
        assertEquals(4, SectionNumberUtil.getLevel("1.1.1.1"));
        assertEquals(4, SectionNumberUtil.getLevel("5.3.2.4"));
        
        // 五级目录
        assertEquals(5, SectionNumberUtil.getLevel("1.1.1.1.1"));
        assertEquals(5, SectionNumberUtil.getLevel("5.3.2.4.6"));
        
        // 边界情况
        assertEquals(0, SectionNumberUtil.getLevel(null));
        assertEquals(0, SectionNumberUtil.getLevel(""));
    }
    
    @Test
    @DisplayName("测试章节编号提取")
    void testExtractSectionNumber() {
        // 标准格式
        assertEquals("1.1", SectionNumberUtil.extractSectionNumber("1.1 功能测试"));
        assertEquals("5.3.2", SectionNumberUtil.extractSectionNumber("5.3.2 登录功能测试"));
        assertEquals("5.3.2.1", SectionNumberUtil.extractSectionNumber("5.3.2.1 用户名验证测试"));
        assertEquals("5.3.2.1.1", SectionNumberUtil.extractSectionNumber("5.3.2.1.1 特殊字符验证"));
        
        // 无效格式
        assertNull(SectionNumberUtil.extractSectionNumber("功能测试"));
        assertNull(SectionNumberUtil.extractSectionNumber(null));
        assertNull(SectionNumberUtil.extractSectionNumber(""));
    }
    
    @Test
    @DisplayName("测试父级编号获取")
    void testGetParentNumber() {
        assertEquals("5", SectionNumberUtil.getParentNumber("5.3"));
        assertEquals("5.3", SectionNumberUtil.getParentNumber("5.3.2"));
        assertEquals("5.3.2", SectionNumberUtil.getParentNumber("5.3.2.1"));
        assertEquals("5.3.2.1", SectionNumberUtil.getParentNumber("5.3.2.1.1"));
        
        // 顶级没有父级
        assertNull(SectionNumberUtil.getParentNumber("5"));
        assertNull(SectionNumberUtil.getParentNumber(null));
    }
    
    @Test
    @DisplayName("测试子章节判断")
    void testIsChildOf() {
        // 直接子章节
        assertTrue(SectionNumberUtil.isChildOf("5.3", "5"));
        assertTrue(SectionNumberUtil.isChildOf("5.3.2", "5.3"));
        assertTrue(SectionNumberUtil.isChildOf("5.3.2.1", "5.3.2"));
        assertTrue(SectionNumberUtil.isChildOf("5.3.2.1.1", "5.3.2.1"));
        
        // 非直接子章节（后代但不是子）
        assertFalse(SectionNumberUtil.isChildOf("5.3.2", "5"));
        assertFalse(SectionNumberUtil.isChildOf("5.3.2.1", "5"));
        
        // 不是子章节
        assertFalse(SectionNumberUtil.isChildOf("6.1", "5"));
        assertFalse(SectionNumberUtil.isChildOf("5", "5.3"));
    }
    
    @Test
    @DisplayName("测试后代章节判断")
    void testIsDescendantOf() {
        // 所有后代
        assertTrue(SectionNumberUtil.isDescendantOf("5.3", "5"));
        assertTrue(SectionNumberUtil.isDescendantOf("5.3.2", "5"));
        assertTrue(SectionNumberUtil.isDescendantOf("5.3.2.1", "5"));
        assertTrue(SectionNumberUtil.isDescendantOf("5.3.2.1.1", "5"));
        
        assertTrue(SectionNumberUtil.isDescendantOf("5.3.2", "5.3"));
        assertTrue(SectionNumberUtil.isDescendantOf("5.3.2.1", "5.3"));
        assertTrue(SectionNumberUtil.isDescendantOf("5.3.2.1.1", "5.3"));
        
        // 不是后代
        assertFalse(SectionNumberUtil.isDescendantOf("6.1", "5"));
        assertFalse(SectionNumberUtil.isDescendantOf("5", "5.3"));
    }
    
    @Test
    @DisplayName("测试子章节编号生成")
    void testGenerateChildNumber() {
        assertEquals("5.1", SectionNumberUtil.generateChildNumber("5", 1));
        assertEquals("5.3", SectionNumberUtil.generateChildNumber("5", 3));
        
        assertEquals("5.3.1", SectionNumberUtil.generateChildNumber("5.3", 1));
        assertEquals("5.3.2", SectionNumberUtil.generateChildNumber("5.3", 2));
        
        assertEquals("5.3.2.1", SectionNumberUtil.generateChildNumber("5.3.2", 1));
        assertEquals("5.3.2.5", SectionNumberUtil.generateChildNumber("5.3.2", 5));
        
        assertEquals("5.3.2.1.1", SectionNumberUtil.generateChildNumber("5.3.2.1", 1));
        
        // 顶级
        assertEquals("1", SectionNumberUtil.generateChildNumber(null, 1));
        assertEquals("1", SectionNumberUtil.generateChildNumber("", 1));
    }
    
    @Test
    @DisplayName("测试章节编号比较")
    void testCompare() {
        // 同级比较
        assertTrue(SectionNumberUtil.compare("5.1", "5.2") < 0);
        assertTrue(SectionNumberUtil.compare("5.2", "5.1") > 0);
        assertEquals(0, SectionNumberUtil.compare("5.1", "5.1"));
        
        // 不同级比较
        assertTrue(SectionNumberUtil.compare("5", "5.1") < 0);
        assertTrue(SectionNumberUtil.compare("5.1", "5") > 0);
        assertTrue(SectionNumberUtil.compare("5.3.2", "5.3.2.1") < 0);
        
        // 跨分支比较
        assertTrue(SectionNumberUtil.compare("5.3", "5.4") < 0);
        assertTrue(SectionNumberUtil.compare("5.3.9", "5.4.1") < 0);
    }
    
    @Test
    @DisplayName("测试占位符识别")
    void testPlaceholder() {
        // 有效占位符
        assertTrue(SectionNumberUtil.isPlaceholder("5.x"));
        assertTrue(SectionNumberUtil.isPlaceholder("5.X"));
        assertTrue(SectionNumberUtil.isPlaceholder("5.x 功能测试"));
        assertTrue(SectionNumberUtil.isPlaceholder("5.3.x"));
        assertTrue(SectionNumberUtil.isPlaceholder("5.3.x 登录功能测试"));
        assertTrue(SectionNumberUtil.isPlaceholder("5.3.2.x"));
        
        // 无效占位符
        assertFalse(SectionNumberUtil.isPlaceholder("5.3"));
        assertFalse(SectionNumberUtil.isPlaceholder("功能测试"));
        
        // 提取父级
        assertEquals("5", SectionNumberUtil.extractPlaceholderParent("5.x"));
        assertEquals("5.3", SectionNumberUtil.extractPlaceholderParent("5.3.x"));
        assertEquals("5.3.2", SectionNumberUtil.extractPlaceholderParent("5.3.2.x"));
    }
    
    @Test
    @DisplayName("测试样式判断")
    void testStyles() {
        // Heading样式
        assertTrue(SectionNumberUtil.isHeadingStyle("2"));
        assertTrue(SectionNumberUtil.isHeadingStyle("3"));
        assertTrue(SectionNumberUtil.isHeadingStyle("4"));
        assertTrue(SectionNumberUtil.isHeadingStyle("5"));
        assertTrue(SectionNumberUtil.isHeadingStyle("Heading 1"));
        assertTrue(SectionNumberUtil.isHeadingStyle("标题 3"));
        assertFalse(SectionNumberUtil.isHeadingStyle("11"));
        
        // TOC样式
        assertTrue(SectionNumberUtil.isTocStyle("22"));
        assertTrue(SectionNumberUtil.isTocStyle("25"));
        assertTrue(SectionNumberUtil.isTocStyle("16"));
        assertTrue(SectionNumberUtil.isTocStyle("toc1"));
        assertFalse(SectionNumberUtil.isTocStyle("4"));
        
        // Caption样式
        assertTrue(SectionNumberUtil.isCaptionStyle("11"));
        assertTrue(SectionNumberUtil.isCaptionStyle("Caption"));
        assertTrue(SectionNumberUtil.isCaptionStyle("题注"));
        assertFalse(SectionNumberUtil.isCaptionStyle("4"));
    }
}

