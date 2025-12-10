package pub.developers.docautogenbyexcel.generator;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import pub.developers.docautogenbyexcel.manager.RequirementManager;
import pub.developers.docautogenbyexcel.model.Requirement;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.*;

/**
 * 需求树Word文档生成器
 * 将需求树以树状结构输出到Word文档
 */
public class RequirementTreeWordGenerator {
    
    /**
     * 生成需求树Word文档
     * 
     * @param requirementManager 需求管理器
     * @param outputPath 输出文件路径
     * @throws Exception 生成异常
     */
    public void generateRequirementTreeWord(RequirementManager requirementManager, String outputPath) throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(outputPath)) {
            
            // 1. 添加标题
            addTitle(document, "需求树结构");
            
            // 2. 添加需求树
            RequirementManager.RequirementTree tree = requirementManager.getRequirementTree();
            List<Requirement> roots = tree.getRoots();
            
            if (roots.isEmpty()) {
                addParagraph(document, "暂无需求数据", false, 0);
            } else {
                for (Requirement root : roots) {
                    addRequirementNode(document, root, 0);
                }
            }
            
            // 3. 添加统计信息
            addStatistics(document, requirementManager);
            
            // 保存文档
            document.write(out);
            System.out.println("需求树Word文档生成成功: " + outputPath);
        } catch (IOException e) {
            throw new Exception("生成Word文档失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 添加标题
     */
    private void addTitle(XWPFDocument document, String title) {
        XWPFParagraph titlePara = document.createParagraph();
        titlePara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = titlePara.createRun();
        titleRun.setText(title);
        titleRun.setBold(true);
        titleRun.setFontSize(18);
        
        // 添加空行
        document.createParagraph();
    }
    
    /**
     * 递归添加需求节点（树状结构）
     */
    private void addRequirementNode(XWPFDocument document, Requirement requirement, int level) {
        // 创建段落
        XWPFParagraph para = document.createParagraph();
        
        // 设置缩进（每级缩进2个字符）
        CTPPr ppr = para.getCTP().isSetPPr() ? para.getCTP().getPPr() : para.getCTP().addNewPPr();
        CTInd ind = ppr.isSetInd() ? ppr.getInd() : ppr.addNewInd();
        ind.setLeft(BigInteger.valueOf(level * 400)); // 每级缩进400 twips (约0.2英寸)
        
        // 创建文本运行
        XWPFRun run = para.createRun();
        
        // 构建节点文本
        StringBuilder nodeText = new StringBuilder();
        
        // 添加树状符号
        if (level > 0) {
            // 使用不同的符号表示层级
            if (requirement.hasChildren()) {
                nodeText.append("├─ "); // 有子节点的节点
            } else {
                nodeText.append("└─ "); // 叶子节点
            }
        }
        
        // 添加需求编号和名称
        nodeText.append(requirement.getRequirementNumber())
                .append(" ")
                .append(requirement.getRequirementName());
        
        // 添加需求类型和优先级（可选）
        if (requirement.getType() != null) {
            nodeText.append(" [").append(requirement.getType().getDescription()).append("]");
        }
        if (requirement.getPriority() != null) {
            nodeText.append(" (").append(requirement.getPriority().getDescription()).append(")");
        }
        
        run.setText(nodeText.toString());
        run.setFontSize(12);
        
        // 如果是根节点或重要节点，加粗
        if (level == 0 || requirement.getPriority() == Requirement.Priority.HIGH) {
            run.setBold(true);
        }
        
        // 如果有描述，添加描述（缩进更多）
        if (requirement.getDescription() != null && !requirement.getDescription().trim().isEmpty()) {
            XWPFParagraph descPara = document.createParagraph();
            CTPPr descPpr = descPara.getCTP().isSetPPr() ? descPara.getCTP().getPPr() : descPara.getCTP().addNewPPr();
            CTInd descInd = descPpr.isSetInd() ? descPpr.getInd() : descPpr.addNewInd();
            descInd.setLeft(BigInteger.valueOf((level + 1) * 400));
            
            XWPFRun descRun = descPara.createRun();
            descRun.setText("    " + requirement.getDescription());
            descRun.setFontSize(10);
            descRun.setItalic(true);
            descRun.setColor("666666");
        }
        
        // 递归添加子节点
        List<Requirement> children = requirement.getChildren();
        if (children != null && !children.isEmpty()) {
            for (Requirement child : children) {
                addRequirementNode(document, child, level + 1);
            }
        }
    }
    
    /**
     * 添加统计信息
     */
    private void addStatistics(XWPFDocument document, RequirementManager manager) {
        // 添加分隔线
        document.createParagraph();
        XWPFParagraph separator = document.createParagraph();
        XWPFRun sepRun = separator.createRun();
        sepRun.setText("────────────────────────────────────────────────────────");
        sepRun.setColor("CCCCCC");
        document.createParagraph();
        
        // 添加统计标题
        XWPFParagraph statTitle = document.createParagraph();
        XWPFRun statTitleRun = statTitle.createRun();
        statTitleRun.setText("统计信息");
        statTitleRun.setBold(true);
        statTitleRun.setFontSize(14);
        document.createParagraph();
        
        // 统计信息
        List<Requirement> allReqs = manager.getAllRequirements();
        List<Requirement> rootReqs = manager.getRootRequirements();
        List<Requirement> leafReqs = manager.getLeafRequirements();
        
        addStatisticItem(document, "总需求数", String.valueOf(allReqs.size()));
        addStatisticItem(document, "根需求数", String.valueOf(rootReqs.size()));
        addStatisticItem(document, "叶子需求数", String.valueOf(leafReqs.size()));
        addStatisticItem(document, "非叶子需求数", String.valueOf(allReqs.size() - leafReqs.size()));
        
        // 计算最大层级深度
        int maxDepth = 0;
        for (Requirement req : allReqs) {
            int depth = manager.getRequirementLevel(req.getRequirementId());
            maxDepth = Math.max(maxDepth, depth);
        }
        addStatisticItem(document, "最大层级深度", String.valueOf(maxDepth + 1));
    }
    
    /**
     * 添加统计项
     */
    private void addStatisticItem(XWPFDocument document, String label, String value) {
        XWPFParagraph para = document.createParagraph();
        para.setIndentationLeft(400);
        
        XWPFRun run = para.createRun();
        run.setText(label + ": " + value);
        run.setFontSize(11);
    }
    
    /**
     * 添加普通段落
     */
    private void addParagraph(XWPFDocument document, String text, boolean bold, int indentLevel) {
        XWPFParagraph para = document.createParagraph();
        
        if (indentLevel > 0) {
            CTPPr ppr = para.getCTP().isSetPPr() ? para.getCTP().getPPr() : para.getCTP().addNewPPr();
            CTInd ind = ppr.isSetInd() ? ppr.getInd() : ppr.addNewInd();
            ind.setLeft(BigInteger.valueOf(indentLevel * 400));
        }
        
        XWPFRun run = para.createRun();
        run.setText(text);
        run.setBold(bold);
        run.setFontSize(12);
    }
    
    /**
     * 生成需求树Word文档（带追溯信息）
     * 
     * @param requirementManager 需求管理器
     * @param traceabilityManager 追溯管理器（可选）
     * @param outputPath 输出文件路径
     * @throws Exception 生成异常
     */
    public void generateRequirementTreeWordWithTraceability(
            RequirementManager requirementManager,
            pub.developers.docautogenbyexcel.manager.TraceabilityManager traceabilityManager,
            String outputPath) throws Exception {
        
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(outputPath)) {
            
            // 1. 添加标题
            addTitle(document, "需求树结构（含追溯信息）");
            
            // 2. 添加需求树（带追溯信息）
            RequirementManager.RequirementTree tree = requirementManager.getRequirementTree();
            List<Requirement> roots = tree.getRoots();
            
            if (roots.isEmpty()) {
                addParagraph(document, "暂无需求数据", false, 0);
            } else {
                for (Requirement root : roots) {
                    addRequirementNodeWithTraceability(document, root, 0, requirementManager, traceabilityManager);
                }
            }
            
            // 3. 添加统计信息
            addStatistics(document, requirementManager);
            
            // 4. 添加追溯统计（如果有追溯管理器）
            if (traceabilityManager != null) {
                addTraceabilityStatistics(document, requirementManager, traceabilityManager);
            }
            
            // 保存文档
            document.write(out);
            System.out.println("需求树Word文档（含追溯信息）生成成功: " + outputPath);
        } catch (IOException e) {
            throw new Exception("生成Word文档失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 递归添加需求节点（带追溯信息）
     */
    private void addRequirementNodeWithTraceability(
            XWPFDocument document, 
            Requirement requirement, 
            int level,
            RequirementManager requirementManager,
            pub.developers.docautogenbyexcel.manager.TraceabilityManager traceabilityManager) {
        
        // 创建段落
        XWPFParagraph para = document.createParagraph();
        
        // 设置缩进
        CTPPr ppr = para.getCTP().isSetPPr() ? para.getCTP().getPPr() : para.getCTP().addNewPPr();
        CTInd ind = ppr.isSetInd() ? ppr.getInd() : ppr.addNewInd();
        ind.setLeft(BigInteger.valueOf(level * 400));
        
        // 创建文本运行
        XWPFRun run = para.createRun();
        
        // 构建节点文本
        StringBuilder nodeText = new StringBuilder();
        
        // 添加树状符号
        if (level > 0) {
            if (requirement.hasChildren()) {
                nodeText.append("├─ ");
            } else {
                nodeText.append("└─ ");
            }
        }
        
        // 添加需求编号和名称
        nodeText.append(requirement.getRequirementNumber())
                .append(" ")
                .append(requirement.getRequirementName());
        
        // 添加需求类型和优先级
        if (requirement.getType() != null) {
            nodeText.append(" [").append(requirement.getType().getDescription()).append("]");
        }
        if (requirement.getPriority() != null) {
            nodeText.append(" (").append(requirement.getPriority().getDescription()).append(")");
        }
        
        run.setText(nodeText.toString());
        run.setFontSize(12);
        
        if (level == 0 || requirement.getPriority() == Requirement.Priority.HIGH) {
            run.setBold(true);
        }
        
        // 添加追溯信息
        if (traceabilityManager != null) {
            List<String> testCaseIds = traceabilityManager.getTestCaseIdsByRequirement(requirement.getRequirementId());
            if (!testCaseIds.isEmpty()) {
                XWPFParagraph tracePara = document.createParagraph();
                CTPPr tracePpr = tracePara.getCTP().isSetPPr() ? tracePara.getCTP().getPPr() : tracePara.getCTP().addNewPPr();
                CTInd traceInd = tracePpr.isSetInd() ? tracePpr.getInd() : tracePpr.addNewInd();
                traceInd.setLeft(BigInteger.valueOf((level + 1) * 400));
                
                XWPFRun traceRun = tracePara.createRun();
                traceRun.setText("    → 测试用例: " + String.join(", ", testCaseIds));
                traceRun.setFontSize(10);
                traceRun.setColor("0066CC");
            }
        }
        
        // 添加描述
        if (requirement.getDescription() != null && !requirement.getDescription().trim().isEmpty()) {
            XWPFParagraph descPara = document.createParagraph();
            CTPPr descPpr = descPara.getCTP().isSetPPr() ? descPara.getCTP().getPPr() : descPara.getCTP().addNewPPr();
            CTInd descInd = descPpr.isSetInd() ? descPpr.getInd() : descPpr.addNewInd();
            descInd.setLeft(BigInteger.valueOf((level + 1) * 400));
            
            XWPFRun descRun = descPara.createRun();
            descRun.setText("    " + requirement.getDescription());
            descRun.setFontSize(10);
            descRun.setItalic(true);
            descRun.setColor("666666");
        }
        
        // 递归添加子节点
        List<Requirement> children = requirement.getChildren();
        if (children != null && !children.isEmpty()) {
            for (int i = 0; i < children.size(); i++) {
                Requirement child = children.get(i);
                addRequirementNodeWithTraceability(document, child, level + 1, requirementManager, traceabilityManager);
            }
        }
    }
    
    /**
     * 添加追溯统计信息
     */
    private void addTraceabilityStatistics(
            XWPFDocument document,
            RequirementManager requirementManager,
            pub.developers.docautogenbyexcel.manager.TraceabilityManager traceabilityManager) {
        
        document.createParagraph();
        XWPFParagraph traceTitle = document.createParagraph();
        XWPFRun traceTitleRun = traceTitle.createRun();
        traceTitleRun.setText("追溯统计");
        traceTitleRun.setBold(true);
        traceTitleRun.setFontSize(14);
        document.createParagraph();
        
        List<Requirement> allReqs = requirementManager.getAllRequirements();
        long tracedCount = allReqs.stream()
                .map(Requirement::getRequirementId)
                .map(traceabilityManager::getTestCaseIdsByRequirement)
                .filter(list -> !list.isEmpty())
                .count();
        
        addStatisticItem(document, "已追溯需求数", tracedCount + "/" + allReqs.size());
        addStatisticItem(document, "追溯覆盖率", 
                String.format("%.2f%%", (double) tracedCount / allReqs.size() * 100));
    }
}

