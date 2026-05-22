package pub.developers.docautogenbyexcel.builder;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.processor.TableFillProcessor;
import pub.developers.docautogenbyexcel.processor.WordProcessor;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Map;

/**
 * Word document construction facade.
 * Encapsulates section generation and table filling operations.
 */
public class WordDocumentBuilder {

    private final WordProcessor wordProcessor;
    private final TableFillProcessor tableFillProcessor;

    public WordDocumentBuilder() {
        this.wordProcessor = new WordProcessor();
        this.tableFillProcessor = new TableFillProcessor();
    }

    public int buildModuleSections(String templatePath, String outputPath,
                                   Map<String, ModuleData> moduleDataMap) throws Exception {
        return wordProcessor.processWord(templatePath, outputPath, moduleDataMap);
    }

    public void fillAdditionalTables(String outputPath,
                                     Map<String, BasicInfoData> basicInfoMap,
                                     Map<String, ListTableData> listTableMap,
                                     Map<String, ModuleData> moduleDataMap) throws Exception {
        // 即使basicInfoMap和listTableMap为空，只要moduleDataMap不为空就继续处理追踪表
        boolean hasBasicOrList = (basicInfoMap != null && !basicInfoMap.isEmpty())
                || (listTableMap != null && !listTableMap.isEmpty());
        boolean hasModule = moduleDataMap != null && !moduleDataMap.isEmpty();
        
        if (!hasBasicOrList && !hasModule) {
            return;
        }

        try (FileInputStream fis = new FileInputStream(outputPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            if (hasBasicOrList) {
                tableFillProcessor.fillBasicInfoTables(document, basicInfoMap);
                tableFillProcessor.fillListTables(document, listTableMap);
            }
            
            // 填充测试项追踪表（表9.1 测试依据到测试项的追踪）
            if (hasModule) {
                tableFillProcessor.fillTestTraceabilityTables(document, moduleDataMap);
            }

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }
        }
    }
}
