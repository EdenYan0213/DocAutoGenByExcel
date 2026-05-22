package pub.developers.docautogenbyexcel.baseline;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.stereotype.Service;
import pub.developers.docautogenbyexcel.hub.DataHub;
import pub.developers.docautogenbyexcel.model.Requirement;

import java.io.FileInputStream;
import java.util.List;

/**
 * 需求基线建立组件。
 */
@Service
public class RequirementBaselineBuilder {

  private final DataHub dataHub;
  private final LLMRequirementParser parser;

  public RequirementBaselineBuilder(DataHub dataHub, LLMRequirementParser parser) {
    this.dataHub = dataHub;
    this.parser = parser;
  }

  public int buildBaseline(String srsDocxPath, String excelHubPath) throws Exception {
    String srsText = readWordText(srsDocxPath);
    List<Requirement> requirements = parser.parseRequirements(srsText);
    dataHub.writeRequirements(excelHubPath, requirements);
    return requirements.size();
  }

  private String readWordText(String filePath) throws Exception {
    StringBuilder sb = new StringBuilder();
    try (FileInputStream fis = new FileInputStream(filePath);
        XWPFDocument document = new XWPFDocument(fis)) {
      for (XWPFParagraph paragraph : document.getParagraphs()) {
        String text = paragraph.getText();
        if (text != null && !text.isBlank()) {
          sb.append(text).append(System.lineSeparator());
        }
      }
    }
    return sb.toString();
  }
}
