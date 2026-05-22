package pub.developers.docautogenbyexcel.baseline;

import org.springframework.stereotype.Component;
import pub.developers.docautogenbyexcel.model.Requirement;

import java.util.ArrayList;
import java.util.List;

/**
 * 默认解析器：无LLM配置时基于规则提取需求。
 */
@Component
public class HeuristicRequirementParser implements LLMRequirementParser {

  @Override
  public List<Requirement> parseRequirements(String srsText) {
    List<Requirement> requirements = new ArrayList<>();
    if (srsText == null || srsText.isBlank()) {
      return requirements;
    }

    String[] lines = srsText.split("\\R");
    int seq = 1;
    for (String line : lines) {
      String text = line == null ? "" : line.trim();
      if (text.isEmpty()) {
        continue;
      }

      boolean looksLikeReq = text.matches("^(REQ[-_ ]?\\d+.*|\\d+(\\.\\d+)+.*)")
          || text.contains("应")
          || text.contains("shall");
      if (!looksLikeReq) {
        continue;
      }

      Requirement requirement = new Requirement();
      requirement.setRequirementId(String.format("REQ-%03d", seq));
      requirement.setRequirementNumber(requirement.getRequirementId());
      requirement.setRequirementName(text.length() > 30 ? text.substring(0, 30) : text);
      requirement.setDescription(text);
      requirements.add(requirement);
      seq++;
    }

    return requirements;
  }
}
