package pub.developers.docautogenbyexcel.baseline;

import pub.developers.docautogenbyexcel.model.Requirement;

import java.util.List;

/**
 * 大语言模型需求解析抽象。
 */
public interface LLMRequirementParser {
  List<Requirement> parseRequirements(String srsText);
}
