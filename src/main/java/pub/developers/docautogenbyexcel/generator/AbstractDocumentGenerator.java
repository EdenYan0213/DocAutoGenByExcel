package pub.developers.docautogenbyexcel.generator;

import pub.developers.docautogenbyexcel.builder.WordDocumentBuilder;
import pub.developers.docautogenbyexcel.hub.DataHub;

/**
 * 抽象生成层基类
 * 封装 extractData() → generateContent() → save() 模板方法
 *
 * 所有文档生成器都应继承此基类，
 * 确保生成流程遵循统一的数据提取、内容生成、保存三步流程。
 */
public abstract class AbstractDocumentGenerator {

  protected final DataHub dataHub;
  protected final WordDocumentBuilder wordDocumentBuilder;

  protected AbstractDocumentGenerator(DataHub dataHub) {
    this.dataHub = dataHub;
    this.wordDocumentBuilder = new WordDocumentBuilder();
  }

  /**
   * 模板方法：提取数据 → 生成内容 → 保存
   * 不可被子类覆盖，确保执行顺序为 extractData → generateContent → save
   *
   * @param excelPath    Excel数据源路径
   * @param templatePath Word模板路径
   * @param outputPath   输出文件路径
   * @return 生成结果
   * @throws Exception 任何步骤失败时抛出
   */
  public final GenerateResult generate(String excelPath, String templatePath, String outputPath) throws Exception {
    ExtractedData extractedData = extractData(excelPath);
    int contentResult = generateContent(templatePath, outputPath, extractedData);
    return save(outputPath, extractedData, contentResult);
  }

  /**
   * 第一步：从数据源提取所需数据
   * 子类实现具体的数据提取逻辑
   *
   * @param excelPath 数据源路径
   * @return 提取的数据容器
   * @throws Exception 提取失败时抛出
   */
  protected abstract ExtractedData extractData(String excelPath) throws Exception;

  /**
   * 第二步：基于提取的数据生成文档内容
   * 子类实现具体的文档内容生成逻辑
   *
   * @param templatePath  模板文件路径
   * @param outputPath    输出文件路径
   * @param extractedData 提取的数据
   * @return 生成的内容统计（如模块数量）
   * @throws Exception 生成失败时抛出
   */
  protected abstract int generateContent(String templatePath, String outputPath, ExtractedData extractedData)
      throws Exception;

  /**
   * 第三步：保存最终结果并返回生成结果
   * 默认实现直接返回结果，子类可覆盖以添加后处理（如追加表格）
   *
   * @param outputPath    输出文件路径
   * @param extractedData 提取的数据
   * @param contentResult 生成的内容统计
   * @return 最终生成结果
   * @throws Exception 保存失败时抛出
   */
  protected GenerateResult save(String outputPath, ExtractedData extractedData, int contentResult) throws Exception {
    return new GenerateResult(contentResult);
  }

  /**
   * 提取数据的容器接口
   * 子类可定义自己的数据容器实现
   */
  public interface ExtractedData {
  }

  /**
   * 生成结果记录
   */
  public record GenerateResult(int moduleCount) {
  }
}