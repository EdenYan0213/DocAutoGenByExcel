package pub.developers.docautogenbyexcel.writer;

/**
 * 数据写入层：需求基线构建器
 * 继承AbstractDataWriter，封装validate() → write()模板方法
 *
 * 负责调用LLM解析SRS（软件需求规格说明），建立需求基线
 * 当前为架构预留类，待LLM集成后实现完整功能
 */
public class RequirementBaselineBuilder extends AbstractDataWriter {

  /**
   * 校验SRS文档是否合法
   * 检查文件存在性、格式、可读性
   */
  @Override
  protected void validate(String sourcePath) throws Exception {
    if (sourcePath == null || sourcePath.trim().isEmpty()) {
      throw new Exception("SRS文档路径不能为空");
    }
    java.io.File file = new java.io.File(sourcePath);
    if (!file.exists()) {
      throw new Exception("SRS文档不存在: " + sourcePath);
    }
    if (!file.canRead()) {
      throw new Exception("无法读取SRS文档: " + sourcePath);
    }
  }

  /**
   * 解析SRS文档，建立需求基线
   * 预留LLM集成接口，当前输出提示信息
   */
  @Override
  protected void write(String sourcePath) throws Exception {
    // TODO: 集成LLM解析SRS文档，自动提取需求并建立需求基线
    System.out.println("RequirementBaselineBuilder: 需求基线构建功能待LLM集成后实现");
    System.out.println("  SRS文档路径: " + sourcePath);
  }
}