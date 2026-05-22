package pub.developers.docautogenbyexcel.writer;

/**
 * 抽象写入层基类
 * 封装 validate() → write() 模板方法
 * 
 * 所有数据写入/导入操作都应继承此基类，
 * 确保数据在写入前经过校验。
 */
public abstract class AbstractDataWriter {

  /**
   * 模板方法：校验后写入
   * 不可被子类覆盖，确保执行顺序为 validate → write
   *
   * @param sourcePath 数据源路径（如Excel文件路径）
   * @throws Exception 校验或写入失败时抛出
   */
  public final void execute(String sourcePath) throws Exception {
    validate(sourcePath);
    write(sourcePath);
  }

  /**
   * 校验数据源是否合法
   * 子类必须实现具体的校验逻辑
   *
   * @param sourcePath 数据源路径
   * @throws Exception 校验失败时抛出
   */
  protected abstract void validate(String sourcePath) throws Exception;

  /**
   * 执行数据写入操作
   * 子类必须实现具体的写入逻辑
   *
   * @param sourcePath 数据源路径
   * @throws Exception 写入失败时抛出
   */
  protected abstract void write(String sourcePath) throws Exception;
}