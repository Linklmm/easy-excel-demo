package net.educoder.read;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.converters.DefaultConverterLoader;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.fastjson.JSON;
import com.google.common.collect.Lists;
import java.io.File;
import java.util.List;
import lombok.extern.slf4j.Slf4j;
import net.educoder.listener.ConvertDataListener;
import net.educoder.pojo.ConverterData;
import net.educoder.listener.DemoDataListener;
import net.educoder.listener.HeadDataListener;
import net.educoder.listener.IndexDataListener;
import net.educoder.listener.NameDataListener;
import net.educoder.pojo.DemoData;
import net.educoder.pojo.IndexData;
import net.educoder.pojo.NameData;
import org.junit.Test;

/**
 * @author minmin.liu
 * @version 1.0
 */
@Slf4j
public class ReadTest {

  private static final String resource = "demo" + File.separator + "demo.xlsx";

  /**
   * 2.1:最简单的读 最简单的读
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link DemoData}
   * <p>
   * 2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器
   * <p>
   * 3. 直接读即可
   */
  @Test
  public void simpleRead() {
    // 简单读写 ，第一种写法 使用匿名内部类，不用额外写一个监听器
    String fileName = this.getClass().getClassLoader().getResource(resource).getPath();
    // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
    EasyExcel.read(fileName, DemoData.class, new ReadListener<DemoData>() {

      /**
       * 临时缓存
       */
      private List<DemoData> demoDataListCache = Lists.newArrayList();

      /**
       * 分析一行数据时触发的方法
       * @param data 一行对应的数据
       * @param context 数据分析上下文
       */
      public void invoke(DemoData data, AnalysisContext context) {
        //获取
        demoDataListCache.add(data);
        saveData(data);
      }

      /**
       * 所有数据解析完成了后调用
       * @param analysisContext 数据分析上下文
       */
      public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.info("一共存入{}条数据", demoDataListCache.size());
        log.info("所有数据解析完成！");
      }

      /**
       * 模拟存入数据库
       */
      private void saveData(DemoData data) {
        log.info("解析到一条数据:{}", JSON.toJSONString(data));
        log.info("存储数据库成功！");
      }

    }).sheet().doRead();
  }

  /**
   * 2.2:最简单的读监听器
   */
  @Test
  public void listenerTest() {
    // 简单读写 ，第一种写法 使用匿名内部类，不用额外写一个监听器
    String fileName = this.getClass().getClassLoader().getResource(resource).getPath();
    // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
    EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
  }

  /**
   * 2.3：指定列的下标
   */
  @Test
  public void indexData() {
    String fileName = this.getClass().getClassLoader().getResource(resource).getPath();
    EasyExcel.read(fileName, IndexData.class, new IndexDataListener()).sheet().doRead();
  }

  /**
   * 2.4：指定列的列名
   */
  @Test
  public void nameData() {
    String fileName = this.getClass().getClassLoader().getResource(resource).getPath();
    EasyExcel.read(fileName, NameData.class, new NameDataListener()).sheet().doRead();
  }

  /**
   * 2.5: 读多个sheet 读多个或者全部sheet,这里注意一个sheet不能读取多次，多次读取需要重新读取文件
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link DemoData}
   * <p>
   * 2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link DemoDataListener}
   * <p>
   * 3. 直接读即可
   */
  @Test
  public void repeatedRead() {
    String fileName = this.getClass().getClassLoader().getResource(resource).getPath();
    // 读取全部sheet
    // 这里需要注意 DemoDataListener的doAfterAllAnalysed 会在每个sheet读取完毕后调用一次。
    // 然后所有sheet都会往同一个DemoDataListener里面写
    EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).doReadAll();

    //第一种写法 部分sheet
    log.error("读取部分sheet");
    EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet(1).doRead();

    //第二种写法 读取部分sheet
    log.error("读取部分sheet");
    try (ExcelReader excelReader = EasyExcel.read(fileName).build()) {
      ReadSheet readSheet1 = EasyExcel.readSheet(1).head(DemoData.class)
          .registerReadListener(new DemoDataListener()).build();

      excelReader.read(readSheet1);
    }
  }

  /**
   * 2.6:自定义转换格式 日期、数字或者自定义格式转换
   * <p>
   * 默认读的转换器{@link DefaultConverterLoader#loadDefaultReadConverter()}
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link ConverterData}.里面可以使用注解{@link DateTimeFormat}、{@link
   * NumberFormat}或者自定义注解
   * <p>
   * 2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link ConvertDataListener}
   * <p>
   * 3. 直接读即可
   */
  @Test
  public void convertData() {
    String fileName = this.getClass().getClassLoader().getResource(resource).getPath();
    EasyExcel.read(fileName, ConverterData.class, new ConvertDataListener()).sheet().doRead();
  }

  /**
   * 2.7:多行头
   *
   * <p>1. 创建excel对应的实体对象 参照{@link DemoData}
   * <p>2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link DemoDataListener}
   * <p>3. 设置headRowNumber参数，然后读。 这里要注意headRowNumber如果不指定，
   * 会根据你传入的class的{@link ExcelProperty#value()}里面的表头的数量来决定行数，
   * 如果不传入class则默认为1.当然你指定了headRowNumber不管是否传入class都是以你传入的为准。
   */
  @Test
  public void complexHeaderReader() {
    String fileName = this.getClass().getClassLoader().getResource(resource).getPath();
    EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet(2)
        .headRowNumber(12).doRead();
  }

  /**
   * 2.8:读取表头数据
   *
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link DemoData}
   * <p>
   * 2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照 { @link DemoHeadDataListener }
   * <p>
   * 3. 直接读即可
   */
  @Test
  public void headerRead() {
    String fileName = this.getClass().getClassLoader().getResource(resource).getPath();
    EasyExcel.read(fileName, DemoData.class, new HeadDataListener()).sheet().doRead();
  }
}
