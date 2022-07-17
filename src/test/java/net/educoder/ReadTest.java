package net.educoder;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.fastjson.JSON;
import com.google.common.collect.Lists;
import java.io.File;
import java.util.List;
import lombok.extern.slf4j.Slf4j;
import net.educoder.pojo.DemoData;
import org.junit.Test;

/**
 * @author minmin.liu
 * @version 1.0
 */
@Slf4j
public class ReadTest {

  private static final String resource = "demo" + File.separator + "demo.xlsx";

  /**
   * 1.1:简单读写
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
}
