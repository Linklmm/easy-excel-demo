package net.educoder.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.fastjson.JSON;
import com.google.common.collect.Lists;
import java.util.List;
import lombok.extern.slf4j.Slf4j;
import net.educoder.pojo.IndexData;
import net.educoder.pojo.NameData;

/**
 * 模板的读取类
 *
 */
// 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
@Slf4j
public class NameDataListener implements ReadListener<NameData> {

  /**
   * 缓存的数据
   */
  private List<NameData> cachedDataList = Lists.newArrayList();


  /**
   * 这个每一条数据解析都会来调用
   *
   * @param data    one row value. Is is same as {@link AnalysisContext#readRowHolder()}
   * @param context
   */
  @Override
  public void invoke(NameData data, AnalysisContext context) {
    log.info("解析到一条数据:{}", JSON.toJSONString(data));
    cachedDataList.add(data);
    // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
    saveData(data);
  }

  /**
   * 所有数据解析完成了 都会来调用
   *
   * @param context
   */
  @Override
  public void doAfterAllAnalysed(AnalysisContext context) {
    // 这里也要保存数据，确保最后遗留的数据也存储到数据库
    log.info("一共存入{}条数据", cachedDataList.size());
    log.info("所有数据解析完成！");
  }

  /**
   * 加上存储数据库
   */
  private void saveData(NameData data) {
    log.info("解析到一条数据:{}", JSON.toJSONString(data));
    log.info("存储数据库成功！");
  }
}
