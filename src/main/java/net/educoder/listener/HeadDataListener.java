package net.educoder.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.fastjson.JSON;
import com.google.common.collect.Lists;
import java.util.List;
import java.util.Map;
import lombok.extern.slf4j.Slf4j;
import net.educoder.pojo.DemoData;

/**
 * 模板的读取类
 */
// 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
@Slf4j
public class HeadDataListener implements ReadListener<DemoData> {

  /**
   * 缓存的数据
   */
  private List<DemoData> cachedDataList = Lists.newArrayList();


  /**
   * 在转换异常 获取其他异常下会调用本接口。抛出异常则停止读取。如果这里不抛出异常则 继续读取下一行。
   *
   * @param exception
   * @param context
   * @throws Exception
   */
  @Override
  public void onException(Exception exception, AnalysisContext context) {
    log.error("解析失败，但是继续解析下一行:{}", exception.getMessage());
    if (exception instanceof ExcelDataConvertException) {
      ExcelDataConvertException excelDataConvertException = (ExcelDataConvertException) exception;
      log.error("第{}行，第{}列解析异常，数据为:{}", excelDataConvertException.getRowIndex(),
          excelDataConvertException.getColumnIndex(), excelDataConvertException.getCellData());
    }
  }

  /**
   * 这个每一条数据解析都会来调用
   *
   * @param data    one row value. Is is same as {@link AnalysisContext#readRowHolder()}
   * @param context
   */
  @Override
  public void invoke(DemoData data, AnalysisContext context) {
    log.info("解析到一条数据:{}", JSON.toJSONString(data));
    cachedDataList.add(data);
    // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
    saveData(data);
  }

  /**
   * 这里会一行行的返回头
   *
   * @param headMap
   * @param context
   */
  @Override
  public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
    log.info("解析到一条头数据:{}", JSON.toJSONString(headMap));
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
    log.info("存储数据库成功！");
    log.info("所有数据解析完成！");
  }

  /**
   * 加上存储数据库
   */
  private void saveData(DemoData data) {
    log.info("解析到一条数据:{}", JSON.toJSONString(data));
  }
}
