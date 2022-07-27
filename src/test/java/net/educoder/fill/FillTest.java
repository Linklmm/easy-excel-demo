package net.educoder.fill;

import com.alibaba.excel.EasyExcel;
import com.google.common.collect.Maps;
import java.io.File;
import java.util.Map;
import net.educoder.pojo.FillData;
import net.educoder.write.WriteTest;
import org.junit.Test;

public class FillTest {

  private static String path = WriteTest.class.getResource("/").getPath();

  /**
   * 4.1 最简单的填充
   *
   * @since 2.1.1
   */
  @Test
  public void simpleFill() {
    // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
    String templateFileName =
        path + "fill" + File.separator + "template" + File.separator + "simple.xlsx";
    String fileName =
        path + "fill" + File.separator + "simple01.xlsx";
    //方案1 根据对象填充
    FillData fillData = new FillData();
    fillData.setName("张三");
    fillData.setNumber(5.2);
    EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(fillData);

    //方案2 根据Map填充
    fileName = path + "fill" + File.separator + "simple02.xlsx";
    Map<String, Object> map = Maps.newHashMap();
    map.put("name", "张三");
    map.put("number", 5.2);
    EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(map);
  }
  /**
   * 4.2填充列表
   *
   * @since 2.1.1
   */
  @Test
  public void listFill(){
    String templateFileName =
        path + "fill" + File.separator + "template" + File.separator + "complexFillWithTable.xlsx";
    String fileName =
        path + "fill" + File.separator + "complexFillWithTable01.xlsx";
  }
}
