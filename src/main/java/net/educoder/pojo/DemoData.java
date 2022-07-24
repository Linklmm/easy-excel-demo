package net.educoder.pojo;

import com.alibaba.excel.annotation.ExcelProperty;
import java.util.Date;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

/**
 * 基础数据类.这里的排序和excel里面的排序一致
 **/
@Getter
@Setter
@EqualsAndHashCode
public class DemoData {

  @ExcelProperty("字符串标题")
  private String string;
  @ExcelProperty("日期标题")
  private Date date;
  @ExcelProperty("数字标题")
  private Double doubleData;
}
