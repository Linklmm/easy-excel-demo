package net.educoder.pojo;

import com.alibaba.excel.annotation.ExcelProperty;
import java.util.Date;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

/**
 * 指定列名
 */
@Getter
@Setter
@EqualsAndHashCode
public class NameData {

  /**
   * 强制读取第三个，index从0开始
   */
  @ExcelProperty(value = "字符串标题")
  private String string;

  @ExcelProperty(value = "日期标题")
  private Date date;
}
