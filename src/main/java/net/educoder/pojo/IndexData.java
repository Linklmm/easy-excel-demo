package net.educoder.pojo;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

/**
 * 指定下标
 */
@Getter
@Setter
@EqualsAndHashCode
public class IndexData {

  /**
   * 强制读取第三个，index从0开始
   */
  @ExcelProperty(index = 2)
  private double doubleData;
}
