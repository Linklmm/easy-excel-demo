package net.educoder.pojo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import java.util.Date;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import net.educoder.convert.CustomStringStringConverter;

/**
 * 基础数据类.这里的排序和excel里面的排序一致
 **/
@Getter
@Setter
@EqualsAndHashCode
public class ConverterData {

  /**
   * 我自定义 转换器，不管数据库传过来什么 。我给他加上“自定义：”
   */
  @ExcelProperty(value = "字符串标题", converter = CustomStringStringConverter.class)
  private String string;
  /**
   * 这里用string 去接日期才能格式化。我想接收年月日格式
   */
  @DateTimeFormat("yyyy年MM月dd日HH时mm分ss秒")
  @ExcelProperty("日期标题")
  private Date date;
  /**
   * 我想接收百分比的数字
   */
  @NumberFormat("#.##%")
  @ExcelProperty(value = "数字标题")
  private Double doubleData;
}