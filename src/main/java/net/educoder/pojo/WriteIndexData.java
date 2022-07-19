package net.educoder.pojo;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

import java.util.Date;

/**
 * 基础数据类
 *
 * @author Jiaju Zhuang
 **/
@Getter
@Setter
@EqualsAndHashCode
public class WriteIndexData {
    @ExcelProperty(value = "字符串标题", index = 0)
    private String string;
    @ExcelProperty(value = "日期标题", index = 1)
    private Date date;
    /**
     * 这里设置3 会导致第二列空的
     */
    @ExcelProperty(value = "数字标题", index = 3)
    private Double doubleData;

    // 这里需要注意 在使用ExcelProperty注解的使用，如果想不空列则需要加入order字段，
    // 而不是index,order会忽略空列，然后继续往后，而index，不会忽略空列，在第几列就是第几列。
    //@ExcelProperty(value = "整数标题", order = 5)
    //private Integer intData;
}
