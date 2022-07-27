package net.educoder.action;

import com.alibaba.fastjson.JSON;
import java.io.File;
import net.educoder.pojo.DemoData;

public class SimpleRead {

  private static final String fileName =
      SimpleRead.class.getResource("/").getPath()
          + File.separator
          + "action"
          + File.separator
          + "demo.xlsx";

  public static void main(String[] args) {
    int size = 0;
    //这里只是模拟从Excel中读取想关数据。
    for (int i = 0; i < 5; i++) {
      DemoData data = new DemoData();
      System.out.println("第" + i + "条数据" + JSON.toJSONString(data));
      size++;
    }
    System.out.println("一共存入" + size + "条数据");
  }
}
