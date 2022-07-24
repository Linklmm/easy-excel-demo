package net.educoder.action;

import java.io.File;

public class SimpleRead {

  private static final String fileName =
      SimpleRead.class.getResource("/").getPath()
          + File.separator
          + "action"
          + File.separator
          + "demo.xlsx";

  public static void main(String[] args) {

  }
}
