package net.educoder.write;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.data.CommentData;
import com.alibaba.excel.metadata.data.FormulaData;
import com.alibaba.excel.metadata.data.HyperlinkData;
import com.alibaba.excel.metadata.data.HyperlinkData.HyperlinkType;
import com.alibaba.excel.metadata.data.ImageData;
import com.alibaba.excel.metadata.data.ImageData.ImageType;
import com.alibaba.excel.metadata.data.RichTextStringData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.util.FileUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.google.common.collect.Lists;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import net.educoder.pojo.ComplexHeadData;
import net.educoder.pojo.ConverterData;
import net.educoder.pojo.DemoData;
import net.educoder.pojo.ImageDemoData;
import net.educoder.pojo.WriteCellDemoData;
import net.educoder.pojo.WriteIndexData;
import net.educoder.pojo.WriteOrderData;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;

/**
 * @author minmin.liu
 * @version 1.0
 */
public class WriteTest {

  private static final String path = WriteTest.class.getResource("/").getPath();

  /**
   * 3.1 最简单的写
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link DemoData}
   * <p>
   * 2. 直接写即可
   */
  @Test
  public void simpleWrite() {
    //写法1
    String fileName = path + "simpleWrite01.xlsx";
    EasyExcel.write(fileName, DemoData.class)
        .sheet("简单写")
        .doWrite(this::data);
    //写法2
    fileName = path + "simpleWrite02.xlsx";
    EasyExcel.write(fileName, DemoData.class).sheet("写法2").doWrite(data());
    //写法3
    fileName = path + "simpleWrite03.xlsx";
    try (ExcelWriter excelWriter = EasyExcel.write(fileName, DemoData.class).build()) {
      WriteSheet writeSheet = EasyExcel.writerSheet("写法3").build();
      excelWriter.write(data(), writeSheet);
    }
  }

  /**
   * 3.2:根据参数只导出指定列
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link DemoData}
   * <p>
   * 2. 根据自己或者排除自己需要的列
   * <p>
   * 3. 直接写即可
   *
   * @since 2.1.1
   */
  @Test
  public void excludeOrIncludeWrite() {
    //忽略某个字段
    String fileName = path + "excludeWrite.xlsx";
    // 根据用户传入字段 假设我们要忽略 date
    Set<String> excludeColumnFieldNames = new HashSet<>();
    excludeColumnFieldNames.add("date");
    // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为"忽略字段" 然后文件流会自动关闭
    EasyExcel.write(fileName, DemoData.class).excludeColumnFieldNames(excludeColumnFieldNames)
        .sheet("忽略字段")
        .doWrite(data());

    //只写入某个字段
    fileName = path + "includeWrite.xlsx";
    Set<String> includeColumnFieldNames = new HashSet<>();
    includeColumnFieldNames.add("string");
    EasyExcel.write(fileName, DemoData.class)
        .includeColumnFieldNames(includeColumnFieldNames)
        .sheet("只写入某些字段")
        .doWrite(data());
  }

  /**
   * 3.3 指定写入的列
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link WriteIndexData}
   * <p>
   * 2. 使用{@link ExcelProperty}注解指定写入的列
   * <p>
   * 3. 直接写即可
   */
  @Test
  public void indexWrite() {
    String fileName = path + "indexWrite.xlsx";
    //会有空列
    EasyExcel.write(fileName, WriteIndexData.class).sheet("指定写入的列").doWrite(writeIndexData());
    //不会有空列
    fileName = path + "orderWrite.xlsx";
    EasyExcel.write(fileName, WriteOrderData.class).sheet("指定写入的列").doWrite(writeOrderData());

  }

  /**
   * 3.4 复杂头写入
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link ComplexHeadData}
   * <p>
   * 2. 使用{@link ExcelProperty}注解指定复杂的头
   * <p>
   * 3. 直接写即可
   */
  @Test
  public void complexHeadWrite() {
    String fileName = path + "complexHeadWrite.xlsx";
    EasyExcel.write(fileName, ComplexHeadData.class).sheet("复杂的头")
        .doWrite(data());
  }

  /**
   * 3.5 重复多次写入
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link ComplexHeadData}
   * <p>
   * 2. 使用{@link ExcelProperty}注解指定复杂的头
   * <p>
   * 3. 直接调用二次写入即可
   */
  @Test
  public void repeatedWrite() {
    //方法1：写到同一个sheet
    String fileName = path + "repeatedWrite01.xlsx";
    //指定class
    try (ExcelWriter excelWriter = EasyExcel.write(fileName, DemoData.class).build()) {
      WriteSheet writeSheet = EasyExcel.writerSheet("重复写").build();
      // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来
      for (int i = 0; i < 5; i++) {
        // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
        List<DemoData> data = data();
        excelWriter.write(data, writeSheet);
      }
    }

    // 方法2: 如果写到不同的sheet 同一个对象
    fileName = path + "repeatedWrite02.xlsx";
    // 这里 指定文件
    try (ExcelWriter excelWriter2 = EasyExcel.write(fileName, DemoData.class).build()) {
      // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
      for (int i = 0; i < 5; i++) {
        // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样
        WriteSheet writeSheet2 = EasyExcel.writerSheet(i, "模板" + i).build();
        // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
        List<DemoData> data = data();
        excelWriter2.write(data, writeSheet2);
      }
    }

    //方法3:写到不同的sheet 不同的对象
    fileName = path + "repeatedWrite03.xlsx";
    try (ExcelWriter excelWriter3 = EasyExcel.write(fileName).build()) {
      for (int i = 0; i < 5; i++) {
        // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样。
        // 这里注意DemoData.class 可以每次都变，我这里为了方便 所以用的同一个class
        // 实际上可以一直变

        // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
        if (i == 2) {
          WriteSheet writeSheet3 = EasyExcel.writerSheet(i, "模板" + i).head(WriteIndexData.class)
              .build();
          List<WriteIndexData> writeIndexData = writeIndexData();
          excelWriter3.write(writeIndexData, writeSheet3);

        } else {
          WriteSheet writeSheet3 = EasyExcel.writerSheet(i, "模板" + i).head(DemoData.class).build();
          List<DemoData> data = data();
          excelWriter3.write(data, writeSheet3);
        }
      }
    }

  }

  /**
   * 3.6 日期、数字或者自定义格式转换
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link ConverterData}
   * <p>
   * 2. 使用{@link ExcelProperty}配合使用注解{@link DateTimeFormat}、{@link NumberFormat}或者自定义注解
   * <p>
   * 3. 直接写即可
   */
  @Test
  public void convertWrite() {
    String fileName = path + "converterWrite.xlsx";
    EasyExcel.write(fileName, ConverterData.class).sheet("模板").doWrite(data());
  }

  /**
   * 3.7 图片导出
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link ImageDemoData}
   * <p>
   * 2. 直接写即可
   */
  @Test
  public void imageWrite() throws IOException {
    String fileName = path + "imageWrite.xlsx";
    String imagePath = path + "imgs" + File.separator + "img.jpg";
    List<ImageDemoData> list = Lists.newArrayList();
    try (InputStream inputStream = FileUtils.openInputStream(new File(imagePath))) {
      ImageDemoData imageDemoData = new ImageDemoData();
      // 放入五种类型的图片 实际使用只要选一种即可
      imageDemoData.setByteArray(FileUtils.readFileToByteArray(new File(imagePath)));
      imageDemoData.setFile(new File(imagePath));
      imageDemoData.setString(imagePath);
      imageDemoData.setInputStream(inputStream);
      imageDemoData.setUrl(new URL(
          "https://lmm.myfloweryourgrass.cn/blog/logo.png"));

      // 这里演示
      // 需要额外放入文字
      // 而且需要放入2个图片
      // 第一个图片靠左
      // 第二个靠右 而且要额外的占用他后面的单元格
      WriteCellData<Void> writeCellData = new WriteCellData<>();
      // 这里可以设置为 EMPTY 则代表不需要其他数据了
      writeCellData.setType(CellDataTypeEnum.STRING);
      writeCellData.setStringValue("额外的放一些文字");

      // 可以放入多个图片
      List<ImageData> imageDataList = Lists.newArrayList();
      ImageData imageData = new ImageData();
      // 放入2进制图片
      imageData.setImage(FileUtils.readFileToByteArray(new File(imagePath)));
      // 图片类型
      imageData.setImageType(ImageType.PICTURE_TYPE_PNG);
      // 上 右 下 左 需要留空
      // 这个类似于 css 的 margin
      // 这里实测 不能设置太大 超过单元格原始大小后 打开会提示修复。暂时未找到很好的解法。
      imageData.setTop(5);
      imageData.setRight(40);
      imageData.setBottom(5);
      imageData.setLeft(5);

      // 放入第二个图片
      ImageData imageData2 = new ImageData();
      imageData2.setImage(FileUtils.readFileToByteArray(new File(imagePath)));
      imageData2.setImageType(ImageType.PICTURE_TYPE_PNG);
      imageData2.setTop(5);
      imageData2.setRight(5);
      imageData2.setBottom(5);
      imageData2.setLeft(50);

      // 设置图片的位置 假设 现在目标 是 覆盖 当前单元格 和当前单元格右边的单元格
      // 起点相对于当前单元格为0 当然可以不写
      imageData2.setRelativeFirstRowIndex(0);
      imageData2.setRelativeFirstColumnIndex(0);
      imageData2.setRelativeLastRowIndex(0);
      // 前面3个可以不写  下面这个需要写 也就是 结尾 需要相对当前单元格 往右移动一格
      // 也就是说 这个图片会覆盖当前单元格和 后面的那一格
      imageData2.setRelativeLastColumnIndex(1);

      imageDataList.add(imageData);
      imageDataList.add(imageData2);
      writeCellData.setImageDataList(imageDataList);

      imageDemoData.setWriteCellDataFile(writeCellData);
      list.add(imageDemoData);

      // 写入数据
      EasyExcel.write(fileName, ImageDemoData.class).sheet().doWrite(list);
    }


  }

  /**
   * 3.8 超链接、备注、公式、指定单个单元格的样式、单个单元格多种样式
   * <p>
   * 1. 创建excel对应的实体对象 参照{@link WriteCellDemoData}
   * <p>
   * 2. 直接写即可
   *
   * @since 3.0.0-beta1
   */
  @Test
  public void writeCellDataWrite() {
    String fileName = path + "writeCellDataWrite.xlsx";
    WriteCellDemoData writeCellDemoData = new WriteCellDemoData();

    //设置超链接
    WriteCellData<String> hyperlink = new WriteCellData<>("官方网站");
    HyperlinkData hyperlinkData = new HyperlinkData();
    hyperlinkData.setAddress("https://www.educoder.net");
    hyperlinkData.setHyperlinkType(HyperlinkType.URL);

    hyperlink.setHyperlinkData(hyperlinkData);
    writeCellDemoData.setHyperlink(hyperlink);

    //设置备注
    WriteCellData<String> comment = new WriteCellData<>("备注信息");
    CommentData commentData = new CommentData();
    commentData.setAuthor("playboy");
    commentData.setRichTextStringData(new RichTextStringData("这是一个备注"));
    commentData.setRelativeLastColumnIndex(1);
    commentData.setRelativeLastRowIndex(1);
    comment.setCommentData(commentData);
    writeCellDemoData.setCommentData(comment);

    //设置公式
    WriteCellData<String> formula = new WriteCellData<>("设置公式");
    FormulaData formulaData = new FormulaData();
    // 将 123456789 中的第一个数字替换成 2
    // 这里只是例子 如果真的涉及到公式 能内存算好尽量内存算好 公式能不用尽量不用
    formulaData.setFormulaValue("REPLACE(123456789,1,1,2)");
    formula.setFormulaData(formulaData);
    writeCellDemoData.setFormulaData(formula);

    // 设置单个单元格的样式 当然样式 很多的话 也可以用注解等方式。
    // writeCellDemoData-> WriteCellData<String> ->  WriteCellStyle
    WriteCellData<String> writeCellStyle = new WriteCellData<>("设置单元格样式");
    writeCellStyle.setType(CellDataTypeEnum.STRING);
    WriteCellStyle writeCellStyleData = new WriteCellStyle();
    // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.
    writeCellStyleData.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
    // 背景绿色
    writeCellStyleData.setFillForegroundColor(IndexedColors.GREEN.getIndex());
    //设置样式
    writeCellStyle.setWriteCellStyle(writeCellStyleData);
    writeCellDemoData.setWriteCellStyle(writeCellStyle);

    // 设置单个单元格多种样式
    WriteCellData<String> richTest = new WriteCellData<>("设置单个单元格多种样式");
    richTest.setType(CellDataTypeEnum.RICH_TEXT_STRING);
    RichTextStringData richTextStringData = new RichTextStringData();
    richTextStringData.setTextString("红色绿色默认");
    // 前2个字红色
    WriteFont writeFont = new WriteFont();
    writeFont.setColor(IndexedColors.RED.getIndex());
    richTextStringData.applyFont(0, 2, writeFont);
    // 接下来2个字绿色
    WriteFont writeFont2 = new WriteFont();
    writeFont2.setColor(IndexedColors.GREEN.getIndex());
    richTextStringData.applyFont(2, 4, writeFont2);

    richTest.setRichTextStringDataValue(richTextStringData);

    writeCellDemoData.setRichText(richTest);

    List<WriteCellDemoData> data = Lists.newArrayList();
    data.add(writeCellDemoData);
    EasyExcel.write(fileName, WriteCellDemoData.class)
        .inMemory(true)
        .sheet("设置单元格")
        .doWrite(data);

  }

  /**
   * 3.9 根据模板写入
   * <p>
   * 1. 创建excel对应的实体对象
   * <p>
   * 2. 使用{@link ExcelProperty}注解指定写入的列
   * <p>
   * 3. 使用withTemplate 写取模板
   * <p>
   * 4. 直接写即可
   */
  @Test
  public void templateWrite() {
    String templateFileName = path + "demo" + File.separator + "demo.xlsx";
    String fileName = path + "templateWrite.xlsx";
    // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
    EasyExcel.write(fileName, DemoData.class).withTemplate(templateFileName).sheet()
        .doWrite(data());
  }

  private List<DemoData> data() {
    List<DemoData> datas = Lists.newArrayList();
    for (int i = 0; i < 10; i++) {
      DemoData data = new DemoData();
      data.setString("数据" + i);
      data.setDate(new Date());
      data.setDoubleData(i * Math.random());
      datas.add(data);
    }
    return datas;
  }

  private List<WriteIndexData> writeIndexData() {
    List<WriteIndexData> datas = Lists.newArrayList();
    for (int i = 0; i < 10; i++) {
      WriteIndexData data = new WriteIndexData();
      data.setString("数据" + i);
      data.setDate(new Date());
      data.setDoubleData(i * Math.random());
      //data.setIntData(i);
      datas.add(data);
    }
    return datas;
  }

  private List<WriteOrderData> writeOrderData() {
    List<WriteOrderData> datas = Lists.newArrayList();
    for (int i = 0; i < 10; i++) {
      WriteOrderData data = new WriteOrderData();
      data.setString("数据" + i);
      data.setDate(new Date());
      //data.setDoubleData(i * Math.random());
      data.setIntData(i);
      datas.add(data);
    }
    return datas;
  }
}
