package net.educoder.write;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.google.common.collect.Lists;
import net.educoder.pojo.DemoData;
import net.educoder.pojo.WriteIndexData;
import net.educoder.pojo.WriteOrderData;
import org.junit.Test;

import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

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
        EasyExcel.write(fileName, DemoData.class).excludeColumnFieldNames(excludeColumnFieldNames).sheet("忽略字段")
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
