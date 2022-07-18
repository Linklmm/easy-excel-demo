package net.educoder.write;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.google.common.collect.Lists;
import net.educoder.pojo.DemoData;
import org.junit.Test;

import java.util.Date;
import java.util.List;

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
        String fileName = path + "simpleWrite01" + ".xlsx";
        EasyExcel.write(fileName, DemoData.class)
                .sheet("简单写")
                .doWrite(this::data);
        //写法2
        fileName = path + "simpleWrite02" + ".xlsx";
        EasyExcel.write(fileName, DemoData.class).sheet("写法2").doWrite(data());
        //写法3
        fileName = path + "simpleWrite03" + ".xlsx";
        try (ExcelWriter excelWriter = EasyExcel.write(fileName, DemoData.class).build()) {
            WriteSheet writeSheet = EasyExcel.writerSheet("写法3").build();
            excelWriter.write(data(), writeSheet);
        }
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
}
