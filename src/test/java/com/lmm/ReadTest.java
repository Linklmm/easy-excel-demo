package com.lmm;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

import java.io.File;

/**
 * @author minmin.liu
 * @version 1.0
 */
public class ReadTest {
    @Test
    public void simpleRead() {
        String fileName = TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
    }
}
