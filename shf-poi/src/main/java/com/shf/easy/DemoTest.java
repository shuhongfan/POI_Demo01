package com.shf.easy;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class DemoTest {
    String PATH = "C:\\Users\\shuho\\Documents\\Code\\POI_Demo01\\shf-poi\\src\\main\\java\\com\\shf\\";

    private List<DemoData> data(){
        ArrayList<DemoData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串："+i);
            data.setDate(new Date());
            data.setDoubleDate(0.56);
            list.add(data);
        }
        return list;
    }

//    根据list  写入excel
    @Test
    public void simpleWrite() {
        String fileName = PATH+"EasyTest.xlsx";
//        String fileName = TestFileUtil.getPath() + "write" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去读，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
//        write(fileName,格式类)
//        sheet(表明)
//        doWrite(数据)
        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(data());
    }

    /**
     * 不创建对象的读
     */
    @Test
    public void noModelRead() {
        String fileName = PATH+"EasyTest.xlsx";
        EasyExcel.read(fileName,DemoData.class, new DemoDataListener()).sheet().doRead();
    }
}
