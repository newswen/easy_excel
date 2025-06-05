package com.yw.easy_excel_test.simple;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.yw.easy_excel_test.utils.CommonListener;

import java.util.List;

public class SimpleReadDemo {
    public static void readDemo() {
        String fileName = "竹子数据-1746684382081.xlsx";
        ZhuZiListener zhuZiListener = new ZhuZiListener();
        // 读取指定路径的文件，并转换为ZhuZi对象，默认会读取第一个sheet单元
        EasyExcel.read(fileName, ZhuZi.class, zhuZiListener).sheet().doRead();
        List<ZhuZi> zhuZis = zhuZiListener.getData();
        System.out.println("读取excel文件结束，总计解析到" + zhuZis.size() + "条数据！");
    }

    public static void readDemo2() {
        String fileName = "竹子数据-1746684382081.xlsx";
        CommonListener<ZhuZi> listener = new CommonListener<>(ZhuZi.class);
        EasyExcelFactory.read(fileName, ZhuZi.class, listener).sheet().doRead();
        List<ZhuZi> zhuZis = listener.getData();
        System.out.println(zhuZis);
        System.out.println("读取excel文件结束，总计解析到" + zhuZis.size() + "条数据！");
    }

    public static void main(String[] args) {
      readDemo2();
    }
}
