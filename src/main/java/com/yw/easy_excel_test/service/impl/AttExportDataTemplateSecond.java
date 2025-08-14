package com.yw.easy_excel_test.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.yw.easy_excel_test.entity.ReasonRowVO;
import com.yw.easy_excel_test.utils.CellMergeStrategy;
import com.yw.easy_excel_test.utils.CenterCellStyleHandler;
import org.apache.commons.compress.utils.Sets;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.List;
import java.util.Random;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class AttExportDataTemplateSecond {

    public static void main(String[] args) {

        String writeFileName = "att-second-" + System.currentTimeMillis() + ".xlsx";
//        EasyExcel.write(writeFileName)
//                .head(createDynamicHead())
//                .registerWriteHandler(new CellMergeStrategy(0, 0, Sets.newHashSet(0, 1, 2)))
//                .excelType(ExcelTypeEnum.XLSX)
//                .sheet("考勤报表")
//                .doWrite(createDataByList());

//        EasyExcel.write(writeFileName)
//                .head(createAsinHead())
//                .registerWriteHandler(defaultStyleStrategy())
//                .registerWriteHandler(new CellMergeStrategy(0, 0, Sets.newHashSet(0, 1)))
////                .registerWriteHandler(new CenterCellStyleHandler()) // 新加的居中处理器
//                .excelType(ExcelTypeEnum.XLSX)
//                .sheet("分析&总结")
//                .doWrite(createAsinDate());

        ExcelWriter excelWriter = EasyExcel.write(writeFileName)
                .excelType(ExcelTypeEnum.XLSX)
                .build();

        WriteSheet writeSheet = EasyExcel.writerSheet("分析&总结")
                //自动调整列宽
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                .needHead(Boolean.FALSE)
                .build();

        WriteTable accountInfo = EasyExcel.writerTable(0)
                .head(createAsinHead())
                .needHead(Boolean.TRUE)
                .registerWriteHandler(defaultStyleStrategy())
                .registerWriteHandler(new CellMergeStrategy(0, 0, Sets.newHashSet(0, 1)))
//                .registerWriteHandler(new CenterCellStyleHandler()) // 新加的居中处理器
                .build();
        WriteTable tweetInfo = EasyExcel.writerTable(1)
                .head(ReasonRowVO.class)
                .needHead(Boolean.TRUE)
                .relativeHeadRowIndex(4)
                .build();
        List<List<Object>> tweetExcelDTOS = ListUtils.newArrayList();
        excelWriter.write(createAsinDate(), writeSheet, accountInfo);
        excelWriter.write(tweetExcelDTOS, writeSheet, tweetInfo);
        excelWriter.finish();
    }

    private static List<List<Object>> createAsinDate() {
        List<List<Object>> list = ListUtils.newArrayList();


        List<Object> data = ListUtils.newArrayList();
        data.add("US");
        data.add("下单-拣货（100%正常）");
        for (int i = 0; i < 3; i++) {
            data.add("");
        }
        list.add(data);
        for (int i = 0; i < 3; i++) {
            List<Object> data1 = ListUtils.newArrayList();
            data1.add("US");
            data1.add("拣货-扫描（20%扫描慢）");
            data1.add("晚了" + i);
            data1.add(i);
            data1.add("100%");
            list.add(data1);
        }
        return list;
    }

    // 创建动态日期表头
    private static List<List<String>> createAsinHead() {
        List<List<String>> list = ListUtils.newArrayList();

        List<String> jobNo = ListUtils.newArrayList();
        jobNo.add("ASIN 时效≥4天原因分析");
        jobNo.add("国家");
        list.add(jobNo);

        List<String> name = ListUtils.newArrayList();
        name.add("ASIN 时效≥4天原因分析");
        name.add("流程节点");
        list.add(name);

        List<String> reason = ListUtils.newArrayList();
        reason.add("ASIN 时效≥4天原因分析");
        reason.add("原因");
        list.add(reason);

        List<String> count = ListUtils.newArrayList();
        count.add("ASIN 时效≥4天原因分析");
        count.add("计数");
        list.add(count);

        List<String> ratio = ListUtils.newArrayList();
        ratio.add("ASIN 时效≥4天原因分析");
        ratio.add("比例");
        list.add(ratio);

        return list;

    }


    // 模拟填充数据
    private static List<List<Object>> createDataByList() {
        List<List<Object>> list = ListUtils.newArrayList();

        for (int i = 0; i < 2; i++) {

            List<Object> data = ListUtils.newArrayList();
            data.add("RY044" + i);
            data.add("张三" + i);
            data.add("技术部");

            data.add("出勤工时");
            List<Integer> dayKeyList = IntStream.rangeClosed(1, 28)
                    .boxed()
                    .collect(Collectors.toList());
            for (Integer integer : dayKeyList) {
                data.add("");
            }
            list.add(data);

            List<Object> data1 = ListUtils.newArrayList();
            data1.add("RY044" + i);
            data1.add("张三" + i);
            data1.add("技术部");
            data1.add("加班工时");
            for (Integer integer : dayKeyList) {
                int hour = new Random().nextInt(5);
                if (0 == hour) {
                    data1.add("");
                } else {
                    data1.add(new Random().nextInt(5));
                }
            }
            list.add(data1);

            List<Object> data2 = ListUtils.newArrayList();
            data2.add("RY044" + i);
            data2.add("张三" + i);
            data2.add("技术部");

            data2.add("考勤补贴");
            for (Integer integer : dayKeyList) {
                data2.add("");
            }
            list.add(data2);

            List<Object> data3 = ListUtils.newArrayList();
            data3.add("RY044" + i);
            data3.add("张三" + i);
            data3.add("技术部");

            data3.add("请假旷工");
            for (Integer integer : dayKeyList) {
                data3.add("");
            }
            list.add(data3);
        }
        return list;
    }

    // 创建动态日期表头
    private static List<List<String>> createDynamicHead() {
        List<List<String>> list = ListUtils.newArrayList();

        List<String> jobNo = ListUtils.newArrayList();
        jobNo.add("202201月考勤报表");
        jobNo.add("工号");
        list.add(jobNo);

        List<String> name = ListUtils.newArrayList();
        name.add("202201月考勤报表");
        name.add("姓名");
        list.add(name);

        List<String> dept = ListUtils.newArrayList();
        dept.add("202201月考勤报表");
        dept.add("部门");
        list.add(dept);

        List<String> attendanceItem = ListUtils.newArrayList();
        attendanceItem.add("202201月考勤报表");
        attendanceItem.add("考勤项");
        list.add(attendanceItem);

        /**2022-02-01 到 2022-02-28*/
        List<Integer> dayKeyList = IntStream.rangeClosed(1, 28)
                .boxed()
                .collect(Collectors.toList());

        for (Integer day : dayKeyList) {
            List<String> oneDay = ListUtils.newArrayList();
            oneDay.add("202201月考勤报表");
            oneDay.add(day + "");
            list.add(oneDay);
        }

        List<String> total = ListUtils.newArrayList();
        total.add("合计");
        list.add(total);
        return list;
    }

    /**
     * 构建默认的样式策略，带表头灰底加粗、边框
     */
    public static WriteHandler defaultStyleStrategy() {
        // 表头样式
        WriteCellStyle headStyle = new WriteCellStyle();
        headStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        WriteFont headFont = new WriteFont();
        headFont.setBold(true);
        headFont.setFontHeightInPoints((short) 12);
        headStyle.setWriteFont(headFont);
        setBorderStyle(headStyle);

        // 内容样式
        WriteCellStyle contentStyle = new WriteCellStyle();
        WriteFont contentFont = new WriteFont();
        contentFont.setFontHeightInPoints((short) 11);
        contentStyle.setWriteFont(contentFont);
        setBorderStyle(contentStyle);

        return new HorizontalCellStyleStrategy(headStyle, contentStyle);
    }

    /**
     * 设置四周边框样式
     */
    private static void setBorderStyle(WriteCellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }
}
