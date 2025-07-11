package com.yw.easy_excel_test.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.yw.easy_excel_test.service.ICspExcelExportService;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.StringUtil;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * @Author: yw
 * @Date: 2025/7/11 09:13
 * @Description:
 **/
@Service
@Slf4j
public class CspExcelExportServiceImpl implements ICspExcelExportService {

    @Override
    public void excelExportTest() throws IOException {
        List<List<String>> headList = getHeadList();
        List<List<String>> dataList = getDateList();
        // 这个自己更改
        String downloadPath = "D:\\idea\\easy_excel_test\\gzb.xlsx";

        // 这个自己更改
        String downloadBusinessFileName = "test2.xlsx";


        EasyExcel.write(downloadPath)
                .head(roateHeadFields(headList))
                .sheet("sheet1")
                .registerWriteHandler(registerWriteHandler())
                .doWrite(dataList);


        // 加载已导出的文件c
        File exportedFile = new File(downloadPath);
        Workbook workbook = WorkbookFactory.create(exportedFile);
        Sheet sheet = workbook.getSheet("sheet1");
        // 移除已有的合并区域
        // sheet.removeMergedRegion(0);
        // 创建样式对象
        CellStyle cellStyle = workbook.createCellStyle();
        // 设置边框
        // cellStyle.setBorderTop(BorderStyle.THIN);
        // cellStyle.setBorderBottom(BorderStyle.THIN);
        // cellStyle.setBorderLeft(BorderStyle.THIN);
        // cellStyle.setBorderRight(BorderStyle.THIN);
        // 设置水平对齐方式
        cellStyle.setAlignment(HorizontalAlignment.CENTER); // 居中对齐
        // 设置合并单元格区域1 月度客诉率
        CellRangeAddress mergedRegionRemark = new CellRangeAddress(3, 3, 1, 3);
        sheet.addMergedRegion(mergedRegionRemark);
        Cell remarkCell = sheet.getRow(3).getCell(0);
        remarkCell.setCellStyle(cellStyle);
        // 设置合并单元格区域2   月度工单总量
        CellRangeAddress mergedRegionRemark2 = new CellRangeAddress(4, 4, 1, 3);
        sheet.addMergedRegion(mergedRegionRemark2);
        Cell remarkCell2 = sheet.getRow(4).getCell(0);
        remarkCell2.setCellStyle(cellStyle);

        // 设置合并单元格区域3   客诉类型序号
        CellRangeAddress mergedRegionRemark3 = new CellRangeAddress(5, 7, 0, 0);
        sheet.addMergedRegion(mergedRegionRemark3);
        Cell remarkCell3 = sheet.getRow(5).getCell(0);
        remarkCell3.setCellStyle(cellStyle);

        // 设置合并单元格区域4   客诉类型
        CellRangeAddress mergedRegionRemark4 = new CellRangeAddress(5, 7, 1, 1);
        sheet.addMergedRegion(mergedRegionRemark4);
        Cell remarkCell4 = sheet.getRow(5).getCell(0);
        remarkCell4.setCellStyle(cellStyle);

        // 设置合并单元格区域5 客户原因
        CellRangeAddress mergedRegionRemark5 = new CellRangeAddress(5, 5, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark5);
        Cell remarkCell5 = sheet.getRow(5).getCell(0);
        remarkCell5.setCellStyle(cellStyle);

        // 设置合并单元格区域6 绝配原因
        CellRangeAddress mergedRegionRemark6 = new CellRangeAddress(6, 7, 2, 2);
        sheet.addMergedRegion(mergedRegionRemark6);
        Cell remarkCell6 = sheet.getRow(6).getCell(0);
        remarkCell6.setCellStyle(cellStyle);

        // 设置合并单元格区域7   工单类型
        CellRangeAddress mergedRegionRemark7 = new CellRangeAddress(8, 9, 1, 1);
        sheet.addMergedRegion(mergedRegionRemark7);
        Cell remarkCell7 = sheet.getRow(8).getCell(0);
        remarkCell7.setCellStyle(cellStyle);

        // 设置合并单元格区域7   工单类型-序号
        CellRangeAddress mergedRegionRemark7Number = new CellRangeAddress(8, 9, 0, 0);
        sheet.addMergedRegion(mergedRegionRemark7Number);
        Cell remarkCell7Number = sheet.getRow(8).getCell(0);
        remarkCell7Number.setCellStyle(cellStyle);

        // 设置合并单元格区域8   工单类型-普通工单
        CellRangeAddress mergedRegionRemark8 = new CellRangeAddress(8, 8, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark8);
        Cell remarkCell8 = sheet.getRow(8).getCell(0);
        remarkCell8.setCellStyle(cellStyle);

        // 设置合并单元格区域9   工单类型-理赔工单
        CellRangeAddress mergedRegionRemark9 = new CellRangeAddress(9, 9, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark9);
        Cell remarkCell9 = sheet.getRow(9).getCell(0);
        remarkCell9.setCellStyle(cellStyle);

        // 设置合并单元格区域10   盘点差异
        CellRangeAddress mergedRegionRemark10 = new CellRangeAddress(10, 12, 1, 1);
        sheet.addMergedRegion(mergedRegionRemark10);
        Cell remarkCell10 = sheet.getRow(10).getCell(0);
        remarkCell10.setCellStyle(cellStyle);

        // 设置合并单元格区域10   盘点差异-序号
        CellRangeAddress mergedRegionRemark10Number = new CellRangeAddress(10, 12, 0, 0);
        sheet.addMergedRegion(mergedRegionRemark10Number);
        Cell remarkCell10Number = sheet.getRow(10).getCell(0);
        remarkCell10Number.setCellStyle(cellStyle);


        // 设置合并单元格区域11   盘点差异-差异SKU数量
        CellRangeAddress mergedRegionRemark11 = new CellRangeAddress(10, 10, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark11);
        Cell remarkCell11 = sheet.getRow(10).getCell(0);
        remarkCell11.setCellStyle(cellStyle);

        // 设置合并单元格区域12   盘点差异-库存准确率
        CellRangeAddress mergedRegionRemark12 = new CellRangeAddress(11, 11, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark12);
        Cell remarkCell12 = sheet.getRow(11).getCell(0);
        remarkCell12.setCellStyle(cellStyle);

        // 设置合并单元格区域13   盘点差异-库存准确率
        CellRangeAddress mergedRegionRemark13 = new CellRangeAddress(12, 12, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark13);
        Cell remarkCell13 = sheet.getRow(12).getCell(0);
        remarkCell13.setCellStyle(cellStyle);

        // 设置合并单元格区域14   理赔承担
        CellRangeAddress mergedRegionRemark14 = new CellRangeAddress(13, 17, 1, 1);
        sheet.addMergedRegion(mergedRegionRemark14);
        Cell remarkCell14 = sheet.getRow(13).getCell(0);
        remarkCell14.setCellStyle(cellStyle);

        // 设置合并单元格区域14   理赔承担-序号
        CellRangeAddress mergedRegionRemark14Number = new CellRangeAddress(13, 17, 0, 0);
        sheet.addMergedRegion(mergedRegionRemark14Number);
        Cell remarkCell14Number = sheet.getRow(13).getCell(0);
        remarkCell14Number.setCellStyle(cellStyle);

        // 设置合并单元格区域15   理赔承担-客户承担
        CellRangeAddress mergedRegionRemark15 = new CellRangeAddress(13, 13, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark15);
        Cell remarkCell15 = sheet.getRow(13).getCell(0);
        remarkCell15.setCellStyle(cellStyle);

        // 设置合并单元格区域16   理赔承担-下游承担
        CellRangeAddress mergedRegionRemark16 = new CellRangeAddress(14, 15, 2, 2);
        sheet.addMergedRegion(mergedRegionRemark16);
        Cell remarkCell16 = sheet.getRow(14).getCell(0);
        remarkCell16.setCellStyle(cellStyle);

        // 设置合并单元格区域17   理赔承担-绝配承担
        CellRangeAddress mergedRegionRemark17 = new CellRangeAddress(16, 16, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark17);
        Cell remarkCell17 = sheet.getRow(16).getCell(0);
        remarkCell17.setCellStyle(cellStyle);

        // 设置合并单元格区域18   理赔承担-月度理赔总额
        CellRangeAddress mergedRegionRemark18 = new CellRangeAddress(17, 17, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark18);
        Cell remarkCell18 = sheet.getRow(17).getCell(0);
        remarkCell18.setCellStyle(cellStyle);

        // 设置合并单元格区域19   原因归类
        CellRangeAddress mergedRegionRemark19 = new CellRangeAddress(18, 19, 1, 1);
        sheet.addMergedRegion(mergedRegionRemark19);
        Cell remarkCell19 = sheet.getRow(18).getCell(0);
        remarkCell19.setCellStyle(cellStyle);

        // 设置合并单元格区域19   原因归类-序号
        CellRangeAddress mergedRegionRemark19Number = new CellRangeAddress(18, 19, 0, 0);
        sheet.addMergedRegion(mergedRegionRemark19Number);
        Cell remarkCell19Number = sheet.getRow(18).getCell(0);
        remarkCell19Number.setCellStyle(cellStyle);

        // 设置合并单元格区域20   原因归类-月度工单量产生比例最高
        CellRangeAddress mergedRegionRemark20 = new CellRangeAddress(18, 18, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark20);
        Cell remarkCell20 = sheet.getRow(18).getCell(0);
        remarkCell20.setCellStyle(cellStyle);

        // 设置合并单元格区域21   原因归类-单笔理赔金额最高
        CellRangeAddress mergedRegionRemark21 = new CellRangeAddress(19, 19, 2, 3);
        sheet.addMergedRegion(mergedRegionRemark21);
        Cell remarkCell21 = sheet.getRow(19).getCell(0);
        remarkCell21.setCellStyle(cellStyle);

        // 自定义列宽
        int maxLength = 50;
        if (!CollectionUtils.isEmpty(dataList)) {
            for (int i = 0; i < dataList.get(0).size(); i++) {
                sheet.setColumnWidth(i, (maxLength) * 125);
            }
        }


        // 保存文件
        try (FileOutputStream fos = new FileOutputStream(downloadBusinessFileName)) {
            workbook.write(fos);
            workbook.close();
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }


    /**
     * 表格样式设置
     */
    private HorizontalCellStyleStrategy registerWriteHandler() {
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 头背景设置
        headWriteCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 10);
        headWriteCellStyle.setWriteFont(headWriteFont);
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        WriteFont font = new WriteFont();
//        font.setFontName("微软雅黑");
        font.setColor(IndexedColors.WHITE.getIndex()); // 设置文本颜色为白色
        headWriteCellStyle.setWriteFont(font);

        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
//        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short) 10);
//        contentWriteFont.setFontName("微软雅黑");
        contentWriteCellStyle.setWriteFont(contentWriteFont);
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        contentWriteCellStyle.setWrapped(true); // 设置自动换行
        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置边框
//        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
//        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);

        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        return new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
    }

    private List<List<String>> roateHeadFields(List<List<String>> headFields) {
        List<List<String>> result = new ArrayList<>();
        for (List<String> row : headFields) {
            for (int j = 0; j < row.size(); j++) {
                if (result.size() > j) {
                    // 往对应第j个List<String> 添加添加值
                    result.get(j).add(row.get(j));
                } else {
                    // 分割成单个List<String>
                    result.add(new ArrayList<>(Collections.singletonList(row.get(j))));
                }
            }
        }
        return result;
    }

    private List<List<String>> getHeadList() {
        List<List<String>> headList = new ArrayList<>();
        List<String> titleOne = Arrays.asList("XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据", "XX项目管理基础数据");
        headList.add(makeUpTitle(titleOne));
        List<String> titleTwo = Arrays.asList("序号", "分析维度", "分析维度", "分析维度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度", "2023年时间跨度");
        headList.add(makeUpTitle(titleTwo));
        List<String> titleThree = Arrays.asList("序号", "一级", "二级", "三级", "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月");
        headList.add(makeUpTitle(titleThree));
        return headList;
    }

    private List<String> makeUpTitle(List<String> titleList) {
        List<String> makeUpList = new ArrayList<>();
        String temp = "";
        for (String e : org.apache.commons.collections4.ListUtils.emptyIfNull(titleList)) {
            if (StringUtil.isNotBlank(e)) {
                temp = e;
            }
            makeUpList.add(temp);
        }
        return makeUpList;
    }


    private List<List<String>> getDateList() {
        List<List<String>> dataList = new ArrayList<>();
        for (int i = 0; i < 17; i++) {
            dataList.add(buildData(i));
        }
        return dataList;
    }

    public List<String> buildData(int i) {
        List<String> list = new ArrayList<>();
        switch (i) {
            case 0:
                list.add(String.valueOf(i + 1));
                list.add("月度客诉率-%");
                list.add("月度客诉率-%");
                list.add("月度客诉率-%");
                list.add("3.01");
                return list;
            case 1:
                list.add(String.valueOf(i + 1));
                list.add("月度工单总量");
                list.add("月度工单总量");
                list.add("月度工单总量");
                list.add("10");
                return list;
            case 2:
                list.add(String.valueOf(i + 1));
                list.add("客诉类型");
                list.add("客户原因");
                list.add("客户原因");
                list.add("2");
                return list;
            case 3:
                list.add(String.valueOf(i));
                list.add("客诉类型");
                list.add("绝配原因");
                list.add("仓");
                list.add("5");
                return list;
            case 4:
                list.add(String.valueOf(i));
                list.add("客诉类型");
                list.add("绝配原因");
                list.add("配");
                list.add("1");
                return list;
            case 5:
                list.add(String.valueOf(i - 1));
                list.add("工单类型");
                list.add("普通工单");
                list.add("普通工单");
                list.add("8");
                return list;
            case 6:
                list.add(String.valueOf(i));
                list.add("工单类型");
                list.add("理赔工单");
                list.add("理赔工单");
                list.add("2");
                return list;
            case 7:
                list.add(String.valueOf(i - 2));
                list.add("盘点差异");
                list.add("差异SKU数量");
                list.add("差异SKU数量");
                list.add("13");
                return list;
            case 8:
                list.add(String.valueOf(i));
                list.add("盘点差异");
                list.add("库存准确率");
                list.add("库存准确率");
                list.add("");
                return list;
            case 9:
                list.add(String.valueOf(i));
                list.add("盘点差异");
                list.add("月度理赔金额");
                list.add("月度理赔金额");
                list.add("1500.00");
                return list;
            case 10:
                list.add(String.valueOf(i - 4));
                list.add("理赔承担");
                list.add("客户承担");
                list.add("客户承担");
                list.add("1500.00");
                return list;
            case 11:
                list.add(String.valueOf(i));
                list.add("理赔承担");
                list.add("下游承担");
                list.add("仓");
                list.add("1500.00");
                return list;
            case 12:
                list.add(String.valueOf(i));
                list.add("理赔承担");
                list.add("下游承担");
                list.add("配");
                list.add("1500.00");
                return list;
            case 13:
                list.add(String.valueOf(i));
                list.add("理赔承担");
                list.add("绝配承担");
                list.add("绝配承担");
                list.add("1500.00");
                return list;
            case 14:
                list.add(String.valueOf(i));
                list.add("理赔承担");
                list.add("月度理赔总额");
                list.add("月度理赔总额");
                list.add("1500.00");
                return list;
            case 15:
                list.add(String.valueOf(i - 8));
                list.add("原因归类");
                list.add("月度工单量产生比例最高");
                list.add("月度工单量产生比例最高");
                list.add("少货（75%）");
                return list;
            case 16:
                list.add(String.valueOf(i));
                list.add("原因归类");
                list.add("单笔理赔金额最高");
                list.add("单笔理赔金额最高");
                list.add("库存差异（30%）");
                return list;
            default:
                return null;
        }
    }



}
