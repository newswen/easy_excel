package com.yw.easy_excel_test.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.yw.easy_excel_test.entity.ReasonRowVO;
import com.yw.easy_excel_test.service.ICspExcelExportService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * @Author: yw
 * @Date: 2025/7/11 09:37
 * @Description:
 **/
@Service
@Slf4j
public class AsinExportServiceImpl implements ICspExcelExportService {

    @Override
    public void excelExportTest() throws IOException {

//        List<List<String>> dataList = getDateList();
        // 这个自己更改
        String downloadPath = "D:\\idea\\easy_excel_test\\gzb.xlsx";

        // 这个自己更改
        String downloadBusinessFileName = "test2.xlsx";


//        EasyExcel.write(downloadPath, ReasonRowVO.class)
//                .sheet("sheet1")
//                .registerWriteHandler(registerWriteHandler())
//                .doWrite(dataList);

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
        //合并单元格区域1 US
        CellRangeAddress mergedRegionRemark = new CellRangeAddress(1, 14, 0, 0);
        sheet.addMergedRegion(mergedRegionRemark);
        Cell remarkCell = sheet.getRow(1).getCell(0);
        remarkCell.setCellStyle(cellStyle);
        // 下单-拣货（100%正常）
        CellRangeAddress mergedRegionRemark2 = new CellRangeAddress(1, 3, 1, 3);
        sheet.addMergedRegion(mergedRegionRemark2);
        Cell remarkCell2 = sheet.getRow(1).getCell(0);
        remarkCell2.setCellStyle(cellStyle);
        // 合并“拣货-扫描（20%扫描慢）” A5:A8
        CellRangeAddress mergedRegionPickScan = new CellRangeAddress(4, 7, 1, 1);
        sheet.addMergedRegion(mergedRegionPickScan);
        Cell remarkCellPickScan = sheet.getRow(4).getCell(0);
        remarkCellPickScan.setCellStyle(cellStyle);
        // 合并“扫描-送达（12%派送慢）” A10:A14
        CellRangeAddress mergedRegionScanDelivery = new CellRangeAddress(8, 13, 1, 1);
        sheet.addMergedRegion(mergedRegionScanDelivery);
        Cell remarkCellScanDelivery = sheet.getRow(8).getCell(0);
        remarkCellScanDelivery.setCellStyle(cellStyle);


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
}
