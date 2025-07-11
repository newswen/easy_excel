package com.yw.easy_excel_test.utils;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;

public class CenterCellStyleHandler implements CellWriteHandler {

    private CellStyle centeredStyle;

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        Workbook workbook = writeSheetHolder.getSheet().getWorkbook();

        if (centeredStyle == null) {
            centeredStyle = workbook.createCellStyle();
            centeredStyle.setAlignment(HorizontalAlignment.CENTER);
            centeredStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // 保持边框样式可选（如果你用了边框）
            centeredStyle.setBorderTop(BorderStyle.THIN);
            centeredStyle.setBorderBottom(BorderStyle.THIN);
            centeredStyle.setBorderLeft(BorderStyle.THIN);
            centeredStyle.setBorderRight(BorderStyle.THIN);

            // 你可以加字体设置等
        }

        int colIndex = cell.getColumnIndex();

        // 判断列索引，0 和 1 列做居中
        if (colIndex == 0) {
            centeredStyle.setAlignment(HorizontalAlignment.CENTER);
            centeredStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cell.setCellStyle(centeredStyle);
        } else {
            //设置靠左
            centeredStyle.setAlignment(HorizontalAlignment.LEFT);
            centeredStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cell.setCellStyle(workbook.createCellStyle());
        }
    }
}
