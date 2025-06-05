//package com.yw.easy_excel_test.fa_eu;
//
//import com.alibaba.excel.write.handler.CellWriteHandler;
//import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
//import com.alibaba.excel.write.metadata.RowData;
//
//public class HighlightCellWriteHandler implements CellWriteHandler {
//
//    @Override
//    public void afterCellDispose(CellWriteHandlerContext context) {
//        WriteRowData rowData = context.getFirstCellData().getRowData();
//        if (rowData instanceof RowData) {
//            Object vo = ((RowData<?>) rowData).getOriginal();
//            if (vo instanceof StockMovementModelVO) {
//                StockMovementModelVO data = (StockMovementModelVO) vo;
//                if (data.isHighlight()) {
//                    // 设置背景颜色为淡黄色
//                    CellStyle cellStyle = context.getWriteWorkbookHolder()
//                        .getCachedWorkbook().createCellStyle();
//                    cellStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
//                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//
//                    context.getCell().setCellStyle(cellStyle);
//                }
//            }
//        }
//    }
//}
