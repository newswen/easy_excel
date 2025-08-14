package com.yw.easy_excel_test.utils;

import java.awt.*;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
 
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.AbstractCellStyleStrategy;
 
/**
 * 实施进度报表 主标题行样式
 * 
 * @author 
 * @date 2024/9/19 星期四
 * @since JDK 17
 */
public class ImplProgressReportTitleRowStyleStrategy extends AbstractCellStyleStrategy {
 
	@Override
	protected void setHeadCellStyle(CellWriteHandlerContext context) {
		// 获取和创建CellStyle
		WriteCellData<?> cellData = context.getFirstCellData();
		CellStyle originCellStyle = cellData.getOriginCellStyle();
		Cell cell = context.getCell();
 
		if (originCellStyle == null) {
			originCellStyle = context.getWriteWorkbookHolder().getWorkbook().createCellStyle();
		}
		// 设置背景颜色
		((XSSFCellStyle) originCellStyle)
				.setFillForegroundColor(new XSSFColor(Color.WHITE, new DefaultIndexedColorMap()));
		// 背景色充满
		originCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		originCellStyle.setWrapText(true);
		// 重点！！！
		// 由于在FillStyleCellWriteHandler，会把OriginCellStyle和WriteCellStyle合并，会已WriteCellStyle样式为主，所有必须把WriteCellStyle设置的背景色清空
		// 具体合并规则请看WriteWorkbookHolder.createCellStyle方法
		WriteCellStyle writeCellStyle = cellData.getWriteCellStyle();
		writeCellStyle.setFillForegroundColor(null);
		// 重点！！！ 必须设置OriginCellStyle
		cellData.setOriginCellStyle(originCellStyle);
 
		// 字体
		WriteFont headWriteFont = new WriteFont();
		if (cell.getRowIndex() == 0) {
			headWriteFont.setFontName("方正小标宋_GBK");
			// 字体颜色
			headWriteFont.setColor(IndexedColors.BLACK.getIndex());
			// 字体大小
			headWriteFont.setFontHeightInPoints((short) 20);
		}
		if (0 == context.getRowIndex()) {
			writeCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
		}
//		else {
//			writeCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
//		}
		cellData.getWriteCellStyle().setWriteFont(headWriteFont);
	}
 
	// 设置填充数据样式
	@Override
	protected void setContentCellStyle(CellWriteHandlerContext context) {
//		WriteFont contentWriteFont = new WriteFont();
//		contentWriteFont.setFontName("方正小标宋_GBK");
//		contentWriteFont.setFontHeightInPoints((short) 20);
//
//		WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
//		// 设置数据填充后的实线边框
//		contentWriteCellStyle.setWriteFont(contentWriteFont);
//		// 前景色充满
//		contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
//		// 前景色为白色，见：https://www.cnblogs.com/hezemin/p/17272591.html
//		contentWriteCellStyle.setFillForegroundColor((short) 9);
//		contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
//		contentWriteCellStyle.setBorderTop(BorderStyle.THIN);
//		contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
//		contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
//		DataFormatData dataFormatData = new DataFormatData();
//		dataFormatData.setIndex((short) 49);
//		contentWriteCellStyle.setDataFormatData(dataFormatData);
//		contentWriteCe llStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
//		WriteCellData<?> cellData = context.getFirstCellData();
//		WriteCellStyle.merge(contentWriteCellStyle, cellData.getOrCreateStyle());
	}


}