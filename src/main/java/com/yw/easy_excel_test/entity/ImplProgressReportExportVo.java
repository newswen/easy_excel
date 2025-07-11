package com.yw.easy_excel_test.entity;

import java.math.BigDecimal;
 
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentFontStyle;
import com.alibaba.excel.annotation.write.style.ContentStyle;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.enums.poi.BorderStyleEnum;
import com.alibaba.excel.enums.poi.FillPatternTypeEnum;
import com.alibaba.excel.enums.poi.HorizontalAlignmentEnum;
 
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;
 
/**
 * 导出实施进度报表
 * 
 * @author 
 * @date 2024/9/14 星期六
 * @since JDK 17
 */
@Data
@EqualsAndHashCode(callSuper = false)
@ToString(callSuper = true)
public class ImplProgressReportExportVo {
	/** 序号 */
	// 这一列 每隔2行 合并单元格
	// @ContentLoopMerge(eachRow = 2)
	@ExcelProperty(value = "序号", index = 0)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(10)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 1)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 1, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private String number;
 
	@ExcelProperty(value = "项目", index = 1)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(20)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 1)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 1, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private String fundName;
 
	/** 单位 */
	@ExcelProperty(value = "单位", index = 2)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(10)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 1)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 1, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private String unit;
 
	/** 初步设计指标/投资（设计单位填报） */
	@ExcelProperty(value = "初步设计指标", index = 3)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 42)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 42, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal totalPlan;
 
	/** 实施合计（当月完成指标） */
	@ExcelProperty(value = { "实施合计", "当月完成指标" }, index = 4)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 1)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 1, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal implCurMonthComplete;
 
	/** 实施合计（累计完成指标） */
	@ExcelProperty(value = { "实施合计", "累计完成指标" }, index = 5)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 1)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 1, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal implTotalComplete;
 
	/** 环北公司（当月完成指标） */
	@ExcelProperty(value = { "环北公司", "当月完成指标" }, index = 6)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 47)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 47, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal hbCurMonthComplete;
 
	/** 环北公司（累计完成指标） */
	@ExcelProperty(value = { "环北公司", "累计完成指标" }, index = 7)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 47)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 47, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal hbTotalComplete;
 
	/** 南宁市（当月完成指标） */
	@ExcelProperty(value = { "南宁市", "当月完成指标" }, index = 8)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal nanNingCurMonthComplete;
 
	/** 南宁市（累计完成指标） */
	@ExcelProperty(value = { "南宁市", "累计完成指标" }, index = 9)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal nanNingTotalComplete;
 
	/** 北海市（当月完成指标） */
	@ExcelProperty(value = { "北海市", "当月完成指标" }, index = 10)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal beiHaiCurMonthComplete;
 
	/** 北海市（累计完成指标） */
	@ExcelProperty(value = { "北海市", "累计完成指标" }, index = 11)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal beiHaiTotalComplete;
 
	/** 防城港市（当月完成指标） */
	@ExcelProperty(value = { "防城港市", "当月完成指标" }, index = 12)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal fcgCurMonthComplete;
 
	/** 防城港市（累计完成指标） */
	@ExcelProperty(value = { "防城港市", "累计完成指标" }, index = 13)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal fcgTotalComplete;
 
	/** 钦州市（当月完成指标） */
	@ExcelProperty(value = { "钦州市", "当月完成指标" }, index = 14)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal qinZhouCurMonthComplete;
 
	/** 钦州市（累计完成指标） */
	@ExcelProperty(value = { "钦州市", "累计完成指标" }, index = 15)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal qinZhouTotalComplete;
 
	/** 玉林市（当月完成指标） */
	@ExcelProperty(value = { "玉林市", "当月完成指标" }, index = 16)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal yuLinCurMonthComplete;
 
	/** 玉林市（累计完成指标） */
	@ExcelProperty(value = { "玉林市", "累计完成指标" }, index = 17)
	@HeadFontStyle(fontHeightInPoints = 10) // 字体大小
	@ColumnWidth(12)
	// HorizontalAlignmentEnum.CENTER 居中
	// FillPatternTypeEnum.SOLID_FOREGROUND : 充满
	// fillForegroundColor 见 https://www.cnblogs.com/hezemin/p/17272591.html
	@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, fillForegroundColor = 31)
	@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND,
			fillForegroundColor = 31, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
	@ContentFontStyle(fontHeightInPoints = 10) // 字体大小
	private BigDecimal yuLinTotalComplete;
}