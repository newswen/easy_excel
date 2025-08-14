package com.yw.easy_excel_test.controller;

import cn.hutool.core.io.IoUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.merge.OnceAbsoluteMergeStrategy;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.yw.easy_excel_test.entity.ImplProgressReportExportVo;
import com.yw.easy_excel_test.entity.ImplProgressReportRequest;
import com.yw.easy_excel_test.utils.ImplProgressReportTitleRowHeightStyleStrategy;
import com.yw.easy_excel_test.utils.ImplProgressReportTitleRowStyleStrategy;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.*;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * xxx统计 API接口
 * 
 * @author 
 * @date 2024/8/26
 * @since JDK 17
 */
@Slf4j
@RestController
@RequestMapping(value = "/api")
public class MigrateBidFillStatController {
 
	@GetMapping("/export/implProgressReport")
	public void implProgressReportExport(HttpServletResponse response) {
		/* 参考：https://blog.csdn.net/qq_43932985/article/details/141644977 */

		ServletOutputStream os = null;
		ExcelWriter excelWriter = null;
 
		try {
			this.setExcelResponseProp(response, "实施进度报表");
 
			/* 第一个报表 */
			ImplProgressReportExportVo vo1 = new ImplProgressReportExportVo();
			vo1.setNumber("一");
			vo1.setFundName("土地");
			vo1.setUnit("亩");
			vo1.setTotalPlan(new BigDecimal("31751.62"));
			vo1.setImplCurMonthComplete(new BigDecimal("391.46"));
			vo1.setImplTotalComplete(new BigDecimal("3185.2"));
			//vo1.setHbCurMonthComplete(BigDecimal.ZERO);
			//vo1.setHbTotalComplete(BigDecimal.ZERO);
			vo1.setNanNingCurMonthComplete(new BigDecimal("167.7"));
			vo1.setNanNingTotalComplete(new BigDecimal("620.31"));
			vo1.setBeiHaiCurMonthComplete(new BigDecimal("45.22"));
			vo1.setBeiHaiTotalComplete(new BigDecimal("437.43"));
			vo1.setFcgCurMonthComplete(new BigDecimal("7.53"));
			vo1.setFcgTotalComplete(new BigDecimal("101.98"));
			vo1.setQinZhouCurMonthComplete(new BigDecimal("171.01"));
			vo1.setQinZhouTotalComplete(new BigDecimal("1581.62"));
			//vo1.setYuLinCurMonthComplete(BigDecimal.ZERO);
			vo1.setYuLinTotalComplete(new BigDecimal("443.86"));
 
			ImplProgressReportExportVo vo2 = new ImplProgressReportExportVo();
			vo2.setNumber("1");
			vo2.setFundName("永久用地");
			vo2.setUnit("亩");
			vo2.setTotalPlan(new BigDecimal("2605.88"));
//			vo2.setImplCurMonthComplete(BigDecimal.ZERO);
//			vo2.setImplTotalComplete(BigDecimal.ZERO);
//			vo2.setHbCurMonthComplete(BigDecimal.ZERO);
//			vo2.setHbTotalComplete(BigDecimal.ZERO);
//			vo2.setNanNingCurMonthComplete(BigDecimal.ZERO);
//			vo2.setNanNingTotalComplete(BigDecimal.ZERO);
//			vo2.setBeiHaiCurMonthComplete(BigDecimal.ZERO);
//			vo2.setBeiHaiTotalComplete(BigDecimal.ZERO);
//			vo2.setFcgCurMonthComplete(BigDecimal.ZERO);
//			vo2.setFcgTotalComplete(BigDecimal.ZERO);
//			vo2.setQinZhouCurMonthComplete(BigDecimal.ZERO);
//			vo2.setQinZhouTotalComplete(BigDecimal.ZERO);
//			vo2.setYuLinCurMonthComplete(BigDecimal.ZERO);
//			vo2.setYuLinTotalComplete(BigDecimal.ZERO);
 
			ImplProgressReportExportVo vo3 = new ImplProgressReportExportVo();
			vo3.setNumber("2");
			vo3.setFundName("临时用地");
			vo3.setUnit("亩");
			vo3.setTotalPlan(new BigDecimal("29145.74"));
			vo3.setImplCurMonthComplete(new BigDecimal("391.46"));
			vo3.setImplTotalComplete(new BigDecimal("3185.2"));
			//vo3.setHbCurMonthComplete(BigDecimal.ZERO);
			//vo3.setHbTotalComplete(BigDecimal.ZERO);
			vo3.setNanNingCurMonthComplete(new BigDecimal("167.7"));
			vo3.setNanNingTotalComplete(new BigDecimal("620.31"));
			vo3.setBeiHaiCurMonthComplete(new BigDecimal("45.22"));
			vo3.setBeiHaiTotalComplete(new BigDecimal("437.43"));
			vo3.setFcgCurMonthComplete(new BigDecimal("7.53"));
			vo3.setFcgTotalComplete(new BigDecimal("101.98"));
			vo3.setQinZhouCurMonthComplete(new BigDecimal("171.01"));
			vo3.setQinZhouTotalComplete(new BigDecimal("1581.62"));
			//vo3.setYuLinCurMonthComplete(BigDecimal.ZERO);
			vo3.setYuLinTotalComplete(new BigDecimal("443.86"));
 
			os = response.getOutputStream();
 
//			EasyExcel.write(os).head(ImplProgressReportExportVo.class).excelType(ExcelTypeEnum.XLSX)
//					.registerWriteHandler(new ImplProgressReportCellWriteHandler(2, 2, new int[]{
//					// 需要合并部门列、部门描述列、工资范围列
//					0, 1, 2}, ((cur, pre) -> {
//				// 部门名称相同 && 工资范围相同才需要合并
//				String curDept = cur.getCell(0).getStringCellValue();
//				String preDept = pre.getCell(0).getStringCellValue();
//				String curSalaryRange = cur.getCell(2).getStringCellValue();
//				String preSalaryRange = pre.getCell(2).getStringCellValue();
//				return curDept.equals(preDept) && curSalaryRange.equals(preSalaryRange) ? true : false;
//			}))).sheet("实施进度报表").doWrite(Arrays.asList(vo1, vo2, vo3));
 
			List<ImplProgressReportExportVo> firstList = new ArrayList<>(Arrays.asList(vo1, vo2, vo3));
 
			/* 第二个报表 */
			ImplProgressReportExportVo vo4 = new ImplProgressReportExportVo();
			vo4.setNumber("第一部分");
			vo4.setFundName("农村部分补偿费");
			vo4.setUnit("万元");
			vo4.setTotalPlan(new BigDecimal("98257.23"));
			vo4.setImplCurMonthComplete(new BigDecimal("1561.39"));
			vo4.setImplTotalComplete(new BigDecimal("5073.29"));
			//vo4.setHbCurMonthComplete(BigDecimal.ZERO);
			//vo4.setHbTotalComplete(BigDecimal.ZERO);
			vo4.setNanNingCurMonthComplete(new BigDecimal("1187.82"));
			vo4.setNanNingTotalComplete(new BigDecimal("1825.92"));
			vo4.setBeiHaiCurMonthComplete(new BigDecimal("167.54"));
			vo4.setBeiHaiTotalComplete(new BigDecimal("780.02"));
			vo4.setFcgCurMonthComplete(new BigDecimal("3.38"));
			vo4.setFcgTotalComplete(new BigDecimal("96.55"));
			vo4.setQinZhouCurMonthComplete(new BigDecimal("202.65"));
			vo4.setQinZhouTotalComplete(new BigDecimal("1889.58"));
			//vo4.setYuLinCurMonthComplete(BigDecimal.ZERO);
			vo4.setYuLinTotalComplete(new BigDecimal("481.22"));
 
			ImplProgressReportExportVo vo5 = new ImplProgressReportExportVo();
			vo5.setNumber("一");
			vo5.setFundName("征地补偿费及青苗、林木补偿费");
			vo5.setUnit("万元");
			vo5.setTotalPlan(new BigDecimal("90672.63"));
			vo5.setImplCurMonthComplete(new BigDecimal("1464.47"));
			vo5.setImplTotalComplete(new BigDecimal("4854.12"));
			//	vo5.setHbCurMonthComplete(BigDecimal.ZERO);
			//	vo5.setHbTotalComplete(BigDecimal.ZERO);
			vo5.setNanNingCurMonthComplete(BigDecimal.ZERO);
			vo5.setNanNingTotalComplete(new BigDecimal("4854.12"));
			vo5.setBeiHaiCurMonthComplete(new BigDecimal("167.21"));
			vo5.setBeiHaiTotalComplete(new BigDecimal("684.13"));
			vo5.setFcgCurMonthComplete(new BigDecimal("3.38"));
			vo5.setFcgTotalComplete(new BigDecimal("96.2"));
			vo5.setQinZhouCurMonthComplete(new BigDecimal("111.55"));
			vo5.setQinZhouTotalComplete(new BigDecimal("1780.64"));
			//	vo5.setYuLinCurMonthComplete(BigDecimal.ZERO);
			vo5.setYuLinTotalComplete(new BigDecimal("472.72"));
 
			ImplProgressReportExportVo vo6 = new ImplProgressReportExportVo();
			vo6.setNumber("二");
			vo6.setFundName("移民搬迁及房屋设施补偿费");
			vo6.setUnit("万元");
			vo6.setTotalPlan(new BigDecimal("1873.8"));
			vo6.setImplCurMonthComplete(new BigDecimal("91.68"));
			vo6.setImplTotalComplete(new BigDecimal("157.15"));
			//vo6.setHbCurMonthComplete(BigDecimal.ZERO);
			//vo6.setHbTotalComplete(BigDecimal.ZERO);
			vo6.setNanNingCurMonthComplete(new BigDecimal("0.25"));
			vo6.setNanNingTotalComplete(new BigDecimal("0.25"));
			vo6.setBeiHaiCurMonthComplete(new BigDecimal("0.33"));
			vo6.setBeiHaiTotalComplete(new BigDecimal("52.96"));
			// vo6.setFcgCurMonthComplete(new BigDecimal("7.53"));
			// vo6.setFcgTotalComplete(new BigDecimal("101.98"));
			vo6.setQinZhouCurMonthComplete(new BigDecimal("91.10"));
			vo6.setQinZhouTotalComplete(new BigDecimal("95.44"));
			//vo6.setYuLinCurMonthComplete(BigDecimal.ZERO);
			vo6.setYuLinTotalComplete(new BigDecimal("8.50"));
 
			List<ImplProgressReportExportVo> secondList = new ArrayList<>(Arrays.asList(vo4, vo5, vo6));
 
			excelWriter = EasyExcel.write(os).build();
			// 把sheet设置为不需要头 不然会输出sheet的头 这样看起来第一个table 就有2个头了
			WriteSheet writeSheet = EasyExcel.writerSheet().needHead(Boolean.FALSE).sheetName("实施进度报表").build();
 
			// 主标题
			List<String> title = Arrays.asList("xxx报表");
            List<List<String>> titleHead = new ArrayList<>(Arrays.asList(title));
 
			// 合并主标题行：将第1行的第1-18列合并
			OnceAbsoluteMergeStrategy absoluteMergeStrategy = new OnceAbsoluteMergeStrategy(0, 0, 0, 17);
 
			// 主标题
			WriteTable titleTable = EasyExcel.writerTable(0)
					.head(titleHead)
					.needHead(Boolean.TRUE)
					// 自动合并头，头中相同的字段上下左右都会去尝试匹配
					.automaticMergeHead(Boolean.TRUE)
					.registerWriteHandler(absoluteMergeStrategy)
					// 行高
					.registerWriteHandler(new ImplProgressReportTitleRowHeightStyleStrategy())
					// 样式，参考：https://blog.csdn.net/weixin_44077141/article/details/139008521
					.registerWriteHandler(new ImplProgressReportTitleRowStyleStrategy())
					.build();
 
			// 这里必须指定需要头，table 会继承sheet的配置，sheet配置了不需要，table 默认也是不需要
			WriteTable writeTable1 = EasyExcel.writerTable(1)
					.head(ImplProgressReportExportVo.class)
					.needHead(Boolean.TRUE)
					.build();
 
			// 第二个对象 读取对象的excel实体类中的标题
			WriteTable writeTable2 = EasyExcel.writerTable(2)
					.head(ImplProgressReportExportVo.class)
					.needHead(Boolean.TRUE)
					// 和第一个报表间隔两行
					.relativeHeadRowIndex(2)
					.build();
 
			// 写入主标题
			excelWriter.write(new ArrayList<>(), writeSheet, titleTable);
			// 第一次写入会创建头
			excelWriter.write(firstList, writeSheet, writeTable1);
			// 第二次写如也会创建头，然后在第一次的后面写入数据
			excelWriter.write(secondList, writeSheet, writeTable2);
 
		} catch (Exception e) {
            excelWriter.finish();
        }
	}
 
	/**
	 * 设置响应结果
	 *
	 * @param response    响应结果对象
	 * @param rawFileName 文件名
	 * @throws UnsupportedEncodingException 不支持编码异常
	 */
	private void setExcelResponseProp(HttpServletResponse response, String rawFileName) throws UnsupportedEncodingException {
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		response.setCharacterEncoding("utf-8");
		String fileName = URLEncoder.encode(rawFileName, "UTF-8").replaceAll("\\+", "%20");
		response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
	}
}