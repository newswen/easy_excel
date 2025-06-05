package com.yw.easy_excel_test.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.net.URL;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class CompetitorProductModelVO implements Serializable {

	private static final long serialVersionUID = 2826374031094971392L;

	@ExcelProperty(value = "竞品链接", index = 0)
	@ColumnWidth(30)
	private String productLink;

	@ExcelProperty(value = "竞品主图", index = 1)
	@ColumnWidth(50)
	private URL productImageUrl;

	@ExcelProperty(value = "竞品标题", index = 2)
	@ColumnWidth(40)
	private String productTitle;

	@ExcelProperty(value = "竞品五点", index = 3)
	@ColumnWidth(50)
	private String productFivePoints;

	@ExcelProperty(value = "竞品review数", index = 4)
	@ColumnWidth(15)
	private Integer reviewCount;


}
