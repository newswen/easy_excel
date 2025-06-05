package com.yw.easy_excel_test.entity;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum ExcelFileTypeEnum {

	XLSX("xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
	XLS("xls", "application/vnd.ms-excel"),
	CSV("csv", "text/csv");

	private final String extension;  // 文件扩展名
	private final String contentType;  // 文件的 Content-Type 响应头格式

	// 构造方法
	ExcelFileTypeEnum(String extension, String description, String contentType) {
		this.extension = extension;
		this.contentType = contentType;
	}
}
