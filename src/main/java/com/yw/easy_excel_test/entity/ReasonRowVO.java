package com.yw.easy_excel_test.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

@Data
public class ReasonRowVO {
    @ExcelProperty("原因")
    @ColumnWidth(40)
    private String reason;

    @ExcelProperty("计数")
    @ColumnWidth(10)
    private Integer count;

    @ExcelProperty("比例")
    @ColumnWidth(10)
    private String ratio;
} 