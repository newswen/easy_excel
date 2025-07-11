package com.yw.easy_excel_test.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.*;
import com.alibaba.excel.enums.poi.BorderStyleEnum;
import com.alibaba.excel.enums.poi.FillPatternTypeEnum;
import com.alibaba.excel.enums.poi.HorizontalAlignmentEnum;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
// 头背景设置
@HeadStyle(fillPatternType = FillPatternTypeEnum.SOLID_FOREGROUND, horizontalAlignment = HorizontalAlignmentEnum.CENTER, borderLeft = BorderStyleEnum.THIN, borderTop = BorderStyleEnum.THIN, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
//标题高度
@HeadRowHeight(30)
//内容高度
@ContentRowHeight(20)
//内容居中,左、上、右、下的边框显示
@ContentStyle(horizontalAlignment = HorizontalAlignmentEnum.CENTER, borderLeft = BorderStyleEnum.THIN, borderTop = BorderStyleEnum.THIN, borderRight = BorderStyleEnum.THIN, borderBottom = BorderStyleEnum.THIN)
public class OrderDetailEntity {

    @ExcelProperty(value = "订单号")
    @ColumnWidth(25)
    private String orderCode;

    @ExcelProperty(value = "订单明细")
    @ColumnWidth(40)
    private String orderDetailCode;

    @ExcelProperty(value = "商品分类")
    @ColumnWidth(20)
    private String productCategory;

    @ExcelProperty(value = "商品编码")
    @ColumnWidth(20)
    private String productCode;

    @ExcelProperty(value = "商品名称")
    @ColumnWidth(20)
    private String productName;

    @ExcelProperty(value = "单价")
    @ColumnWidth(10)
    private BigDecimal price;

    @ExcelProperty(value = "数量")
    @ColumnWidth(10)
    private BigDecimal quantity;

    @ExcelProperty(value = "状态")
    @ColumnWidth(10)
    private String status;

    @ExcelProperty(value = "分类总数")
    @ColumnWidth(20)
    //@ExcelIgnore // 案例一、案例二放开该注解
    private BigDecimal categoryTotalQuantity;

    @ExcelProperty(value = "分类总金额")
    @ColumnWidth(20)
    //@ExcelIgnore // 案例一、案例二放开该注解
    private BigDecimal categoryTotalPrice;

    @ExcelProperty(value = "总数")
    @ColumnWidth(10)
    //@ExcelIgnore // 案例一放开该注解
    private BigDecimal totalQuantity;

    @ExcelProperty(value = "总金额")
    @ColumnWidth(10)
    //@ExcelIgnore // 案例一放开该注解
    private BigDecimal totalPrice;
}

