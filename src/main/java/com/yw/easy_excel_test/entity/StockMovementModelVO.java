package com.yw.easy_excel_test.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.time.LocalDate;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class StockMovementModelVO implements Serializable {

    private static final long serialVersionUID = 1L;

    @ExcelProperty(value = "index", index = 0)
    private Integer index;

    @ExcelProperty(value = "sku", index = 1)
    @ColumnWidth(20)
    private String sku;

    @ExcelProperty(value = "invoiceNumber", index = 2)
    @ColumnWidth(20)
    private String invoiceNumber;

    @ExcelProperty(value = "invoiceDate", index = 3)
    @ColumnWidth(20)
    private LocalDate invoiceDate;

    @ExcelProperty(value = "inboundDate", index = 4)
    @ColumnWidth(20)
    private LocalDate inboundDate;

    @ExcelProperty(value = "beginOfTheMonthStock", index = 5)
    private Integer beginOfTheMonthStock;

    @ExcelProperty(value = "endOfTheMonthStock", index = 6)
    private Integer endOfTheMonthStock;

    @ExcelProperty(value = "inbound", index = 7)
    private Integer inbound;

    @ExcelProperty(value = "sales", index = 8)
    private Integer sales;

    @ExcelProperty(value = "transfer", index = 9)
    private Integer transfer;

    @ExcelProperty(value = "lost", index = 10)
    private Integer lost;

    @ExcelProperty(value = "restock", index = 11)
    private Integer restock;

    @ExcelProperty(value = "inPriceEur", index = 12)
    private Double inPriceEur;

    @ExcelProperty(value = "inPriceUsd", index = 13)
    private Double inPriceUsd;

    @ExcelProperty(value = "currencyRate", index = 14)
    private Double currencyRate;

    @ExcelProperty(value = "pln", index = 15)
    private Double pln;

    @ExcelProperty(value = "type", index = 16)
    private String type;

    @ExcelProperty(value = "flag", index = 17)
    private String flag;

    @ExcelProperty(value = "level0", index = 18)
    private Integer level0;

    @ExcelProperty(value = "inPriceUsdX", index = 19)
    private Double inPriceUsdX;

    @ExcelProperty(value = "transferY", index = 20)
    private Integer transferY;

    @ExcelProperty(value = "name",  index = 21)
    private String name;
}
