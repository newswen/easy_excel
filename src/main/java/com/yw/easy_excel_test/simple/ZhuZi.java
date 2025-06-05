package com.yw.easy_excel_test.simple;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ZhuZi {

    @ExcelProperty("编号")
    private Long id;

    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("性别")
    @ExcelIgnore
    private String sex;

    @ExcelProperty("爱好")
    private String hobby;

    @ExcelProperty("生日")
    private Date birthday;
}
