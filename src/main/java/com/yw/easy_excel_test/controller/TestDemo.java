package com.yw.easy_excel_test.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.yw.easy_excel_test.entity.AmazonAsin;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class TestDemo {

    /**
     * 案例一：工资表
     */
    public void salaryList() {
        // 模板文件路径（建议用 .xlsx）
        String templateFilePath = "AsinAnalysisTemplate1.xlsx";
        // 输出文件路径
        String outFilePath = "gzb-filled.xlsx";

        // 创建 ExcelWriter 实例，绑定模板
        ExcelWriter writer = EasyExcel
                .write(outFilePath)
                .withTemplate(templateFilePath)
                .build();

        // ✅ 选中第三个 Sheet（索引从 0 开始，所以是 2）
        WriteSheet sheet = EasyExcel.writerSheet(2).build();

        // 模拟数据
        List<AmazonAsin> amasinList = new ArrayList<>();
        amasinList.add(new AmazonAsin("1", "1", "1"));
        amasinList.add(new AmazonAsin("2", "2", "2"));
        amasinList.add(new AmazonAsin("3", "3", "3"));
        amasinList.add(new AmazonAsin("4", "4", "4"));

        // 填充配置：开启强制新行（模板中要使用“单行区域”语法：如 `{.sku}`）
        FillConfig fillConfig = FillConfig.builder()
                .forceNewRow(true)
                .build();

        // 执行填充到第三个 sheet
        writer.fill(amasinList, fillConfig, sheet);

        writer.finish();
    }


    public static void main(String[] args) {
        TestDemo testDemo = new TestDemo();
        testDemo.salaryList();
    }

}
