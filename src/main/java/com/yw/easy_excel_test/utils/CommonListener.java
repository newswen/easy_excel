package com.yw.easy_excel_test.utils;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class CommonListener<T> extends AnalysisEventListener<T> {
    //创建list集合封装最终的数据
    private final List<T> data;

    // 字段列表
    private final Field[] fields;
    private final Class<T> clazz;
    private boolean validateSwitch = true;

    public CommonListener(Class<T> clazz) {
        fields = clazz.getDeclaredFields();
        this.clazz = clazz;
        this.data = new ArrayList<T>();
    }

    /*
     * 每解析到一行数据都会触发
     * */
    @Override
    public void invoke(T row, AnalysisContext analysisContext) {
        data.add(row);
    }

    /*
     * 读取到excel头信息时触发，会将表头数据转为Map集合
     * */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        // 校验读到的excel表头是否与数据模型类匹配
        if (validateSwitch) {
            ExcelUtil.validateExcelTemplate(headMap, clazz, fields);
        }
    }

    /*
     * 所有数据解析完之后触发
     * */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {}

    /*
     * 关闭excel表头验证
     * */
    public void offValidate() {
        this.validateSwitch = false;
    }

    /*
     * 返回解析到的所有数据
     * */
    public List<T> getData() {
        return data;
    }
}
