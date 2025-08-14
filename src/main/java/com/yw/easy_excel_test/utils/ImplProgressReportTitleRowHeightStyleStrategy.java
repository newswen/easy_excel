package com.yw.easy_excel_test.utils;

import org.apache.poi.ss.usermodel.Row;
 
import com.alibaba.excel.write.style.row.AbstractRowHeightStyleStrategy;
 
/**
 * 实施进度报表 行高设置
 * @author 
 * @date 2024/9/19 星期四
 * @since JDK 17
 */
public class ImplProgressReportTitleRowHeightStyleStrategy extends AbstractRowHeightStyleStrategy {
 
    @Override
    protected void setHeadColumnHeight(Row row, int relativeRowIndex) {
        // 设置主标题行高为 25.5
        if(relativeRowIndex == 0){
            // 25.5*20
            row.setHeight((short) 510);
        }
    }
 
    @Override
    protected void setContentColumnHeight(Row row, int relativeRowIndex) {
        // 设置主标题行高为 25.5
        if(relativeRowIndex == 0){
            // 25.5*20
            row.setHeight((short) 510);
        }
    }
}