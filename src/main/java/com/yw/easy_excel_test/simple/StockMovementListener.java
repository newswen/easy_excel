package com.yw.easy_excel_test.simple;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.yw.easy_excel_test.entity.StockMovementModelVO;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author: yw
 * @Date: 2025/6/5 16:47
 * @Description:
 **/
@Slf4j
public class StockMovementListener extends AnalysisEventListener<StockMovementModelVO> {
    private final List<StockMovementModelVO> data = new ArrayList<>();

    /*
     * 每解析一条数据都会触发一次invoke()方法
     * */
    @Override
    public void invoke(StockMovementModelVO zhuZi, AnalysisContext analysisContext) {
        if (zhuZi.getFlag() == null) {
            return;
        }
        log.info("成功解析到一条数据：{}", zhuZi);
        data.add(zhuZi);
    }

    /*
     * 当一个excel文件所有数据解析完成后，会触发此方法
     * */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.info("所有数据都已解析完毕！");
    }

    public List<StockMovementModelVO> getData() {
        return data;
    }
}
