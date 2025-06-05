//package com.yw.easy_excel_test.utils;
//
//import com.alibaba.excel.support.ExcelTypeEnum;
//import com.yw.easy_excel_test.entity.CompetitorProductModelVO;
//
//import java.io.IOException;
//import java.util.ArrayList;
//import java.util.List;
//
///**
// * @Author: yw
// * @Date: 2025/5/15 17:06
// * @Description:
// **/
//public class TestEasy {
//
//    public static void main(String[] args) {
//        List<CompetitorProductModelVO> competitorProductModelVOList = new ArrayList<>();
//        competitorProductModelVOList.add(CompetitorProductModelVO.builder()
//                .productLink("https://www.taobao.com")
//                .productImageUrl(null)
//                .productTitle("淘宝")
//                .productFivePoints("淘宝five")
//                .reviewCount(100)
//                .build());
//        try {
//            ExcelUtil.exportExcel(null, CompetitorProductModelVO.class, competitorProductModelVOList, "竞品信息", ExcelTypeEnum.XLSX, httpServletResponse);
//        } catch (IOException e) {
//            throw new RuntimeException("获取初步竞品数据excel失败");
//        }
//    }
//}
