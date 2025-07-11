package com.yw.easy_excel_test.controller;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.util.DateUtils;
import com.alibaba.excel.write.style.row.SimpleRowHeightStyleStrategy;
import com.alibaba.fastjson.JSON;
import com.yw.easy_excel_test.entity.OrderDetailEntity;
import com.yw.easy_excel_test.utils.ColumnMergeStrategy;
import com.yw.easy_excel_test.utils.FullCellMergeStrategy;
import org.junit.Test;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

public class TestMergingDemo {

    @Test
    public void exportMerge() {

        // 输出文件路径
        String outFilePath = "exportMerge.xlsx";

        Collection<?> data = data();

        EasyExcelFactory.write(outFilePath, OrderDetailEntity.class)
                // 案例一
                //.registerWriteHandler(new ColumnMergeStrategy(Collections.singletonList(0)))
                // 案例二
                //.registerWriteHandler(new ColumnMergeStrategy(Collections.singletonList(0), Arrays.asList(8, 9)))
                // 案例三
                .registerWriteHandler(new ColumnMergeStrategy(Collections.singletonList(0), Arrays.asList(2, 10, 11)))
                .registerWriteHandler(new ColumnMergeStrategy(Arrays.asList(0, 2), Arrays.asList(8, 9)))
                .sheet("Sheet1").doWrite(data);
    }

    @Test
    public void exportFullMerge() {

        String outFilePath = "D:\\excel-files\\error.xlsx";

        EasyExcelFactory.write(outFilePath)
                .head(getCase4Head())
                // 设置表头行高 30，内容行高 20
                .registerWriteHandler(new SimpleRowHeightStyleStrategy((short)30,(short)20))
                // 自适应表头宽度
                //.registerWriteHandler(new MatchTitleWidthStyleStrategy())
                // 案例四
                .registerWriteHandler(new FullCellMergeStrategy())
                .sheet("Sheet1").doWrite(getCase4Data());
    }

    private Collection<?> data() {

        Map<String, List<String>> productMap = getProductMap();

        List<String> statusList = Arrays.asList("待发货", "已发货", "运输中", "待取货", "已完成");

        List<OrderDetailEntity> dataList = new ArrayList<>();

        Random random = new Random();
        int orderCount = random.nextInt(2) + 2;

        for (int i = 0; i < orderCount; i++) {
            String orderCode = "PL" + DateUtils.format(new Date(), "yyyyMMddHHmm") + "000" + i;
            int orderDetailCount = random.nextInt(10) + 1;

            List<OrderDetailEntity> detailEntities = new ArrayList<>();

            Map<String, BigDecimal> categoryTotalQuantityMap = new HashMap<>();
            Map<String, BigDecimal> categoryTotalPriceMap = new HashMap<>();
            BigDecimal totalQuantity = BigDecimal.ZERO;
            BigDecimal totalPrice = BigDecimal.ZERO;

            for (int j = 0; j < orderDetailCount; j++) {
                String orderDetailCode = UUID.randomUUID().toString();
                String productCategory = new ArrayList<String>(productMap.keySet()).get(random.nextInt(productMap.size()));
                List<String> productList = productMap.get(productCategory);
                String productCode = "SKU" + (random.nextInt(1000)+1000);
                String productName = productList.get(random.nextInt(productList.size())) + "-A" + random.nextInt(50);
                BigDecimal price = new BigDecimal(random.nextInt(2000) + 800);
                BigDecimal quantity = new BigDecimal(random.nextInt(5) + 1);
                String status = statusList.get(random.nextInt(statusList.size()));

                String key = orderCode + "-" + productCategory;
                BigDecimal categoryTotalQuantity = categoryTotalQuantityMap.get(key);
                if (categoryTotalQuantity == null) {
                    categoryTotalQuantity = quantity;
                } else {
                    categoryTotalQuantity = categoryTotalQuantity.add(quantity);
                }
                categoryTotalQuantityMap.put(key, categoryTotalQuantity);

                BigDecimal categoryTotalPrice = categoryTotalPriceMap.get(key);
                if (categoryTotalPrice == null) {
                    categoryTotalPrice = price.multiply(quantity);
                } else {
                    categoryTotalPrice = categoryTotalPrice.add(price.multiply(quantity));
                }
                categoryTotalPriceMap.put(key, categoryTotalPrice);

                totalQuantity = totalQuantity.add(quantity);
                totalPrice = totalPrice.add(price.multiply(quantity));

                detailEntities.add(OrderDetailEntity.builder()
                        .orderCode(orderCode)
                        .orderDetailCode(orderDetailCode)
                        .productCategory(productCategory)
                        .productCode(productCode)
                        .productName(productName)
                        .price(price)
                        .quantity(quantity)
                        .status(status)
                        .build());
            }

            for (OrderDetailEntity item : detailEntities) {
                String key = item.getOrderCode() + "-" + item.getProductCategory();
                item.setCategoryTotalQuantity(categoryTotalQuantityMap.get(key));
                item.setCategoryTotalPrice(categoryTotalPriceMap.get(key));
                item.setTotalQuantity(totalQuantity);
                item.setTotalPrice(totalPrice);
            }
            detailEntities = detailEntities.stream()
                    .sorted(Comparator.comparing(OrderDetailEntity::getOrderCode)
                            .thenComparing(OrderDetailEntity::getProductCategory))
                    .collect(Collectors.toList());

            dataList.addAll(detailEntities);
        }
        return dataList;
    }

    private Map<String, List<String>> getProductMap() {
        Map<String, List<String>> productMap = new HashMap<>();
        // 家电
        List<String> householdList = new ArrayList<>();
        householdList.add("电视机");
        householdList.add("冰箱");
        householdList.add("洗衣机");
        householdList.add("空调");
        productMap.put("家电", householdList);
        // 数码产品
        List<String> digitalList = new ArrayList<>();
        digitalList.add("手机");
        digitalList.add("摄影机");
        digitalList.add("电脑");
        digitalList.add("照相机");
        digitalList.add("投影仪");
        digitalList.add("智能手表");
        productMap.put("数码产品", digitalList);
        // 健身器材
        List<String> gymEquipmentList = new ArrayList<>();
        gymEquipmentList.add("动感单车");
        gymEquipmentList.add("健身椅");
        gymEquipmentList.add("跑步机");
        productMap.put("健身器材", gymEquipmentList);
        return productMap;
    }

    private List<List<String>> getCase4Head() {
        List<List<String>> list = new ArrayList<List<String>>();
        List<String> head0 = new ArrayList<String>();
        head0.add("导出时间");
        head0.add("员工编码");
        head0.add("员工编码");
        List<String> head1 = new ArrayList<String>();
        head1.add("导出时间");
        head1.add("部门信息");
        head1.add("部门编码");
        List<String> head2 = new ArrayList<String>();
        head2.add("导出时间");
        head2.add("部门信息");
        head2.add("部门名称");
        List<String> head3 = new ArrayList<String>();
        head3.add("导出时间");
        head3.add("部门信息");
        head3.add("负责人");
        List<String> head4 = new ArrayList<String>();
        head4.add("导出时间");
        head4.add("个人信息");
        head4.add("用户名称");
        List<String> head5 = new ArrayList<String>();
        head5.add("导出时间");
        head5.add("个人信息");
        head5.add("性别");
        List<String> head6 = new ArrayList<String>();
        head6.add("2024-04-09");
        head6.add("个人信息");
        head6.add("年龄");
        List<String> head7 = new ArrayList<String>();
        head7.add("2024-04-09");
        head7.add("个人信息");
        head7.add("出生日期");
        List<String> head8 = new ArrayList<String>();
        head8.add("2024-04-09");
        head8.add("个人信息");
        head8.add("学历");
        List<String> head9 = new ArrayList<String>();
        head9.add("2024-04-09");
        head9.add("个人信息");
        head9.add("电话号码");
        List<String> head10 = new ArrayList<String>();
        head10.add("2024-04-09");
        head10.add("状态");
        head10.add("状态");

        list.add(head0);
        list.add(head1);
        list.add(head2);
        list.add(head3);
        list.add(head4);
        list.add(head5);
        list.add(head6);
        list.add(head7);
        list.add(head8);
        list.add(head9);
        list.add(head10);

        return list;
    }

    private Collection<?> getCase4Data() {

        List<Map<Integer, Object>> data = new ArrayList<>();
        Map<Integer, Object> map1 = new HashMap<>();
        map1.put(0,"exportTime");
        map1.put(1,"exportTime");
        map1.put(2,"exportTime");
        map1.put(3,"exportTime");
        map1.put(4,"exportTime");
        map1.put(5,"exportTime");
        map1.put(6,"currentData");
        map1.put(7,"currentData");
        map1.put(8,"currentData");
        map1.put(9,"currentData");
        map1.put(10,"currentData");

        Map<Integer, Object> map2 = new HashMap<>();
        map2.put(0,"employeeNo");
        map2.put(1,"deptInfo");
        map2.put(2,"deptInfo");
        map2.put(3,"deptInfo");
        map2.put(4,"userInfo");
        map2.put(5,"userInfo");
        map2.put(6,"userInfo");
        map2.put(7,"userInfo");
        map2.put(8,"userInfo");
        map2.put(9,"userInfo");
        map2.put(10,"status");

        Map<Integer, Object> map3 = new HashMap<>();
        map3.put(0,"employeeNo");
        map3.put(1,"deptCode");
        map3.put(2,"deptName");
        map3.put(3,"deptHead");
        map3.put(4,"username");
        map3.put(5,"gender");
        map3.put(6,"age");
        map3.put(7,"birthday");
        map3.put(8,"educational");
        map3.put(9,"phone");
        map3.put(10,"status");

        data.add(map1);
        data.add(map2);
        data.add(map3);
        return data;
    }
}

