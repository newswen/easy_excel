package com.yw.easy_excel_test.fa_eu;

import com.alibaba.excel.EasyExcel;
import com.yw.easy_excel_test.entity.StockMovementModelVO;
import com.yw.easy_excel_test.entity.TransferYCheckResult;
import com.yw.easy_excel_test.simple.StockMovementListener;
import com.yw.easy_excel_test.simple.ZhuZi;
import lombok.extern.slf4j.Slf4j;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * @Author: yw
 * @Date: 2025/6/5 16:38
 * @Description:
 **/
@Slf4j
public class FaEuServiceImpl {

    public static void readDemo() {
        String fileName = "costway-it_202505.xlsx";
        StockMovementListener zhuZiListener = new StockMovementListener();
        // 读取指定路径的文件，并转换为ZhuZi对象，默认会读取第一个sheet单元
        EasyExcel.read(fileName, StockMovementModelVO.class, zhuZiListener).sheet().doRead();
        List<StockMovementModelVO> zhuZis = zhuZiListener.getData();
        //最终返回
        List<StockMovementModelVO> tempt = new ArrayList<>();
        //通过sku聚合
        Map<String, List<StockMovementModelVO>> skuMap = zhuZis.stream().collect(Collectors.groupingBy(StockMovementModelVO::getSku, Collectors.collectingAndThen(
                Collectors.toList(),
                list -> {
                    list.sort((a, b) -> Double.compare(b.getTransferY(), a.getTransferY())); // 按transfe_y降序
                    return list;
                }
        )));
        for (String sku : skuMap.keySet()) {
            if (sku.equals("AU10035BL")) {
                System.out.println(sku);
            }
            List<StockMovementModelVO> stockMovementModelVOS = skuMap.get(sku);
            //通过invoiceNumber聚合
            Map<String, List<StockMovementModelVO>> invoiceNumberMap = stockMovementModelVOS.stream().collect(Collectors.groupingBy(StockMovementModelVO::getInvoiceNumber));
            //如果该invoiceNumberMap的每一个键对应的List数量相等，就获取对应List元素集合transfer总和
            TransferYCheckResult transferYCheckResult = checkTransferYSumAndListSize(invoiceNumberMap);
            if (transferYCheckResult.getTransferYSum() == 0) {
                tempt.addAll(stockMovementModelVOS);
                continue;
            }
            //每一个列表大小 10
            int size = transferYCheckResult.getListSize();
            //列表数量2
            int listSize = invoiceNumberMap.size();
            if (listSize < size && listSize != 1 && size != 2) {
                tempt.addAll(stockMovementModelVOS);
                continue;
            }
            //最终结算List（判断最终去重后，inbound总和和转运总和是否一致)
            List<StockMovementModelVO> tempt2 = new ArrayList<>();
            //总Transfer大小
            int transferYSum = transferYCheckResult.getTransferYSum();
            int i = 0;
            int y = 0;
            for (String invoiceNumber : invoiceNumberMap.keySet()) {
                List<StockMovementModelVO> stockMovementModelVOS2 = invoiceNumberMap.get(invoiceNumber);
                if (listSize >= size) {
                    if (listSize - i < size) {
                        tempt2.add(stockMovementModelVOS2.get(++y));
                        continue;
                    }
                    tempt2.add(stockMovementModelVOS2.get(y));
                    i++;
                    continue;
                }
                //针对y不为0的时候
                if (listSize == 1 && size == 2) {
                    StockMovementModelVO stockMovementModelVO = stockMovementModelVOS2.get(y++);
                    if (stockMovementModelVO.getInbound() == transferYSum) {
                        stockMovementModelVO.setInbound(stockMovementModelVO.getTransferY());
                    }
                    tempt2.add(stockMovementModelVO);
                }
            }
            int sum = tempt2.stream().mapToInt(StockMovementModelVO::getInbound).sum();
            if (sum == transferYSum) {
                log.info("最终去重后{}", tempt2);
                tempt2.forEach(stockMovementModelVO -> {
                    stockMovementModelVO.setName("okok");
                });
                tempt.addAll(tempt2);
            } else {
                tempt.addAll(stockMovementModelVOS);
            }
        }
        // 可以写绝对路径，没有绝对路径默认放在当前目录下
        fileName = "测试-" + System.currentTimeMillis() + ".xlsx";
        EasyExcel.write(fileName, StockMovementModelVO.class).sheet("测试").doWrite(tempt);
        System.out.println("读取excel文件结束，总计解析到" + tempt.size() + "条数据！");
    }

    public static TransferYCheckResult checkTransferYSumAndListSize(Map<String, List<StockMovementModelVO>> invoiceNumberMap) {
        Set<Integer> transferYSums = new HashSet<>();
        Set<Integer> listSizes = new HashSet<>();

        for (List<StockMovementModelVO> list : invoiceNumberMap.values()) {
            int sum = list.stream()
                    .mapToInt(vo -> vo.getTransferY() != null ? vo.getTransferY() : 0)
                    .sum();
            transferYSums.add(sum);
            listSizes.add(list.size());
        }

        if (transferYSums.size() == 1) {
            int transferY = transferYSums.iterator().next();
            int size = listSizes.size() == 1 ? listSizes.iterator().next() : -1; // -1 表示 list size 不一致
            return new TransferYCheckResult(transferY, size);
        } else {
            return new TransferYCheckResult(0, 0);
        }
    }


    public static void main(String[] args) {
        readDemo();
    }
}
