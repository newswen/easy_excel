package com.yw.easy_excel_test.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.yw.easy_excel_test.entity.CompetitorProductModelVO;
import com.yw.easy_excel_test.entity.excelColumnName;
import com.yw.easy_excel_test.entity.excelContact;
import com.yw.easy_excel_test.utils.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Author: yw
 * @Date: 2025/5/5 09:14
 * @Description:
 **/
@RestController
@Slf4j
@RequestMapping("/api")
public class EasyExcelController {

    @GetMapping("/exportEasyExcel")
    public void exportExcel(HttpServletResponse response) {
        ExcelWriter excelWriter = null;
        try(OutputStream out = response.getOutputStream()) {
            excelWriter = EasyExcel.write(out).build();
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            String fileName = URLEncoder.encode("哈哈哈测试导出", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            WriteSheet dealerSheet = EasyExcel.writerSheet(0, "经销商信息").head(excelColumnName.class).build();
            WriteSheet contactSheet = EasyExcel.writerSheet(1, "联系人").head(excelContact.class).build();
            List<excelColumnName> getCompanyData= getCompany();
            excelWriter.write(getCompanyData, dealerSheet);
            excelWriter.write(getContact(), contactSheet);
            excelWriter.finish();
            out.flush();
        } catch (Exception e) {
            e.printStackTrace();
            log.error("导出失败:" + e.getMessage());
        } finally {
            try {
                excelWriter.finish();
                response.getOutputStream().close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    private void handleExcel(OutputStream out) {
        try {
            ExcelWriter excelWriter = EasyExcelFactory.write(out).build();
            WriteSheet dealerSheet = EasyExcel.writerSheet(0, "经销商信息").head(excelColumnName.class).build();
            WriteSheet contactSheet = EasyExcel.writerSheet(1, "联系人").head(excelColumnName.class).build();
            excelWriter.write(getCompany(), dealerSheet);
            excelWriter.write(getContact(), contactSheet);
        } catch (Exception e) {
            log.error(e.getMessage());
        }
    }

    private List<excelColumnName> getCompany() {
        List<excelColumnName> companyList = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            companyList.add(excelColumnName.builder()
                    .companyName("白小纯公司" + i)
                    .province("上海市")
                    .businessProvinceName("山东省")
                    .businessCityName("临沂市")
                    .businessAreaName("河东区")
                    .entStatus("营业")
                    .netAddress("www.baixiaochun.site")
                    .csdnAddress("https://baixiaochun.blog.csdn.net")
                    .employeeMaxCount("100")
                    .startDate(new Date())
                    .build());
        }
        return companyList;
    }

    private List<excelContact> getContact() {
        List<excelContact> contactList = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            contactList.add(excelContact.builder()
                    .companyName("白小纯公司" + i)
                    .name("白小纯" + i)
                    .mobile("177000000000")
                    .idCard("456224199011111111")
                    .contactPostName("测试后端")
                    .build());
        }
        return contactList;
    }

    @GetMapping("/getInitCompetitorFile")
    public void getInitCompetitorFile(HttpServletResponse httpServletResponse) throws MalformedURLException {
        List<CompetitorProductModelVO> competitorProductModelVOList = new ArrayList<>();
        competitorProductModelVOList.add(CompetitorProductModelVO.builder()
                .productLink("https://www.amazon.com/dp/B01C9D7X")
                .productImageUrl(new URL("https://m.media-amazon.com/images/I/81P9DzrF1NL._AC_UY218_.jpg"))
                .productTitle("豆浆机 Model 1 型号")
                .productFivePoints("【Versatile 6-in-1 Functionality】-This Nut Milk Maker offers six diverse functions, including making nut milk, oat milk, soy milk, preparing fresh juice, boiling water, and self-cleaning. It’s your all-in-one kitchen companion for creating a variety of healthy beverages and dishes;【Easy to Use】-The Nut Milk Maker is designed with a user-friendly interface, making it easy to use. Simply plug in the power adapter, turn on the machine, and select the desired function. The Nut Milk")
                .reviewCount(100)
                .build());
        try {
            ExcelUtil.exportExcel(null, CompetitorProductModelVO.class, competitorProductModelVOList, "竞品信息", ExcelTypeEnum.XLSX, httpServletResponse);
        } catch (IOException e) {
            throw new RuntimeException("获取初步竞品数据excel失败");
        }
    }

}
