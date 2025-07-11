package com.yw.easy_excel_test.controller;

import com.yw.easy_excel_test.service.ICspExcelExportService;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import java.io.IOException;

@RestController
@RequestMapping("/basic/download")
public class CspDownloadCenter {

    @Resource
    private ICspExcelExportService excelExportService;

    @GetMapping("/excelExportTest")
    public void excelExportTest() throws IOException {
        excelExportService.excelExportTest();
    }

}
