package com.yw.easy_excel_test.utils;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelIgnoreUnannotated;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.yw.easy_excel_test.entity.ExcelFileTypeEnum;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.Collection;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelUtil {
    /*
     * 校验excel文件的表头，与数据模型类的映射关系是否匹配
     * */
    public static void validateExcelTemplate(Map<Integer, String> headMap, Class<?> clazz, Field[] fields) {
        Collection<String> headNames = headMap.values();

        // 类上是否存在忽略excel字段的注解
        boolean classIgnore = clazz.isAnnotationPresent(ExcelIgnoreUnannotated.class);
        int count = 0;
        for (Field field : fields) {
            // 如果字段上存在忽略注解，则跳过当前字段
            if (field.isAnnotationPresent(ExcelIgnore.class)) {
                continue;
            }

            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (null == excelProperty) {
                // 如果类上也存在忽略注解，则跳过所有未使用ExcelProperty注解的字段
                if (classIgnore) {
                    continue;
                }
                // 如果检测到既未忽略、又未映射excel列的字段，则抛出异常提示模板不正确
                throw new ExcelAnalysisException("请检查导入的excel文件是否按模板填写!");
            }

            // 校验数据模型类上绑定的名称，是否与excel列名匹配
            String[] value = excelProperty.value();
            String name = value[0];
            if (name != null && 0 != name.length() && !headNames.contains(name)) {
                throw new ExcelAnalysisException("请检查导入的excel文件是否按模板填写!");
            }
            // 更新有效字段的数量
            count++;
        }
        // 最后校验数据模型类的有效字段数量，与读到的excel列数量是否匹配
        if (headMap.size() != count) {
            throw new ExcelAnalysisException("请检查导入的excel文件是否按模板填写!");
        }
    }

    /**
     * 导出excel的通用方法
     *
     * @param inputStream 模板文件
     * @param clazz       导出excel所需的数据模型类；
     * @param excelData   需要导出的数据列表；
     * @param fileName    当前导出的文件名称（不带文件后缀）
     * @param excelType   导出的文件类型（XLSX、XLS、CSV三种）；
     * @param response    网络响应对象。
     * @throws IOException
     */
    public static void exportExcel(InputStream inputStream, Class<?> clazz, List<?> excelData, String fileName, ExcelTypeEnum excelType, HttpServletResponse response) throws IOException {
        // 设置样式
        HorizontalCellStyleStrategy styleStrategy = setCellStyle();
        // 设置文件名及响应类型
        fileName = URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20") + excelType.getValue();

        switch (excelType) {
            case XLS:
                response.setContentType(ExcelFileTypeEnum.XLS.getContentType());
                break;
            case XLSX:
                response.setContentType(ExcelFileTypeEnum.XLSX.getContentType());
                break;
            case CSV:
                response.setContentType(ExcelFileTypeEnum.CSV.getContentType());
                break;
        }
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-Disposition", "attachment;filename*=utf-8''" + fileName);

        try (OutputStream outputStream = response.getOutputStream()) {
            ExcelWriterBuilder writeWork = EasyExcelFactory.write(outputStream, clazz);

            if (inputStream != null) {
                try {
                    writeWork.withTemplate(inputStream)
                            .registerWriteHandler(styleStrategy)
                            .excelType(excelType)
                            .sheet()
                            .doWrite(excelData);
                } finally {
                    // 确保输入流在使用完后关闭
                    inputStream.close();
                }
            } else {
                writeWork.registerWriteHandler(styleStrategy)
                        .excelType(excelType)
                        .sheet()
                        .doWrite(excelData);
            }
        } catch (IOException e) {
            // 捕获IO异常，记录日志
            throw new IOException("Error writing Excel file: " + e.getMessage(), e);
        }
    }

    /*
     * 设置单元格风格
     * */
    public static HorizontalCellStyleStrategy setCellStyle() {
        // 设置表头的样式（背景颜色、字体、居中显示）
        WriteCellStyle headStyle = new WriteCellStyle();
        //设置表头的背景颜色
        headStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        WriteFont headFont = new WriteFont();
        headFont.setFontHeightInPoints((short) 12);
        headFont.setBold(true);
        headStyle.setWriteFont(headFont);
        headStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);

        // 设置Excel内容策略(水平居中)
        WriteCellStyle cellStyle = new WriteCellStyle();
        cellStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
        cellStyle.setWrapped(true);
        return new HorizontalCellStyleStrategy(headStyle, cellStyle);
    }


}
