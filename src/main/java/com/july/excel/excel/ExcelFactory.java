package com.july.excel.excel;

import com.july.excel.entity.ExcelData;
import com.july.excel.entity.ExcelReadData;
import com.july.excel.utils.DateUtils;
import com.july.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.List;

/**
 * excel导出工厂
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-07 09:34
 **/
public class ExcelFactory {

    /**
     * 导入excel信息
     * @param inputStream
     * @param excelClass
     * @return java.util.List<java.util.List < java.util.LinkedHashMap < java.lang.String, java.lang.String>>>
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static <R> List<R> importExcelData(InputStream inputStream, Class<R> excelClass, ExcelData excelData) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(inputStream)) {
            return ExcelOperations.importForExcelData(workbook, excelClass, excelData);
        }
    }

    /**
     * 导出excel信息
     * @param excelData
     * @param excelClass
     * @param httpServletResponse
     * @return void
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static void exportExcelData(ExcelData excelData, Class<?> excelClass, HttpServletResponse httpServletResponse) {
        String fileName = StringUtils.isEmpty(excelData.getFileName()) ? "excel-" + DateUtils.getDateFormatStr() : excelData.getFileName();
        excelData.setFileName(fileName);
        String sheetName = "sheet1";
        //必填项--sheet名称（如果是多表格导出、sheetName也要是多个值！）
        excelData.setSheetName(excelData.getSheetName() == null ? sheetName : excelData.getSheetName());
        ExcelOperations.exportForExcelsOptimize(excelData, excelClass, httpServletResponse);
    }

}
