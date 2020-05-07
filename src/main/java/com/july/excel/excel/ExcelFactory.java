package com.july.excel.excel;

import com.july.excel.constant.ExcelGlobalConstants;
import com.july.excel.entity.ExcelData;
import com.july.excel.entity.ExcelReadData;
import com.july.excel.utils.DateUtils;
import com.july.excel.utils.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.LinkedHashMap;
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
     * @param fileName
     * @param sheetName
     * @param excelReadDataList
     * @return java.util.List<java.util.List < java.util.LinkedHashMap < java.lang.String, java.lang.String>>>
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static List<List<LinkedHashMap<String, String>>> importExcelData(InputStream inputStream, String fileName,
                                                                            String[] sheetName, List<ExcelReadData> excelReadDataList) throws Exception {
        String extName = fileName.substring(fileName.lastIndexOf('.') + 1);
        Workbook workbook = null;
        if (ExcelGlobalConstants.XLS.equals(extName)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (ExcelGlobalConstants.XLSX.equals(extName)) {
            workbook = new XSSFWorkbook(inputStream);
        } else {
            throw new Exception("无法识别的Excel文件，请重新选择Excel模版进行导入!");
        }
        return ExcelOperations.importForExcelData(workbook, sheetName, excelReadDataList);
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
