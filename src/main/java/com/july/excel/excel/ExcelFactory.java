package com.july.excel.excel;

import com.july.excel.entity.ExcelData;
import com.july.excel.exception.BnException;
import com.july.excel.utils.DateUtils;
import com.july.excel.utils.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
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
     * @param file
     * @param excelClass
     * @return java.util.List<java.util.List < java.util.LinkedHashMap < java.lang.String, java.lang.String>>>
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static <R> List<R> importExcelData(MultipartFile file, Class<R> excelClass, ExcelData excelData) throws Exception {
        String fileName = file.getOriginalFilename();
        BnException.of(!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$"), "上传文件格式不正确！");
        return ExcelOperations.importForExcelData(file, excelClass, excelData);
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
        ExcelOperations.exportForExcelsOptimize(excelData, excelClass, httpServletResponse);
    }

}
