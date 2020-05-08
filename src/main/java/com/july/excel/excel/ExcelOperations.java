package com.july.excel.excel;

import com.july.excel.constant.ExcelGlobalConstants;
import com.july.excel.entity.ExcelData;
import com.july.excel.entity.ExcelField;
import com.july.excel.utils.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static com.july.excel.utils.ExcelUtils.*;

/**
 * Excel 导入导出操作相关工具类
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-06 17:42
 **/
@Slf4j
public class ExcelOperations {

    private static final ThreadLocal<SimpleDateFormat> simpleDateFormatThreadLocal = new ThreadLocal<>();
    private static final ThreadLocal<DecimalFormat> decimalFormatThreadLocal = new ThreadLocal<>();
    private static final ThreadLocal<ExcelOperations> UTILS_THREAD_LOCAL = new ThreadLocal<>();

    /**
     * 导出excel数据
     * @param excelData
     * @param excelClass
     * @param httpServletResponse
     * @return java.lang.Boolean
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static Boolean exportForExcelsOptimize(ExcelData excelData, Class<?> excelClass, HttpServletResponse httpServletResponse) {
        long startTime = System.currentTimeMillis();
        log.info("===> Excel tool class export start run!");
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        OutputStream outputStream = null;
        SXSSFRow sxssfRow = null;
        try {
            List<Field> hasExcelFieldList = BeanUtils.getExcelFields(excelClass, excelData.getIgnores());
            setDataList(sxssfWorkbook, sxssfRow, excelData, hasExcelFieldList);
            ExcelUtils.setExcelResponse(sxssfWorkbook, outputStream, excelData, httpServletResponse);
        } catch (Exception e) {
            e.printStackTrace();
        }
        log.info("===> Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
        return true;
    }

    /**
     * excel 模板数据导入
     * @param book      Workbook对象（不可为空）
     * @param excelData
     * @return java.util.List<java.util.List < java.util.LinkedHashMap < java.lang.String, java.lang.String>>>
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static <R> List<R> importForExcelData(Workbook book, Class<R> excelClass, ExcelData excelData) {
        List<Field> excelFields = BeanUtils.getExcelFields(excelClass, null);
        Map<String, Field> hasAnnotationFieldMap = new HashMap<>();
        excelFields.stream().forEach(field -> {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            hasAnnotationFieldMap.put(excelField.value(), field);
        });
        R object = null;
        ExcelOperations excelOperations = UTILS_THREAD_LOCAL.get();
        if (excelOperations == null) {
            excelOperations = new ExcelOperations();
            UTILS_THREAD_LOCAL.set(excelOperations);
        }
        long startTime = System.currentTimeMillis();
        log.info("===> Excel tool class export start run!");
        List<String> excelTitles = new ArrayList<>();
        try {
            List<R> returnDataList = new ArrayList<>();
            for (int k = 0; k <= excelData.getSheetName().split(",").length - 1; k++) {
                //得到第K个工作表对象、得到第K个工作表中的总行数。
                Sheet sheet = book.getSheetAt(k);
                int rowCount = sheet.getLastRowNum() + 1;
                Row valueRow = null;

                //excel首行标题
                excelTitles = getExcelTitle(sheet, excelTitles, excelData.getExportStartRowNum());

                int irow = 1;
                //第k个工作表:从开始获取数据、默认第一行开始获取。
                if (!CollectionUtils.isEmpty(excelData.getExcelReadDataList()) && excelData.getExcelReadDataList().get(k + 1) != null) {
                    irow = Integer.valueOf(excelData.getExcelReadDataList().get(k + 1).getRowNum().toString()) - 1;
                }
                //第k个工作表:数据获取。
                for (int i = irow; i < rowCount; i++) {
                    try {
                        object = excelClass.newInstance();
                    } catch (InstantiationException | IllegalAccessException e) {
                        throw new Exception("Excel model init failure, " + e.getMessage());
                    }
                    valueRow = sheet.getRow(i);
                    if (valueRow == null) {
                        continue;
                    }
                    //第k个工作表:获取列数据。
                    for (int j = 0; j < valueRow.getLastCellNum(); j++) {
                        Field field = hasAnnotationFieldMap.get(excelTitles.get(j));
                        if (field == null) {
                            throw new Exception();
                        }
                        BeanUtils.setFieldValue(object, field, getCellVal(valueRow.getCell(j), excelData));
                    }
                    returnDataList.add(object);
                }
            }
            log.info("===> Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
            return returnDataList;
        } catch (Exception e) {
            log.debug("===> Exception Message：Excel tool class export exception !");
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 获取excel首行的标题
     * @param sheet
     * @param excelTitles
     * @param startRowNum
     * @return java.util.List<java.lang.String>
     * @author zengxueqi
     * @since 2020/5/8
     */
    public static List<String> getExcelTitle(Sheet sheet, List<String> excelTitles, Integer startRowNum) {
        //获取第一行
        Row titlerow = sheet.getRow(startRowNum);
        //有多少列
        int cellNum = titlerow.getLastCellNum();
        for (int i = 0; i < cellNum; i++) {
            //根据索引获取对应的列
            Cell cell = titlerow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            //设置列的类型是字符串
            cell.setCellType(CellType.STRING);
            excelTitles.add(cell.getStringCellValue());
        }
        return excelTitles;
    }

    /**
     * 功能描述: 获取Excel单元格中的值并且转换java类型格式
     * @param cell
     * @param excelData
     * @return java.lang.String
     * @author zengxueqi
     * @since 2020/5/7
     */
    private static String getCellVal(Cell cell, ExcelData excelData) {
        String val = null;
        if (cell != null) {
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        val = DateUtils.getDateFormat(simpleDateFormatThreadLocal, excelData.getExpectDateFormatStr()).format(cell.getDateCellValue());
                    } else {
                        val = NumberUtils.getDecimalFormat(decimalFormatThreadLocal, excelData.getNumeralFormat()).format(cell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    if (cell.getStringCellValue().trim().length() >= ExcelGlobalConstants.DATE_LENGTH
                            && DateUtils.verificationDate(cell.getStringCellValue(), excelData.getDateFormatStr())) {
                        val = DateUtils.strToDateFormat(cell.getStringCellValue(), excelData.getDateFormatStr(),
                                excelData.getExpectDateFormatStr());
                    } else {
                        val = cell.getStringCellValue();
                    }
                    break;
                case BOOLEAN:
                    val = String.valueOf(cell.getBooleanCellValue());
                    break;
                case BLANK:
                    val = cell.getStringCellValue();
                    break;
                case ERROR:
                    val = "错误";
                    break;
                case FORMULA:
                    try {
                        val = String.valueOf(cell.getStringCellValue());
                    } catch (IllegalStateException e) {
                        val = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                default:
                    val = cell.getRichStringCellValue() == null ? null : cell.getRichStringCellValue().toString();
            }
        } else {
            val = "";
        }
        return val;
    }

}

