package com.july.excel.excel;


import com.july.excel.constant.ExcelGlobalConstants;
import com.july.excel.entity.ExcelData;
import com.july.excel.utils.DateUtils;
import com.july.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static com.july.excel.utils.CommonsUtils.*;

/**
 * Excel 导入导出操作相关工具类
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-06 17:42
 **/
public class ExcelUtils {

    private static Logger log = LoggerFactory.getLogger(ExcelUtils.class);
    private static final ThreadLocal<SimpleDateFormat> fmt = new ThreadLocal<>();
    private static final ThreadLocal<DecimalFormat> df = new ThreadLocal<>();
    private static final ThreadLocal<ExcelUtils> UTILS_THREAD_LOCAL = new ThreadLocal<>();

    public static SimpleDateFormat getDateFormat(String expectDateFormatStr) {
        SimpleDateFormat format = fmt.get();
        if (format == null) {
            //默认格式日期： "yyyy-MM-dd"
            format = new SimpleDateFormat(expectDateFormatStr, Locale.getDefault());
            fmt.set(format);
        }
        return format;
    }

    public static DecimalFormat getDecimalFormat(String numeralFormat) {
        DecimalFormat format = df.get();
        if (format == null) {
            //默认数字格式： "#.######" 六位小数点
            format = new DecimalFormat(numeralFormat);
            df.set(format);
        }
        return format;
    }

    public static final ExcelUtils initialization() {
        ExcelUtils excelUtils = UTILS_THREAD_LOCAL.get();
        if (excelUtils == null) {
            excelUtils = new ExcelUtils();
            UTILS_THREAD_LOCAL.set(excelUtils);
        }
        return excelUtils;
    }

    /**
     * web 响应（response）
     * Excel导出：有样式（行、列、单元格样式）、自适应列宽
     * 功能描述: excel 数据导出、导出模板
     * 更新日志:
     * 1.response.reset();注释掉reset，否在会出现跨域错误。
     * 2.新增导出多个单元。[2018-08-08]
     * 3.poi官方建议大数据量解决方案：SXSSFWorkbook。
     * 4.自定义下拉列表：对每个单元格自定义下拉列表。
     * 5.数据遍历方式换成数组(效率较高)。
     * 6.可提供模板下载。
     * 7.每个表格的大标题
     * 8.自定义列宽：对每个单元格自定义列宽
     * 9.自定义样式：对每个单元格自定义样式
     * 10.自定义单元格合并：对每个单元格合并
     * 11.固定表头
     * 12.自定义样式：单元格自定义某一列或者某一行样式
     * 13.忽略边框(默认是有边框)
     * 14.函数式编程换成面向对象编程
     * 15.单表百万数据量导出时样式设置过多，导致速度慢（行、列、单元格样式暂时控制10万行、超过无样式）
     * 版  本:
     * 1.apache poi 3.17
     * 2.apache poi-ooxml  3.17
     * @return
     */
    public Boolean exportForExcelsOptimize(ExcelData excelData) {
        long startTime = System.currentTimeMillis();
        log.info("===> Excel tool class export start run!");
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        OutputStream outputStream = null;
        SXSSFRow sxssfRow = null;
        try {
            //设置数据
            setDataList(sxssfWorkbook, sxssfRow, excelData);
            //io 响应
            setIo(sxssfWorkbook, outputStream, excelData);
        } catch (Exception e) {
            e.printStackTrace();
        }
        log.info("===> Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
        return true;
    }

    /**
     * Excel导出：无样式（行、列、单元格样式）、自适应列宽
     * web 响应（response）
     * @return
     */
    public Boolean exportForExcelsNoStyle(ExcelData excelData) {
        long startTime = System.currentTimeMillis();
        log.info("===> Excel tool class export start run!");
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        OutputStream outputStream = null;
        SXSSFRow sxssfRow = null;
        try {
            setDataListNoStyle(sxssfWorkbook, sxssfRow, excelData);
            setIo(sxssfWorkbook, outputStream, excelData);
        } catch (Exception e) {
            e.printStackTrace();
        }
        log.info("===> Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
        return true;
    }


    /**
     * 功能描述:
     * 1.excel 模板数据导入。
     * <p>
     * 更新日志:
     * 1.共用获取Excel表格数据。
     * 2.多单元数据获取。
     * 3.多单元从第几行开始获取数据[2018-09-20]
     * 4.多单元根据那些列为空来忽略行数据[2018-10-22]
     * <p>
     * 版  本:
     * 1.apache poi 3.17
     * 2.apache poi-ooxml  3.17
     * @param book           Workbook对象（不可为空）
     * @param sheetName      多单元数据获取（不可为空）
     * @param indexMap       多单元从第几行开始获取数据，默认从第二行开始获取（可为空，如 hashMapIndex.put(1,3); 第一个表格从第三行开始获取）
     * @param continueRowMap 多单元根据那些列为空来忽略行数据（可为空，如 mapContinueRow.put(1,new Integer[]{1, 3}); 第一个表格从1、3列为空就忽略）
     * @return
     */
    @SuppressWarnings({"deprecation", "rawtypes"})
    public static List<List<LinkedHashMap<String, String>>> importForExcelData(Workbook book, String[] sheetName, HashMap indexMap, HashMap continueRowMap) {
        long startTime = System.currentTimeMillis();
        log.info("===> Excel tool class export start run!");
        ExcelData excelData = new ExcelData();
        try {
            List<List<LinkedHashMap<String, String>>> returnDataList = new ArrayList<>();
            for (int k = 0; k <= sheetName.length - 1; k++) {
                //得到第K个工作表对象、得到第K个工作表中的总行数。
                Sheet sheet = book.getSheetAt(k);
                int rowCount = sheet.getLastRowNum() + 1;
                Row valueRow = null;

                List<LinkedHashMap<String, String>> rowListValue = new ArrayList<>();
                LinkedHashMap<String, String> cellHashMap = null;

                int irow = 1;
                //第n个工作表:从开始获取数据、默认第一行开始获取。
                if (indexMap != null && indexMap.get(k + 1) != null) {
                    irow = Integer.valueOf(indexMap.get(k + 1).toString()) - 1;
                }
                //第n个工作表:数据获取。
                for (int i = irow; i < rowCount; i++) {
                    valueRow = sheet.getRow(i);
                    if (valueRow == null) {
                        continue;
                    }
                    //第n个工作表:从开始列忽略数据、为空就跳过。
                    if (continueRowMap != null && continueRowMap.get(k + 1) != null) {
                        int continueRowCount = 0;
                        Integer[] continueRow = (Integer[]) continueRowMap.get(k + 1);
                        for (int w = 0; w <= continueRow.length - 1; w++) {
                            Cell valueRowCell = valueRow.getCell(continueRow[w] - 1);
                            if (valueRowCell == null || StringUtils.isEmpty(valueRowCell.toString())) {
                                continueRowCount = continueRowCount + 1;
                            }
                        }
                        if (continueRowCount == continueRow.length) {
                            continue;
                        }
                    }
                    cellHashMap = new LinkedHashMap<>();

                    //第n个工作表:获取列数据。
                    for (int j = 0; j < valueRow.getLastCellNum(); j++) {
                        cellHashMap.put(Integer.toString(j), getCellVal(valueRow.getCell(j), excelData));
                    }
                    if (cellHashMap.size() > 0) {
                        rowListValue.add(cellHashMap);
                    }
                }
                returnDataList.add(rowListValue);
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
     * response 响应
     * @param sxssfWorkbook
     * @param outputStream
     * @param excelData
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    private static void setIo(SXSSFWorkbook sxssfWorkbook, OutputStream outputStream, ExcelData excelData) throws Exception {
        try {
            if (excelData.getResponse() != null) {
                excelData.getResponse().setHeader("Charset", "UTF-8");
                excelData.getResponse().setHeader("Content-Type", "application/force-download");
                excelData.getResponse().setHeader("Content-Type", "application/vnd.ms-excel");
                excelData.getResponse().setHeader("Content-disposition", "attachment; filename="
                        + URLEncoder.encode(StringUtils.isEmpty(excelData.getFileName()) ? excelData.getSheetName()[0]
                        : excelData.getFileName(), "utf8") + ".xlsx");
                excelData.getResponse().flushBuffer();
                outputStream = excelData.getResponse().getOutputStream();
            }
            writeAndColse(sxssfWorkbook, outputStream);
        } catch (Exception e) {
            e.getSuppressed();
        }
    }

    /**
     * 功能描述: 获取Excel单元格中的值并且转换java类型格式
     * @param cell
     * @return
     */
    private static String getCellVal(Cell cell, ExcelData excelData) {
        String val = null;
        if (cell != null) {
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        val = getDateFormat(excelData.getExpectDateFormatStr()).format(cell.getDateCellValue());
                    } else {
                        val = getDecimalFormat(excelData.getNumeralFormat()).format(cell.getNumericCellValue());
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

