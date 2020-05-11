package com.july.excel.excel;

import com.july.excel.constant.ExcelGlobalConstants;
import com.july.excel.entity.ExcelData;
import com.july.excel.entity.ExcelField;
import com.july.excel.entity.ExcelReadData;
import com.july.excel.exception.BnException;
import com.july.excel.utils.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

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
     * @param file
     * @param excelData
     * @return java.util.List<java.util.List < java.util.LinkedHashMap < java.lang.String, java.lang.String>>>
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static <R> List<R> importForExcelData(MultipartFile file, Class<R> excelClass, ExcelData excelData) {
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
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            List<R> returnDataList = new ArrayList<>();
            for (int k = 0; k <= excelData.getSheetName().split(",").length - 1; k++) {
                //得到第K个工作表对象、得到第K个工作表中的总行数。
                Sheet sheet = workbook.getSheetAt(k);
                int rowCount = sheet.getLastRowNum() + 1;
                Row valueRow = null;

                //excel首行标题
                List<String> excelTitles = getExcelTitle(sheet, excelData.getExcelTitleRowNum(), k);
                BnException.of(CollectionUtils.isEmpty(excelTitles), "读取excel标题错误！");

                int irow = 1;
                //第k个工作表:从开始获取数据、默认第一行开始获取。
                irow = getExcelStartRowNum(excelData.getExcelReadDataList(), k);

                //先注释，后续完善
                /*Map<String, PictureData> imgMaplist = null;
                //判断用07还是03的方法获取图片
                if (file.getOriginalFilename().endsWith(".xls")) {
                    imgMaplist = ImageUtils.getPictures1((HSSFSheet) sheet);
                } else if (file.getOriginalFilename().endsWith(".xlsx")) {
                    imgMaplist = ImageUtils.getPictures2((XSSFSheet) sheet);
                }
                if (!CollectionUtils.isEmpty(imgMaplist)) {
                    HashMap addMap = new HashMap();
                    HashMap addValMap = new HashMap();
                    addMap.put("imgMaplist", Class.forName("java.util.Map"));
                    addValMap.put("imgMaplist", imgMaplist);
                    object = (R) new ClassUtils().dynamicClass(object, addMap, addValMap);
                }*/

                //第k个工作表:数据获取。
                for (int i = irow; i < rowCount; i++) {
                    try {
                        object = excelClass.newInstance();
                    } catch (InstantiationException | IllegalAccessException e) {
                        throw BnException.on("Exception Message：Excel model init failure, " + e.getMessage());
                    }
                    valueRow = sheet.getRow(i);
                    if (valueRow == null) {
                        continue;
                    }
                    //第k个工作表:获取列数据。
                    for (int j = 0; j < valueRow.getLastCellNum(); j++) {
                        Field field = hasAnnotationFieldMap.get(excelTitles.get(j));
                        BnException.of(field == null, "excel标题解析失败！");
                        BeanUtils.setFieldValue(object, field, getCellVal(valueRow.getCell(j), excelData));
                    }
                    returnDataList.add(object);
                }
            }
            return returnDataList;
        } catch (Exception e) {
            throw BnException.on("Exception Message：Excel tool class export exception !");
        }
    }

    /**
     * 获取excel首行的标题
     * @param sheet
     * @param excelTitleRowNums
     * @param sheetNum
     * @return java.util.List<java.lang.String>
     * @author zengxueqi
     * @since 2020/5/8
     */
    public static List<String> getExcelTitle(Sheet sheet, List<ExcelReadData> excelTitleRowNums, int sheetNum) {
        int startRowNum = 0;
        if (!CollectionUtils.isEmpty(excelTitleRowNums)) {
            for (ExcelReadData excelTitleRowNum : excelTitleRowNums) {
                if (excelTitleRowNum.getSheetNum().intValue() == sheetNum) {
                    startRowNum = excelTitleRowNum.getRowNum();
                    break;
                }
            }
        }
        List<String> excelTitles = new ArrayList<>();
        //获取第一行
        Row titlerow = sheet.getRow(startRowNum);
        //有多少列
        int cellNum = titlerow.getLastCellNum();
        for (int i = 0; i < cellNum; i++) {
            //根据索引获取对应的列
            Cell cell = titlerow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
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
        String value = null;
        if (cell != null) {
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        value = DateUtils.getDateFormat(simpleDateFormatThreadLocal, excelData.getExpectDateFormatStr()).format(cell.getDateCellValue());
                    } else {
                        value = NumberUtils.getDecimalFormat(decimalFormatThreadLocal, excelData.getNumeralFormat()).format(cell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    if (cell.getStringCellValue().trim().length() >= ExcelGlobalConstants.DATE_LENGTH
                            && DateUtils.verificationDate(cell.getStringCellValue(), excelData.getExpectDateFormatStr())) {
                        value = DateUtils.strToDateFormat(cell.getStringCellValue(), excelData.getExpectDateFormatStr());
                    } else {
                        value = cell.getStringCellValue();
                    }
                    break;
                case BOOLEAN:
                    value = String.valueOf(cell.getBooleanCellValue());
                    break;
                case BLANK:
                    value = cell.getStringCellValue();
                    break;
                case ERROR:
                    value = "错误";
                    break;
                case FORMULA:
                    try {
                        value = String.valueOf(cell.getStringCellValue());
                    } catch (IllegalStateException e) {
                        value = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                default:
                    value = cell.getRichStringCellValue() == null ? null : cell.getRichStringCellValue().toString();
            }
        } else {
            value = "";
        }
        return value;
    }

    /**
     * 获取sheet解析数据的开始行数
     * @param excelReadDataList
     * @param sheetNum
     * @return java.lang.Integer
     * @author zengxueqi
     * @since 2020/5/9
     */
    public static Integer getExcelStartRowNum(List<ExcelReadData> excelReadDataList, int sheetNum) {
        if (!CollectionUtils.isEmpty(excelReadDataList)) {
            for (ExcelReadData excelReadData : excelReadDataList) {
                if (excelReadData.getSheetNum().intValue() == sheetNum) {
                    return excelReadData.getRowNum();
                }
            }
        }
        return 1;
    }

}

