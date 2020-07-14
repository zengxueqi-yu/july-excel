package com.july.excel.utils;

import com.july.excel.constant.ExcelGlobalConstants;
import com.july.excel.property.ExcelData;
import com.july.excel.property.ExcelDropDown;
import com.july.excel.property.ExcelField;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.*;

import static org.apache.poi.ss.util.CellUtil.createCell;

/**
 * 集合操作工具
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-06 17:42
 **/
public class ExcelUtils {

    private static Logger log = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 流操作
     * @param inStream
     * @return byte[]
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static byte[] readInputStream(InputStream inStream) throws Exception {
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        //创建一个Buffer字符串
        byte[] buffer = new byte[1024];
        //每次读取的字符串长度，如果为-1，代表全部读取完毕
        int len = 0;
        //使用一个输入流从buffer里把数据读取出来
        while ((len = inStream.read(buffer)) != -1) {
            //用输出流往buffer里写入数据，中间参数代表从哪个位置开始读，len代表读取的长度
            outStream.write(buffer, 0, len);
        }
        //关闭输入流
        inStream.close();
        //把outStream里的数据写入内存
        return outStream.toByteArray();
    }

    /**
     * 设置数据：无样式（行、列、单元格样式）
     * @param sxssfWorkbook
     * @param sxssfRow
     * @param excelData
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDataList(SXSSFWorkbook sxssfWorkbook, SXSSFRow sxssfRow, ExcelData excelData, List<Field> excelFields) {
        if (CollectionUtils.isEmpty(excelData.getExcelData())) {
            log.debug("===> Exception Message：Export data(type:List<List<String[]>>) cannot be empty!");
        }
        if (excelData.getSheetName() == null) {
            log.debug("===> Exception Message：Export sheet(type:String[]) name cannot be empty!");
        }
        int k = 0;
        SXSSFSheet sxssfSheet = sxssfWorkbook.createSheet();
        sxssfWorkbook.setSheetName(k, excelData.getSheetName());
        CellStyle cellStyle = sxssfWorkbook.createCellStyle();
        XSSFFont font = (XSSFFont) sxssfWorkbook.createFont();
        SXSSFDrawing sxssfDrawing = sxssfSheet.createDrawingPatriarch();

        int jRow = 0;
        //自定义：大标题
        jRow = ExcelStyleUtils.setLabelName(jRow, sxssfWorkbook, excelData.getLabelName(), sxssfRow, sxssfSheet, excelFields);

        //自定义：每个单元格自定义合并单元格：对每个单元格自定义合并单元格（看该方法说明）
        if (!CollectionUtils.isEmpty(excelData.getExcelRegions())) {
            ExcelStyleUtils.setMergedRegion(sxssfSheet, excelData.getExcelRegions());
        }
        //自定义：每个单元格自定义下拉列表：对每个单元格自定义下拉列表（看该方法说明）
        if (!CollectionUtils.isEmpty(excelData.getExcelDropDowns())) {
            setDropDownData(sxssfSheet, excelData.getExcelDropDowns(), excelData.getExcelData().size());
        }
        //默认样式
        ExcelStyleUtils.setCellMainStyle(cellStyle, font, excelData.getFontSize());

        //写入小标题与数据
        Integer SIZE = excelData.getExcelData().size() < ExcelGlobalConstants.MAX_ROWSUM ? excelData.getExcelData().size() : ExcelGlobalConstants.MAX_ROWSUM;

        //设置列标题
        sxssfRow = sxssfSheet.createRow(jRow);
        CellStyle cellTitleStyle = sxssfWorkbook.createCellStyle();
        ExcelStyleUtils.setCellTitleStyle(cellTitleStyle, font, excelData.getIndexedColors());
        for (int i = 0; i < excelFields.size(); i++) {
            sxssfSheet.setColumnWidth(i, excelData.getCellWidth() * ExcelGlobalConstants.EXCEL_WIDTH_UNIT);

            Field field = excelFields.get(i);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            Cell cell = createCell(sxssfRow, i, excelField.value());
            cell.setCellStyle(cellTitleStyle);
        }
        jRow += 1;

        for (int i = 0; i < SIZE; i++) {
            Object excelObject = excelData.getExcelData().get(i);
            sxssfRow = sxssfSheet.createRow(jRow);
            for (int j = 0, headSize = excelFields.size(); j < headSize; j++) {
                Field field = excelFields.get(j);
                Object value = BeanUtils.getFieldValue(excelObject, field);
                Cell cell = null;
                if (ImageUtils.patternIsImg((String) value)) {
                    cell = createCell(sxssfRow, j, " ");
                    ImageUtils.drawPicture(sxssfWorkbook, sxssfDrawing, (String) value, j, jRow);
                } else {
                    cell = createCell(sxssfRow, j, (String) value);
                }
                cell.setCellStyle(cellStyle);
            }
            jRow++;
        }
    }

    /**
     * 写数据与流关闭
     * @param sxssfWorkbook
     * @param outputStream
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void writeAndColse(SXSSFWorkbook sxssfWorkbook, OutputStream outputStream) throws Exception {
        try {
            if (outputStream != null) {
                sxssfWorkbook.write(outputStream);
                sxssfWorkbook.dispose();
                outputStream.flush();
                outputStream.close();
            }
        } catch (Exception e) {
            log.info("===> Exception Message：Output stream is not empty !");
            e.getSuppressed();
        }
    }

    /**
     * 下拉列表
     * @param sheet
     * @param excelDropDowns
     * @param dataListSize
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDropDownData(SXSSFSheet sheet, List<ExcelDropDown> excelDropDowns, int dataListSize) {
        if (!CollectionUtils.isEmpty(excelDropDowns)) {
            for (int i = 0; i < excelDropDowns.size(); i++) {
                ExcelDropDown excelDropDown = excelDropDowns.get(i);
                setDropDownData(sheet, excelDropDown.getDropDownData(), 1, dataListSize < 100 ? 500 : dataListSize,
                        excelDropDown.getColumnNum(), excelDropDown.getColumnNum());
            }
        }
    }

    /**
     * 下拉列表
     * @param xssfWsheet
     * @param dropDownData
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDropDownData(SXSSFSheet xssfWsheet, List<String> dropDownData, Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol) {
        DataValidationHelper helper = xssfWsheet.getDataValidationHelper();
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidationConstraint constraint = helper.createExplicitListConstraint(dropDownData.toArray(new String[dropDownData.size()]));
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        dataValidation.createErrorBox(ExcelGlobalConstants.DataValidationError1, ExcelGlobalConstants.DataValidationError2);
        //处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        xssfWsheet.addValidationData(dataValidation);
    }

    /**
     * response 响应
     * @param sxssfWorkbook
     * @param outputStream
     * @param excelData
     * @return void
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static void setExcelResponse(SXSSFWorkbook sxssfWorkbook, OutputStream outputStream, ExcelData excelData, HttpServletResponse httpServletResponse) {
        try {
            if (httpServletResponse != null) {
                httpServletResponse.setHeader("Charset", "UTF-8");
                httpServletResponse.setHeader("Content-Type", "application/force-download");
                httpServletResponse.setHeader("Content-Type", "application/vnd.ms-excel");
                httpServletResponse.setHeader("Content-disposition", "attachment; filename="
                        + URLEncoder.encode(StringUtils.isEmpty(excelData.getFileName()) ? excelData.getSheetName()
                        : excelData.getFileName(), "utf8") + ".xlsx");
                httpServletResponse.flushBuffer();
                outputStream = httpServletResponse.getOutputStream();
            }
            writeAndColse(sxssfWorkbook, outputStream);
        } catch (Exception e) {
            e.getSuppressed();
        }
    }

}
