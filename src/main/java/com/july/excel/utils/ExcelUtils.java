package com.july.excel.utils;

import com.july.excel.constant.ExcelGlobalConstants;
import com.july.excel.entity.ExcelData;
import com.july.excel.entity.ExcelField;
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
     * 设置数据：有样式（行、列、单元格样式）
     * @param wb
     * @param sxssfRow
     * @param excelData
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDataList(SXSSFWorkbook wb, SXSSFRow sxssfRow, ExcelData excelData, List<Field> excelFields) {
        if (CollectionUtils.isEmpty(excelData.getExcelData())) {
            log.debug("===> Exception Message：Export data(type:List<List<String[]>>) cannot be empty!");
        }
        if (excelData.getSheetName() == null) {
            log.debug("===> Exception Message：Export sheet(type:String[]) name cannot be empty!");
        }
        int k = 0;
        SXSSFSheet sxssfSheet = wb.createSheet();
        sxssfSheet.setDefaultColumnWidth(excelData.getCellWidth());
        wb.setSheetName(k, excelData.getSheetName());
        CellStyle cellStyle = wb.createCellStyle();
        XSSFFont font = (XSSFFont) wb.createFont();
        SXSSFDrawing sxssfDrawing = sxssfSheet.createDrawingPatriarch();

        int jRow = 0;

        //自定义：大标题（看该方法说明）。
        jRow = ExcelStyleUtils.setLabelName(jRow, wb, excelData.getLabelName(), sxssfRow, sxssfSheet, excelFields);

        //自定义：每个表格固定表头（看该方法说明）。
        Integer pane = 1;
        if (!CollectionUtils.isEmpty(excelData.getPaneMap()) && excelData.getPaneMap().get(k + 1) != null) {
            pane = (Integer) excelData.getPaneMap().get(k + 1) + (excelData.getLabelName() != null ? 1 : 0);
            createFreezePane(sxssfSheet, pane);
        }
        //自定义：每个单元格自定义合并单元格：对每个单元格自定义合并单元格（看该方法说明）。
        if (!CollectionUtils.isEmpty(excelData.getRegionMap())) {
            ExcelStyleUtils.setMergedRegion(sxssfSheet, (ArrayList<Integer[]>) excelData.getRegionMap().get(k + 1));
        }
        //自定义：每个单元格自定义下拉列表：对每个单元格自定义下拉列表（看该方法说明）。
        if (!CollectionUtils.isEmpty(excelData.getDropDownMap())) {
            setDataValidation(sxssfSheet, (List<String[]>) excelData.getDropDownMap().get(k + 1), excelData.getExcelData().size());
        }
        //自定义：每个表格自定义列宽：对每个单元格自定义列宽（看该方法说明）。
        if (!CollectionUtils.isEmpty(excelData.getMapColumnWidth())) {
            ExcelStyleUtils.setColumnWidth(sxssfSheet, (HashMap) excelData.getMapColumnWidth().get(k + 1));
        }
        //默认样式。
        ExcelStyleUtils.setCellMainStyle(cellStyle, font, excelData.getFontSize());

        //设置列标题
        sxssfRow = sxssfSheet.createRow(jRow);
        for (int i = 0; i < excelFields.size(); i++) {
            Field field = excelFields.get(i);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            Cell cell = createCell(sxssfRow, i, excelField.value());
            cell.setCellStyle(cellStyle);
        }
        jRow += 1;

        //写入小标题与数据。
        Integer SIZE = excelData.getExcelData().size() < ExcelGlobalConstants.MAX_ROWSUM ? excelData.getExcelData().size() : ExcelGlobalConstants.MAX_ROWSUM;
        Integer MAXSYTLE = excelData.getExcelData().size() < ExcelGlobalConstants.MAX_ROWSTYLE ? excelData.getExcelData().size() : ExcelGlobalConstants.MAX_ROWSTYLE;
        for (int i = 0; i < SIZE; i++) {
            Object excelObject = excelData.getExcelData().get(i);
            sxssfRow = sxssfSheet.createRow(jRow);
            for (int j = 0, headSize = excelFields.size(); j < headSize; j++) {
                Field field = excelFields.get(j);
                Object value = BeanUtils.getFieldValue(excelObject, field);
                //样式过多会导致GC内存溢出！
                try {
                    Cell cell = null;
                    if (ImageUtils.patternIsImg((String) excelObject)) {
                        cell = createCell(sxssfRow, j, " ");
                        ImageUtils.drawPicture(wb, sxssfDrawing, (String) excelObject, j, jRow);
                    } else {
                        cell = createCell(sxssfRow, j, (String) excelObject);
                    }
                    cell.setCellStyle(cellStyle);

                    //自定义：每个表格每一列的样式（看该方法说明）。
                    if (excelData.getColumnStyles() != null && jRow >= pane && i <= MAXSYTLE) {
                        ExcelStyleUtils.setExcelRowStyles(cell, wb, sxssfRow, (List) excelData.getColumnStyles().get(k + 1), j);
                    }
                    //自定义：每个表格每一行的样式（看该方法说明）。
                    if (excelData.getRowStyles() != null && i <= MAXSYTLE) {
                        ExcelStyleUtils.setExcelRowStyles(cell, wb, sxssfRow, (List) excelData.getRowStyles().get(k + 1), jRow);
                    }
                    //自定义：每一个单元格样式（看该方法说明）。
                    if (excelData.getStyles() != null && i <= MAXSYTLE) {
                        ExcelStyleUtils.setExcelStyles(cell, wb, sxssfRow, (List<List<Object[]>>) excelData.getStyles().get(k + 1), j, i);
                    }
                } catch (Exception e) {
                    log.debug("===> Exception Message：The maximum number of cell styles was exceeded. You can define up to 4000 styles!");
                }
            }
            jRow++;
        }
    }

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
    public static void setDataListNoStyle(SXSSFWorkbook sxssfWorkbook, SXSSFRow sxssfRow, ExcelData excelData, List<Field> excelFields) {
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

        int jRow = 0;
        //自定义：大标题（看该方法说明）。
        jRow = ExcelStyleUtils.setLabelName(jRow, sxssfWorkbook, excelData.getLabelName(), sxssfRow, sxssfSheet, excelFields);

        //自定义：每个表格固定表头（看该方法说明）。
        Integer pane = 1;
        if (!CollectionUtils.isEmpty(excelData.getPaneMap()) && excelData.getPaneMap().get(k + 1) != null) {
            pane = (Integer) excelData.getPaneMap().get(k + 1) + (excelData.getLabelName() != null ? 1 : 0);
            createFreezePane(sxssfSheet, pane);
        }
        //自定义：每个单元格自定义合并单元格：对每个单元格自定义合并单元格（看该方法说明）。
        if (!CollectionUtils.isEmpty(excelData.getRegionMap())) {
            ExcelStyleUtils.setMergedRegion(sxssfSheet, (ArrayList<Integer[]>) excelData.getRegionMap().get(k + 1));
        }
        //自定义：每个单元格自定义下拉列表：对每个单元格自定义下拉列表（看该方法说明）。
        if (!CollectionUtils.isEmpty(excelData.getDropDownMap())) {
            setDataValidation(sxssfSheet, (List<String[]>) excelData.getDropDownMap().get(k + 1), excelData.getExcelData().size());
        }
        //自定义：每个表格自定义列宽：对每个单元格自定义列宽（看该方法说明）。
        if (!CollectionUtils.isEmpty(excelData.getMapColumnWidth())) {
            ExcelStyleUtils.setColumnWidth(sxssfSheet, (HashMap) excelData.getMapColumnWidth().get(k + 1));
        }
        //默认样式。
        ExcelStyleUtils.setCellMainStyle(cellStyle, font, excelData.getFontSize());

        //写入小标题与数据。
        Integer SIZE = excelData.getExcelData().size() < ExcelGlobalConstants.MAX_ROWSUM ? excelData.getExcelData().size() : ExcelGlobalConstants.MAX_ROWSUM;

        //设置列标题
        sxssfRow = sxssfSheet.createRow(jRow);
        CellStyle cellTitleStyle = sxssfWorkbook.createCellStyle();
        ExcelStyleUtils.setCellTitleStyle(cellTitleStyle, font, excelData.getIndexedColors());
        for (int i = 0; i < excelFields.size(); i++) {
            sxssfSheet.setColumnWidth(i, excelData.getCellWidth());

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
                Cell cell = createCell(sxssfRow, j, (String) value);
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
            System.out.println(" Andyczy ExcelUtils Exception Message：Output stream is not empty !");
            e.getSuppressed();
        }
    }

    /**
     * 锁定行（固定表头）
     * @param sxssfSheet
     * @param row
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void createFreezePane(SXSSFSheet sxssfSheet, Integer row) {
        if (row != null && row > 0) {
            sxssfSheet.createFreezePane(0, row, 0, 1);
        }
    }

    /**
     * 下拉列表
     * @param sheet
     * @param dropDownListData
     * @param dataListSize
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDataValidation(SXSSFSheet sheet, List<String[]> dropDownListData, int dataListSize) {
        if (dropDownListData.size() > 0) {
            for (int col = 0; col < dropDownListData.get(0).length; col++) {
                Integer colv = Integer.parseInt(dropDownListData.get(0)[col]);
                setDataValidation(sheet, dropDownListData.get(col + 1), 1, dataListSize < 100 ? 500 : dataListSize, colv, colv);
            }
        }
    }

    /**
     * 下拉列表
     * @param xssfWsheet
     * @param list
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDataValidation(SXSSFSheet xssfWsheet, String[] list, Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol) {
        DataValidationHelper helper = xssfWsheet.getDataValidationHelper();
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidationConstraint constraint = helper.createExplicitListConstraint(list);
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
