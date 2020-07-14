package com.july.excel.utils;

import com.july.excel.property.ExcelRegion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.lang.reflect.Field;
import java.util.*;

import static org.apache.poi.ss.util.CellUtil.createCell;

/**
 * excel样式设置工具类
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-07 16:25
 **/
public class ExcelStyleUtils {

    /**
     * 大标题样式
     * @param wb
     * @param cell
     * @param sxssfRow
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setLabelStyles(SXSSFWorkbook wb, Cell cell, SXSSFRow sxssfRow) {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        sxssfRow.setHeight((short) (399 * 2));
        XSSFFont font = (XSSFFont) wb.createFont();
        font.setFontName("宋体");
        font.setFontHeight(16);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置默认样式
     * @param cellStyle
     * @param font
     * @param fontSize
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setCellMainStyle(CellStyle cellStyle, XSSFFont font, Integer fontSize) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        font.setFontName("宋体");
        font.setFontHeight(fontSize);
        cellStyle.setFont(font);
        setBorder(cellStyle, true);
    }

    /**
     * 设置默认样式
     * @param cellStyle
     * @param font
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setCellTitleStyle(CellStyle cellStyle, XSSFFont font, Short indexedColors) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillForegroundColor(indexedColors);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        font.setFontName("宋体");
        font.setBold(true);
        cellStyle.setFont(font);
        setBorder(cellStyle, true);
    }

    /**
     * excel合并单元格
     * @param sheet
     * @param excelRegions
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setMergedRegion(SXSSFSheet sheet, List<ExcelRegion> excelRegions) {
        excelRegions.stream().forEach(excelRegion -> setMergedRegion(sheet, excelRegion.getStartRowNum(),
                excelRegion.getEndRowNum(), excelRegion.getStartColumnNum(), excelRegion.getEndColumnNum()));
    }

    /**
     * 合并单元格
     * @param sheet
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setMergedRegion(SXSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /**
     * 设置边框
     * @param cellStyle
     * @param isBorder
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setBorder(CellStyle cellStyle, Boolean isBorder) {
        if (isBorder) {
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
        } else {
            //添加白色背景，统一设置边框后但不能选择性去掉，只能通过背景覆盖达到效果。
            cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setTopBorderColor(IndexedColors.WHITE.getIndex());
        }
    }

    /**
     * 自定义：大标题
     * @param jRow
     * @param wb
     * @param labelName
     * @param sxssfRow
     * @param sxssfSheet
     * @param excelField
     * @return int
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static int setLabelName(Integer jRow, SXSSFWorkbook wb, String labelName, SXSSFRow sxssfRow, SXSSFSheet sxssfSheet, List<Field> excelField) {
        if (labelName != null) {
            sxssfRow = sxssfSheet.createRow(0);
            Cell cell = createCell(sxssfRow, 0, labelName);
            setMergedRegion(sxssfSheet, 0, 0, 0, excelField.size() - 1);
            setLabelStyles(wb, cell, sxssfRow);
            jRow = 1;
        }
        return jRow;
    }

}
