package com.july.excel.utils;

import com.july.excel.constant.ExcelGlobalConstants;
import com.july.excel.entity.ExcelRegion;
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
     * 所有自定义单元格样式
     * @param cell
     * @param wb
     * @param sxssfRow
     * @param styles
     * @param cellIndex
     * @param rowIndex
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setExcelStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, List<List<Object[]>> styles, int cellIndex, int rowIndex) {
        if (styles != null) {
            for (int z = 0; z < styles.size(); z++) {
                List<Object[]> stylesList = styles.get(z);
                if (stylesList != null) {
                    //样式
                    Boolean[] bool = (Boolean[]) stylesList.get(0);
                    //颜色和字体
                    Integer fontColor = null;
                    Integer fontSize = null;
                    Integer height = null;
                    //当有设置颜色值 、字体大小、行高才获取值
                    if (stylesList.size() >= 2) {
                        int leng = stylesList.get(1).length;
                        if (leng >= 1) {
                            fontColor = (Integer) stylesList.get(1)[0];
                        }
                        if (leng >= 2) {
                            fontSize = (Integer) stylesList.get(1)[1];
                        }
                        if (leng >= 3) {
                            height = (Integer) stylesList.get(1)[2];
                        }
                    }
                    //第几行第几列
                    for (int m = 2; m < stylesList.size(); m++) {
                        Integer[] str = (Integer[]) stylesList.get(m);
                        if (cellIndex + 1 == (str[1]) && rowIndex + 1 == (str[0])) {
                            setExcelStyles(cell, wb, sxssfRow, fontSize, Boolean.valueOf(bool[3]), Boolean.valueOf(bool[0]), Boolean.valueOf(bool[4]), Boolean.valueOf(bool[2]), Boolean.valueOf(bool[1]), fontColor, height);
                        }
                    }
                }
            }
        }
    }

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
    public static void setCellTitleStyle(CellStyle cellStyle, XSSFFont font, IndexedColors indexedColors) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillForegroundColor(indexedColors.getIndex());
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
     * 设置边框颜色
     * @param cellStyle
     * @param isBorder
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setBorderColor(CellStyle cellStyle, Boolean isBorder) {
        if (isBorder) {
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

    /**
     * 设置excel样式
     * @param cell         Cell对象
     * @param wb           SXSSFWorkbook对象
     * @param sxssfRow
     * @param fontSize     字体大小
     * @param bold         是否加粗
     * @param center       是否左右上下居中
     * @param isBorder     是否忽略边框
     * @param leftBoolean  左对齐
     * @param rightBoolean 右对齐
     * @param fontColor    字体颜色
     * @param height       行高
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setExcelStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, Integer fontSize, Boolean bold, Boolean center, Boolean isBorder, Boolean leftBoolean,
                                      Boolean rightBoolean, Integer fontColor, Integer height) {
        CellStyle cellStyle = cell.getRow().getSheet().getWorkbook().createCellStyle();
        //保证了既可以新建一个CellStyle，又可以不丢失原来的CellStyle 的样式
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        //左右居中、上下居中
        if (center != null && center) {
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }
        //右对齐
        if (rightBoolean != null && rightBoolean) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        }
        //左对齐
        if (leftBoolean != null && leftBoolean) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
        }
        //是否忽略边框
        if (isBorder != null && isBorder) {
            ExcelStyleUtils.setBorderColor(cellStyle, isBorder);
        }
        //设置单元格字体样式
        XSSFFont font = (XSSFFont) wb.createFont();
        if (bold != null && bold) {
            font.setBold(bold);
        }
        //行高
        if (height != null) {
            sxssfRow.setHeight((short) (height * 2));
        }
        font.setFontName("宋体");
        font.setFontHeight(fontSize == null ? 12 : fontSize);
        cellStyle.setFont(font);
        //点击可查看颜色对应的值： BLACK(8), WHITE(9), RED(10),
        font.setColor(IndexedColors.fromInt(fontColor == null ? 8 : fontColor).index);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置excel行的样式
     * @param cell
     * @param wb
     * @param sxssfRow
     * @param rowstyleList
     * @param rowIndex
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setExcelRowStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, List<Object[]> rowstyleList, int rowIndex) {
        if (rowstyleList != null && rowstyleList.size() > 0) {
            Integer[] rowstyle = (Integer[]) rowstyleList.get(1);
            for (int i = 0; i < rowstyle.length; i++) {
                if (rowIndex == rowstyle[i]) {
                    Boolean[] bool = (Boolean[]) rowstyleList.get(0);
                    Integer fontColor = null;
                    Integer fontSize = null;
                    Integer height = null;
                    //当有设置颜色值 、字体大小、行高才获取值
                    if (rowstyleList.size() >= 3) {
                        int leng = rowstyleList.get(2).length;
                        if (leng >= 1) {
                            fontColor = (Integer) rowstyleList.get(2)[0];
                        }
                        if (leng >= 2) {
                            fontSize = (Integer) rowstyleList.get(2)[1];
                        }
                        if (leng >= 3) {
                            height = (Integer) rowstyleList.get(2)[2];
                        }
                    }
                    ExcelStyleUtils.setExcelStyles(cell, wb, sxssfRow, fontSize, Boolean.valueOf(bool[3]), Boolean.valueOf(bool[0]), Boolean.valueOf(bool[4]), Boolean.valueOf(bool[2]), Boolean.valueOf(bool[1]), fontColor, height);
                }
            }
        }
    }

}
