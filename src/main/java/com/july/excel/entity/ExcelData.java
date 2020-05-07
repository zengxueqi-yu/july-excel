package com.july.excel.entity;

import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.HashMap;
import java.util.List;

/**
 * excel数据
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-06 19:08
 **/
@Data
public class ExcelData {

    /**
     * 导出excel数据
     */
    public List<?> excelData;
    /**
     * sheet名称必填
     */
    public String sheetName;
    /**
     * 每个表格的大标题
     */
    public String labelName;
    /**
     * 自定义：单元格合并
     */
    public HashMap regionMap;
    /**
     * 自定义：对每个单元格自定义列宽
     */
    public HashMap mapColumnWidth;
    /**
     * 自定义：每一个单元格样式
     */
    public HashMap styles;
    /**
     * 自定义：固定表头
     */
    public HashMap paneMap;
    /**
     * 自定义：某一行样式
     */
    public HashMap rowStyles;
    /**
     * 自定义：某一列样式
     */
    public HashMap columnStyles;
    /**
     * 自定义：对每个单元格自定义下拉列表
     */
    public HashMap dropDownMap;
    /**
     * 文件名称
     */
    public String fileName;
    /**
     * 导出本地路径
     */
    public String filePath;
    /**
     * 导出数字格式化：默认是保留六位小数点
     */
    public String numeralFormat;
    /**
     * 导出日期格式化：默认是"yyyy-MM-dd"格式
     */
    public String dateFormatStr;
    /**
     * 期望转换后的日期格式：默认是 dateFormatStr
     */
    public String expectDateFormatStr;
    /**
     * 默认列宽大小：默认16
     */
    public Integer cellWidth = 20 * 256;
    /**
     * 默认字体大小：默认12号字体
     */
    public Integer fontSize;
    /**
     * 需要忽略生成excel的字段
     */
    public String[] ignores;
    /**
     * 背景颜色
     */
    public IndexedColors indexedColors = IndexedColors.SKY_BLUE;

    public Integer getFontSize() {
        if (fontSize == null) {
            fontSize = 10;
        }
        return fontSize;
    }

    public void setFontSize(Integer fontSize) {
        this.fontSize = fontSize;
    }

    public void setDateFormatStr(String dateFormatStr) {
        if (dateFormatStr == null) {
            dateFormatStr = "yyyy-MM-dd";
        }
        this.dateFormatStr = dateFormatStr;
    }

    public String getDateFormatStr() {
        if (dateFormatStr == null) {
            dateFormatStr = "yyyy-MM-dd";
        }
        return dateFormatStr;
    }

    public String getExpectDateFormatStr() {
        if (expectDateFormatStr == null) {
            expectDateFormatStr = dateFormatStr;
        }
        return expectDateFormatStr;
    }

    public void setExpectDateFormatStr(String expectDateFormatStr) {
        if (expectDateFormatStr == null) {
            expectDateFormatStr = dateFormatStr;
        }
        this.expectDateFormatStr = expectDateFormatStr;
    }

    public void setNumeralFormat(String numeralFormat) {
        if (numeralFormat == null) {
            numeralFormat = "#.######";
        }
        this.numeralFormat = numeralFormat;
    }

    public String getNumeralFormat() {
        if (numeralFormat == null) {
            numeralFormat = "#.######";
        }
        return numeralFormat;
    }

}
