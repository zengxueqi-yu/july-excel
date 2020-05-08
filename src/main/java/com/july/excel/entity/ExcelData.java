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
     * sheet名称(多个时，逗号分开)
     */
    public String sheetName = "sheet1";
    /**
     * 每个表格的大标题
     */
    public String labelName;
    /**
     * 自定义：单元格合并[{1,1,2,5}]
     */
    public List<ExcelRegion> excelRegions;
    /**
     * 自定义：对每个单元格自定义下拉列表
     */
    public List<ExcelDropDown> excelDropDowns;
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
     * excel导入数据开始行数(多单元从第几行开始获取数据，默认从第二行开始获取（可为空，如 [{sheeNum=1,rowNum=3}]; 第一个表格从第三行开始获取）)
     */
    public Integer exportStartRowNum = 0;
    /**
     * 背景颜色
     */
    public IndexedColors indexedColors = IndexedColors.SKY_BLUE;
    /**
     * 从哪个sheet的，多少行开始读取数据
     */
    public List<ExcelReadData> excelReadDataList;

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
