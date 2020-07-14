package com.july.excel.property;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.List;

/**
 * excel数据
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-06 19:08
 **/
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelData {

    /**
     * 导出excel数据
     */
    public List<?> excelData;
    /**
     * sheet名称(多个时，逗号分开)
     */
    @Builder.Default
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
    @Builder.Default
    public String numeralFormat = "#.######";
    /**
     * 期望转换后的日期格式：默认是yyyy-MM-dd
     */
    @Builder.Default
    public String expectDateFormatStr = "yyyy-MM-dd";
    /**
     * 默认列宽大小：默认16
     */
    @Builder.Default
    public Integer cellWidth = 20;
    /**
     * 默认字体大小：默认12号字体
     */
    @Builder.Default
    public Integer fontSize = 10;
    /**
     * 需要忽略生成excel的字段
     */
    public String[] ignores;
    /**
     * 从哪个sheet的，多少行读取标题数据
     */
    public List<ExcelReadData> excelTitleRowNum;
    /**
     * 背景颜色
     */
    @Builder.Default
    public Short indexedColors = IndexedColors.SKY_BLUE.getIndex();
    /**
     * 从哪个sheet的，多少行开始读取数据
     */
    public List<ExcelReadData> excelReadDataList;

}
