package com.july.excel.property;

import lombok.Data;

/**
 * excel合并单元格
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-08 09:25
 **/
@Data
public class ExcelRegion {

    /**
     * 起始行号
     */
    public Integer startRowNum;
    /**
     * 终止行号
     */
    public Integer endRowNum;
    /**
     * 起始列号
     */
    public Integer startColumnNum;
    /**
     * 终止列号
     */
    public Integer endColumnNum;

}
