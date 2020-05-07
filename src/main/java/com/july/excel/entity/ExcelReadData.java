package com.july.excel.entity;

import lombok.Data;

/**
 * excel读取数据信息
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-07 10:27
 **/
@Data
public class ExcelReadData {

    /**
     * sheet数
     */
    public Integer sheetNum;
    /**
     * 行数
     */
    public Integer rowNum;

}
