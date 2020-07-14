package com.july.excel.property;

import lombok.Data;

import java.util.List;

/**
 * excel下拉框
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-08 09:07
 **/
@Data
public class ExcelDropDown {

    /**
     * 列数
     */
    private Integer columnNum;
    /**
     * 下拉框数据
     */
    private List<String> dropDownData;

}
