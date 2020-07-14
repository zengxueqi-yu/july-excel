package com.july.excel.property;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * excel读取数据信息
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-07 10:27
 **/
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
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
