package com.july.excel.utils;

import java.text.DecimalFormat;

/**
 * 数组操作类
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-07 16:32
 **/
public class NumberUtils {

    /**
     * 数字处理
     * @param decimalFormatThreadLocal
     * @param numeralFormat
     * @return java.text.DecimalFormat
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static DecimalFormat getDecimalFormat(ThreadLocal<DecimalFormat> decimalFormatThreadLocal, String numeralFormat) {
        DecimalFormat format = decimalFormatThreadLocal.get();
        if (format == null) {
            //默认数字格式： "#.######" 六位小数点
            format = new DecimalFormat(numeralFormat);
            decimalFormatThreadLocal.set(format);
        }
        return format;
    }

}
