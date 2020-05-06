package com.july.excel.utils;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 日期操作工具类
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-06 17:51
 **/
public class DateUtils {

    /**
     * 验证是否是日期
     * @param strDate
     * @param style
     * @return boolean
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static boolean verificationDate(String strDate, String style) {
        Date date = null;
        if (style == null) {
            style = "yyyy-MM-dd";
        }
        SimpleDateFormat formatter = new SimpleDateFormat(style);
        try {
            formatter.parse(strDate);
        } catch (Exception e) {
            return false;
        }
        return true;
    }

    /**
     * 字符串日期转为指定格式的字符串日期
     * @param strDate
     * @param style
     * @param expectDateFormatStr
     * @return java.lang.String
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static String strToDateFormat(String strDate, String style, String expectDateFormatStr) {
        Date date = null;
        if (style == null) {
            style = "yyyy-MM-dd";
        }
        //日期字符串转成date类型
        SimpleDateFormat formatter = new SimpleDateFormat(style);
        try {
            date = formatter.parse(strDate);
        } catch (Exception e) {
            return null;
        }
        //转成指定的日期格式
        SimpleDateFormat sdf = new SimpleDateFormat(expectDateFormatStr == null ? style : expectDateFormatStr);
        String str = sdf.format(date);
        return str;
    }

}
