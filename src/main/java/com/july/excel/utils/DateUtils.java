package com.july.excel.utils;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

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
     * @param expectDateFormatStr
     * @return java.lang.String
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static String strToDateFormat(String strDate, String expectDateFormatStr) {
        Date date = null;
        //日期字符串转成date类型
        SimpleDateFormat formatter = new SimpleDateFormat(expectDateFormatStr);
        try {
            date = formatter.parse(strDate);
        } catch (Exception e) {
            return null;
        }
        //转成指定的日期格式
        SimpleDateFormat sdf = new SimpleDateFormat(expectDateFormatStr);
        String str = sdf.format(date);
        return str;
    }

    public static SimpleDateFormat getDateFormat(ThreadLocal<SimpleDateFormat> simpleDateFormatThreadLocal, String expectDateFormatStr) {
        SimpleDateFormat format = simpleDateFormatThreadLocal.get();
        if (format == null) {
            //默认格式日期： "yyyy-MM-dd"
            format = new SimpleDateFormat(expectDateFormatStr, Locale.getDefault());
            simpleDateFormatThreadLocal.set(format);
        }
        return format;
    }

    public static String getDateFormatStr() {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
        return simpleDateFormat.format(new Date());
    }

}
