package com.july.excel.utils;

/**
 * 属性工具类
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-07 14:04
 **/
public class ParamUtils {

    public static boolean noContains(String[] arr, String val) {
        if (arr == null || arr.length == 0) {
            return true;
        }
        for (String o : arr) {
            if (o.equals(val)) {
                return false;
            }
        }
        return true;
    }

}
