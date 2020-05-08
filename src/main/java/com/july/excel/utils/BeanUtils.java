package com.july.excel.utils;

import com.july.excel.entity.ExcelField;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 类操作工具
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-07 14:02
 **/
public class BeanUtils {

    /**
     * 获取需要导出的excel字段信息
     * @param excelClass
     * @param ignores
     * @return java.util.List<java.lang.reflect.Field>
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static List<Field> getExcelFields(Class<?> excelClass, String[] ignores) {
        Field[] declaredFields = excelClass.getDeclaredFields();
        List<Field> fieldList = new ArrayList<>(Arrays.asList(declaredFields));
        Class<?> superclass = excelClass.getSuperclass();
        if (superclass != Object.class) {
            fieldList.addAll(Arrays.asList(superclass.getDeclaredFields()));
        }
        return fieldList.stream()
                .filter(e -> e.isAnnotationPresent(ExcelField.class))
                .filter(e -> ParamUtils.noContains(ignores, e.getAnnotation(ExcelField.class).value()))
                .sorted(Comparator.comparing(e -> e.getAnnotation(ExcelField.class).sort()))
                .collect(Collectors.toList());
    }

    /**
     * 获取字段的值
     * @param object
     * @param field
     * @return java.lang.Object
     * @author zengxueqi
     * @since 2020/5/7
     */
    public static Object getFieldValue(Object object, Field field) {
        try {
            field.setAccessible(true);
            return field.get(object);
        } catch (IllegalAccessException e) {
            return null;
        }
    }

    /**
     * 给字段赋值
     * @param o
     * @param field
     * @param value
     * @return void
     * @author zengxueqi
     * @since 2020/5/8
     */
    public static void setFieldValue(Object o, Field field, Object value) {
        try {
            field.setAccessible(true);
            field.set(o, value);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }

}
