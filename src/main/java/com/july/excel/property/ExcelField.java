package com.july.excel.property;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel字段标记
 * @author cengxueqi
 * @since 2020/4/15
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {

    /**
     * excel列标题
     * @param
     * @return java.lang.String
     * @author zengxueqi
     * @since 2020/5/7
     */
    String value();

    int sort() default 99;

}
