package com.july.excel.utils;

import com.july.excel.bean.DynamicBean;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Set;

/**
 * 类操作工具类
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-09 09:51
 **/
public class ClassUtils {

    /**
     * 类的属性添加
     * @param object    旧的对象带值
     * @param addMap    动态需要添加的属性和属性类型
     * @param addValMap 动态需要添加的属性和属性值
     * @return java.lang.Object 新的对象
     * @author zengxueqi
     * @since 2020/5/9
     */
    public Object dynamicClass(Object object, HashMap addMap, HashMap addValMap) throws Exception {
        HashMap returnMap = new HashMap();
        HashMap typeMap = new HashMap();

        Class<?> type = object.getClass();
        BeanInfo beanInfo = Introspector.getBeanInfo(type);
        PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
        for (int i = 0; i < propertyDescriptors.length; i++) {
            PropertyDescriptor descriptor = propertyDescriptors[i];
            String propertyName = descriptor.getName();
            if (!propertyName.equals("class")) {
                Method readMethod = descriptor.getReadMethod();
                Object result = readMethod.invoke(object);
                //可以判断为 NULL不赋值
                returnMap.put(propertyName, result);
                typeMap.put(propertyName, descriptor.getPropertyType());
            }
        }

        returnMap.putAll(addValMap);
        typeMap.putAll(addMap);
        //map转换成实体对象
        DynamicBean bean = new DynamicBean(typeMap);
        //赋值
        Set keys = typeMap.keySet();
        for (Iterator it = keys.iterator(); it.hasNext(); ) {
            String key = (String) it.next();
            bean.setValue(key, returnMap.get(key));
        }
        Object obj = bean.getObject();
        return obj;
    }

}
