package com.wsbxd.excel.formula.calculation.common.util;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * description: Excel 对象工具类
 *
 * @author chenhaoxuan
 * @date 2019/9/11
 */
public class ExcelObjectUtil {

    /**
     * 根据fieldName分类list
     *
     * @param list      list
     * @param fieldName 字段名
     * @param <K>       <K>
     * @param <T>       <T>
     * @return Map<fieldName, List < T>>
     */
    public static <K, T> Map<K, List<T>> classifiedByField(List<T> list, String fieldName) {
        return classifiedByField(list, ExcelReflectUtil.getField(list.get(0).getClass(), fieldName));
    }

    /**
     * 根据fieldName分类list
     *
     * @param list  list
     * @param field 字段
     * @param <K>   <K>
     * @param <T>   <T>
     * @return Map<fieldName, List < T>>
     */
    public static <K, T> Map<K, List<T>> classifiedByField(List<T> list, Field field) {
        HashMap<K, List<T>> hashMap = new HashMap<>(16);
        list.forEach(entry -> {
            K sheetName = ExcelReflectUtil.getValue(entry, field);
            if (hashMap.containsKey(sheetName)) {
                hashMap.get(sheetName).add(entry);
            } else {
                List<T> valueList = new ArrayList<>();
                valueList.add(entry);
                hashMap.put(sheetName, valueList);
            }
        });
        return hashMap;
    }

}
