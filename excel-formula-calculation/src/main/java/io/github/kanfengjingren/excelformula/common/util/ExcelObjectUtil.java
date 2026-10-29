package io.github.kanfengjingren.excelformula.common.util;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 对象集合工具：按指定字段值对列表分组，用于按页签字段拆分工作簿行数据。
 */
public class ExcelObjectUtil {

    /**
     * 根据字段值对列表分组。
     *
     * @param list  源列表
     * @param field 分组依据字段
     * @param <K>   字段值类型（分组键）
     * @param <T>   元素类型
     * @return 字段值到元素列表的映射
     */
    public static <K, T> Map<K, List<T>> classifiedByField(List<T> list, Field field) {
        HashMap<K, List<T>> hashMap = new HashMap<>(16);
        list.forEach(entry -> {
            K key = ExcelReflectUtil.getValue(entry, field);
            hashMap.computeIfAbsent(key, k -> new ArrayList<>()).add(entry);
        });
        return hashMap;
    }

}
