package com.wsbxd.excel.formula.calculation.common.util;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;

/**
 * description: Excel 反射工具类
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelReflectUtil {

    /**
     * 使用反射设置变量值
     *
     * @param target    被调用对象
     * @param fieldName 被调用对象的字段，一般是成员变量或静态变量，不可是常量！
     * @param value     值
     * @param <T>       value类型，泛型
     */
    public static <T> void setValue(Object target, String fieldName, T value) {
        try {
            Class c = target.getClass();
            Field f = c.getDeclaredField(fieldName);
            if (!Modifier.isFinal(f.getModifiers())) {
                f.setAccessible(true);
                f.set(target, value);
            }
        } catch (NoSuchFieldException | IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    /**
     * 使用反射设置变量值
     *
     * @param target    被调用对象
     * @param fieldName 被调用对象的字段，一般是成员变量或静态变量，不可是常量！
     * @param value     值
     * @param <T>       value类型，泛型
     */
    public static <T> void setValue(Object target, Class<?> clazz, String fieldName, T value) {
        try {
            Field f = clazz.getDeclaredField(fieldName);
            if (!Modifier.isFinal(f.getModifiers())) {
                f.setAccessible(true);
                f.set(target, value);
            }
        } catch (NoSuchFieldException | IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    /**
     * 使用反射获取变量值
     *
     * @param target    被调用对象
     * @param fieldName 被调用对象的字段，一般是成员变量或静态变量，不可以是常量
     * @param <T>       返回类型，泛型
     * @return 值
     */
    public static <T> T getValue(Object target, String fieldName) {
        return getValue(target, target.getClass(), fieldName);
    }

    /**
     * 使用反射获取变量值
     *
     * @param target    被调用对象
     * @param fieldName 被调用对象的字段，一般是成员变量或静态变量，不可以是常量
     * @param <T>       返回类型，泛型
     * @return 值
     */
    public static <T> T getValue(Object target, Class<?> clazz, String fieldName) {
        T value = null;
        try {
            value = getValue(target, clazz.getDeclaredField(fieldName));
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        }
        return value;
    }

    /**
     * 使用反射获取变量值
     *
     * @param target 被调用对象
     * @param field  被调用对象的字段，一般是成员变量或静态变量，不可以是常量
     * @param <T>    返回类型，泛型
     * @return 值
     */
    public static <T> T getValue(Object target, Field field) {
        T value = null;
        try {
            field.setAccessible(true);
            value = (T) field.get(target);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return value;
    }

    /**
     * 使用反射根据字段名称获取字段
     *
     * @param target    被调用对象
     * @param fieldName 被调用对象的字段，一般是成员变量或静态变量，不可以是常量
     * @return 字段
     */
    public static Field getField(Object target, String fieldName) {
        Field field = null;
        try {
            field = target.getClass().getField(fieldName);
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        }
        return field;
    }

}
