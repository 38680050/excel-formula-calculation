package io.github.kanfengjingren.excelformula.common.util;

import io.github.kanfengjingren.excelformula.common.exception.ExcelException;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;

/**
 * 反射读写工具：按声明字段读写对象属性；失败时抛出业务异常。
 */
public class ExcelReflectUtil {

    /**
     * 按字段名设置对象属性值（非 final 字段）。
     */
    public static <T> void setValue(Object target, String fieldName, T value) {
        setValue(target, target.getClass(), fieldName, value);
    }

    /**
     * 在指定类上按字段名设置属性值。
     */
    public static <T> void setValue(Object target, Class<?> clazz, String fieldName, T value) {
        try {
            Field f = clazz.getDeclaredField(fieldName);
            if (!Modifier.isFinal(f.getModifiers())) {
                f.setAccessible(true);
                f.set(target, value);
            }
        } catch (NoSuchFieldException | IllegalAccessException e) {
            throw new ExcelException("设置字段失败: " + fieldName + ", " + e.getMessage(), e);
        }
    }

    /**
     * 按字段名读取对象属性值。
     */
    public static <T> T getValue(Object target, String fieldName) {
        return getValue(target, target.getClass(), fieldName);
    }

    /**
     * 在指定类上按字段名读取属性值。
     */
    public static <T> T getValue(Object target, Class<?> clazz, String fieldName) {
        try {
            return getValue(target, clazz.getDeclaredField(fieldName));
        } catch (NoSuchFieldException e) {
            throw new ExcelException("读取字段失败: " + fieldName + ", " + e.getMessage(), e);
        }
    }

    /**
     * 按 Field 对象读取属性值。
     */
    @SuppressWarnings("unchecked")
    public static <T> T getValue(Object target, Field field) {
        try {
            field.setAccessible(true);
            return (T) field.get(target);
        } catch (IllegalAccessException e) {
            throw new ExcelException("读取字段失败: " + field.getName() + ", " + e.getMessage(), e);
        }
    }

    /**
     * 获取声明字段并设为可访问。
     */
    public static Field getDeclaredField(Class<?> clazz, String fieldName) {
        try {
            Field field = clazz.getDeclaredField(fieldName);
            field.setAccessible(true);
            return field;
        } catch (NoSuchFieldException e) {
            throw new ExcelException("获取字段失败: " + fieldName + ", " + e.getMessage(), e);
        }
    }
}
