package com.wsbxd.excel.formula.calculation.common.interfaces;

import java.util.List;

/**
 * description: Excel 公式
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public interface IFunction {

    /**
     * 根据对指定条件的逻辑判断的真假结果，返回相对应的内容
     * =IF(Logical,Value_if_true,Value_if_false)
     *
     * @param valueList 值集合
     * @return 结果
     */
    String IF(List<String> valueList);

    /**
     * 求出相应数字的绝对值
     * ABS(number)
     *
     * @param valueList 值集合
     * @return 结果
     */
    String ABS(List<String> valueList);

    /**
     * 计算所有参数数值的和
     * SUM（Number1,Number2……）
     * SUM（Number1:Number2）
     *
     * @param valueList 值集合
     * @return 结果
     */
    String SUM(List<String> valueList);

    /**
     * 获取传入值中最大的值
     *
     * @param valueList 值集合
     * @return 结果
     */
    String MAX(List<String> valueList);

    /**
     * 获取传入值中最小的值
     *
     * @param valueList 值集合
     * @return 结果
     */
    String MIN(List<String> valueList);

    String ROUND(List<String> valueList);
}
