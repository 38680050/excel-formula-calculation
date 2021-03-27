package com.wsbxd.excel.formula.calculation.common.function;

import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.interfaces.IFunction;
import com.wsbxd.excel.formula.calculation.common.util.ExcelMathUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.List;
import java.util.Objects;

/**
 * description: 默认 Excel 公式 实现
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class DefaultFunctionImpl implements IFunction {

    /**
     * 说明：
     * IF 函数是 Excel 中最常用的函数之一，它可以对值和期待值进行逻辑比较。
     * 因此 IF 语句可能有两个结果。 第一个结果是比较结果为 True，第二个结果是比较结果为 False。
     * 例如，=IF(C2=”Yes”,1,2) 表示 IF(C2 = Yes, 则返回 1, 否则返回 2)。
     * <p>
     * 语法：
     * 使用逻辑函数 IF 函数时，如果条件为真，该函数将返回一个值；如果条件为假，函数将返回另一个值。
     * IF(logical_test, value_if_true, [value_if_false])
     * <p>
     * 例如：
     * =IF(A2>B2,"超出预算","正常")
     * =IF(A2=B2,B4-A4,"")
     * logical_test     必需，要测试的条件。
     * value_if_true    必需，logical_test 的结果为 TRUE 时，您希望返回的值。
     * value_if_false   可选，logical_test 的结果为 FALSE 时，您希望返回的值。
     *
     * @param valueList 值集合
     * @return 结果
     */
    public String IF(List<String> valueList) {
        String flag = valueList.get(0);
        //true 或 != 0 都走 valueList.get(1)
        if (ExcelConstant.BIG_CASE_TRUE.equals(flag) || (!ExcelConstant.TRUE_AND_FALSE.contains(flag) && !ExcelConstant.ZERO.equals(flag))) {
            return ExcelStrUtil.isNotBlank(valueList.get(1)) ? valueList.get(1) : "0";
        } else {
            return valueList.size() > 2 ? ExcelStrUtil.isNotBlank(valueList.get(2)) ? valueList.get(2) : "0" : ExcelConstant.BIG_CASE_FALSE;
        }
    }

    /**
     * 说明：
     * 返回数字的绝对值。 一个数字的绝对值是该数字不带其符号的形式。
     * <p>
     * 语法：
     * ABS(number)
     * ABS 函数语法具有以下参数：
     * Number    必需。 需要计算其绝对值的实数。
     *
     * @param valueList 值集合
     * @return 结果
     */
    public String ABS(List<String> valueList) {
        return ExcelStrUtil.getBigDecimal(valueList.stream().findAny().orElse("0")).abs().toString();
    }

    /**
     * 说明：
     * SUM函数将添加值。 你可以将单个值、单元格引用或是区域相加，或者将三者的组合相加。
     * <p>
     * 语法：
     * SUM(number1,[number2],...)
     * number1      必需，要相加的第一个数字。 该数字可以是 4 之类的数字，B6 之类的单元格引用或 B2:B8 之类的单元格范围。
     * number2-255  可选，这是要相加的第二个数字。 可以按照这种方式最多指定 255 个数字。
     * <p>
     * 例如：
     * =SUM(A2:A10) 对单元格 A2：10 中的值进行加法。
     * =SUM(A2:A10, C2:C10) 对单元格 A2：10 以及单元格 C2：C10 中的值进行加法。
     *
     * @param valueList 值集合
     * @return 结果
     */
    public String SUM(List<String> valueList) {
        return ExcelStrUtil.sum(valueList);
    }

    /**
     * 说明：
     * 返回一组值中的最大值。
     * <p>
     * 语法：
     * MAX(number1, [number2], ...)
     * MAX 函数语法具有下列参数：
     * number1, number2, ...    Number1 是必需的，后续数字是可选的。 要从中查找最大值的 1 到 255 个数字。
     *
     * @param valueList 值集合
     * @return 结果
     */
    public String MAX(List<String> valueList) {
        return ExcelStrUtil.toString(valueList.stream().filter(Objects::nonNull).map(ExcelStrUtil::getBigDecimal).max(BigDecimal::compareTo).orElse(new BigDecimal("0")));
    }

    /**
     * 说明：
     * 返回一组值中的最小值。
     * <p>
     * 语法：
     * MIN(number1, [number2], ...)
     * MIN 函数语法具有下列参数：
     * number1, number2, ...    number1 是可选的，后续数字是可选的。 要从中查找最小值的 1 到 255 个数字。
     *
     * @param valueList 值集合
     * @return 结果
     */
    public String MIN(List<String> valueList) {
        return ExcelStrUtil.toString(valueList.stream().filter(Objects::nonNull).map(ExcelStrUtil::getBigDecimal).min(BigDecimal::compareTo).orElse(new BigDecimal("0")));
    }

    /**
     * 说明：
     * ROUND 函数将数字四舍五入到指定的位数。 例如，如果单元格 A1 包含 23.7825，而且您想要将此数值舍入到两个小数位数，可以使用以下公式：
     * =ROUND(A1, 2)
     * 此函数的结果为 23.78。
     * <p>
     * 语法：
     * ROUND(number, num_digits)
     * ROUND 函数语法具有下列参数：
     * number           必需，要四舍五入的数字。
     * num_digits       必需，要进行四舍五入运算的位数。
     *
     * @param valueList 值集合
     * @return 结果
     */
    public String ROUND(List<String> valueList) {
        return ExcelMathUtil.NUMBER__FORMAT.format(new BigDecimal(valueList.get(0)).setScale(Integer.parseInt(valueList.get(1)), RoundingMode.HALF_UP));
    }

}
