package io.github.kanfengjingren.excelformula.common.function;

import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.interfaces.IFunction;
import io.github.kanfengjingren.excelformula.common.util.ExcelMathUtil;
import io.github.kanfengjingren.excelformula.common.util.ExcelStrUtil;

import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * 内置公式函数实现。
 * <p>
 * 方法名即函数名，由函数注册表按名称反射调用；参数列表为已展开后的字符串值。
 */
public class DefaultFunctionImpl implements IFunction {

    private static final MathContext MATH_CONTEXT = new MathContext(16, RoundingMode.HALF_UP);

    // ---------- 逻辑 ----------

    /**
     * 条件取值：首参为 TRUE 或非零非布尔字符串时返回第二参（空白则 "0"）；
     * 否则返回第三参（缺省为 FALSE，空白则 "0"）。
     */
    public String IF(List<String> valueList) {
        String flag = valueList.get(0);
        if (isTruthy(flag)) {
            return ExcelStrUtil.isNotBlank(valueList.get(1)) ? valueList.get(1) : "0";
        }
        return valueList.size() > 2
                ? (ExcelStrUtil.isNotBlank(valueList.get(2)) ? valueList.get(2) : "0")
                : ExcelConstant.BIG_CASE_FALSE;
    }

    /**
     * 全部参数为真则 TRUE，否则 FALSE；无参为 TRUE。
     */
    public String AND(List<String> valueList) {
        for (String value : valueList) {
            if (!isTruthy(value)) {
                return ExcelConstant.BIG_CASE_FALSE;
            }
        }
        return ExcelConstant.BIG_CASE_TRUE;
    }

    /**
     * 任一参数为真则 TRUE，否则 FALSE；无参为 FALSE。
     */
    public String OR(List<String> valueList) {
        for (String value : valueList) {
            if (isTruthy(value)) {
                return ExcelConstant.BIG_CASE_TRUE;
            }
        }
        return ExcelConstant.BIG_CASE_FALSE;
    }

    /**
     * 对首参取逻辑非。
     */
    public String NOT(List<String> valueList) {
        return isTruthy(valueList.isEmpty() ? null : valueList.get(0))
                ? ExcelConstant.BIG_CASE_FALSE
                : ExcelConstant.BIG_CASE_TRUE;
    }

    /**
     * 首参为错误字面量时返回第二参，否则返回首参。
     */
    public String IFERROR(List<String> valueList) {
        String value = valueList.get(0);
        if (ExcelConstant.EXCEL_ERROR_VALUE_LIST.contains(value)) {
            return valueList.size() > 1 ? nullToEmpty(valueList.get(1)) : ExcelConstant.NULL_STR;
        }
        return nullToEmpty(value);
    }

    /**
     * 首参为 #N/A 时返回第二参，否则返回首参。
     */
    public String IFNA(List<String> valueList) {
        String value = valueList.get(0);
        if (ExcelConstant.EXCEL_ERROR_NA.equals(value)) {
            return valueList.size() > 1 ? nullToEmpty(valueList.get(1)) : ExcelConstant.NULL_STR;
        }
        return nullToEmpty(value);
    }

    /**
     * 首参是否为任意错误字面量。
     */
    public String ISERROR(List<String> valueList) {
        String value = valueList.isEmpty() ? null : valueList.get(0);
        return ExcelConstant.EXCEL_ERROR_VALUE_LIST.contains(value)
                ? ExcelConstant.BIG_CASE_TRUE
                : ExcelConstant.BIG_CASE_FALSE;
    }

    /**
     * 首参是否空白（null 或空串，不含仅空格）。
     */
    public String ISBLANK(List<String> valueList) {
        String value = valueList.isEmpty() ? null : valueList.get(0);
        return ExcelStrUtil.isEmpty(value) ? ExcelConstant.BIG_CASE_TRUE : ExcelConstant.BIG_CASE_FALSE;
    }

    /** 无参，返回 TRUE。 */
    public String TRUE(List<String> valueList) {
        return ExcelConstant.BIG_CASE_TRUE;
    }

    /** 无参，返回 FALSE。 */
    public String FALSE(List<String> valueList) {
        return ExcelConstant.BIG_CASE_FALSE;
    }

    // ---------- 数学 / 统计 ----------

    /** 取绝对值。 */
    public String ABS(List<String> valueList) {
        return formatNumber(ExcelStrUtil.getBigDecimal(firstOrZero(valueList)).abs());
    }

    /** 求和。 */
    public String SUM(List<String> valueList) {
        return ExcelStrUtil.sum(valueList);
    }

    /** 连乘；无有效数字时返回 0。 */
    public String PRODUCT(List<String> valueList) {
        List<BigDecimal> numbers = numericValues(valueList);
        if (numbers.isEmpty()) {
            return ExcelConstant.ZERO_STR;
        }
        BigDecimal product = BigDecimal.ONE;
        for (BigDecimal number : numbers) {
            product = product.multiply(number, MATH_CONTEXT);
        }
        return formatNumber(product);
    }

    /** 算术平均；忽略空白与非数字；无数字时返回 #DIV/0!。 */
    public String AVERAGE(List<String> valueList) {
        List<BigDecimal> numbers = numericValues(valueList);
        if (numbers.isEmpty()) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_DIV0, "AVERAGE 无数字可平均");
        }
        BigDecimal sum = BigDecimal.ZERO;
        for (BigDecimal number : numbers) {
            sum = sum.add(number);
        }
        return formatNumber(sum.divide(BigDecimal.valueOf(numbers.size()), MATH_CONTEXT));
    }

    /** 计数：仅统计可解析为数字的参数。 */
    public String COUNT(List<String> valueList) {
        return Integer.toString(numericValues(valueList).size());
    }

    /** 计数：统计非空参数。 */
    public String COUNTA(List<String> valueList) {
        int count = 0;
        for (String value : valueList) {
            if (ExcelStrUtil.isNotEmpty(value)) {
                count++;
            }
        }
        return Integer.toString(count);
    }

    /** 最大值；全空时返回 0。 */
    public String MAX(List<String> valueList) {
        return ExcelStrUtil.toString(valueList.stream()
                .filter(Objects::nonNull)
                .map(ExcelStrUtil::getBigDecimal)
                .max(BigDecimal::compareTo)
                .orElse(BigDecimal.ZERO));
    }

    /** 最小值；全空时返回 0。 */
    public String MIN(List<String> valueList) {
        return ExcelStrUtil.toString(valueList.stream()
                .filter(Objects::nonNull)
                .map(ExcelStrUtil::getBigDecimal)
                .min(BigDecimal::compareTo)
                .orElse(BigDecimal.ZERO));
    }

    /** 四舍五入到指定小数位（HALF_UP）。 */
    public String ROUND(List<String> valueList) {
        return roundWithMode(valueList, RoundingMode.HALF_UP);
    }

    /** 远离 0 舍入到指定小数位。 */
    public String ROUNDUP(List<String> valueList) {
        return roundWithMode(valueList, RoundingMode.UP);
    }

    /** 朝向 0 舍入到指定小数位。 */
    public String ROUNDDOWN(List<String> valueList) {
        return roundWithMode(valueList, RoundingMode.DOWN);
    }

    /** 向负无穷方向取整。 */
    public String INT(List<String> valueList) {
        return formatNumber(ExcelStrUtil.getBigDecimal(firstOrZero(valueList)).setScale(0, RoundingMode.FLOOR));
    }

    /**
     * 取余，语义为 {@code n - d * FLOOR(n/d)}；除数为 0 时 #DIV/0!。
     */
    public String MOD(List<String> valueList) {
        BigDecimal n = ExcelStrUtil.getBigDecimal(valueList.get(0));
        BigDecimal d = ExcelStrUtil.getBigDecimal(valueList.get(1));
        if (d.compareTo(BigDecimal.ZERO) == 0) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_DIV0, "MOD 除数为 0");
        }
        BigDecimal quotient = n.divide(d, 0, RoundingMode.FLOOR);
        return formatNumber(n.subtract(d.multiply(quotient)));
    }

    /** 幂运算：{@code base^exponent}。 */
    public String POWER(List<String> valueList) {
        BigDecimal base = ExcelStrUtil.getBigDecimal(valueList.get(0));
        BigDecimal exponent = ExcelStrUtil.getBigDecimal(valueList.get(1));
        try {
            if (exponent.stripTrailingZeros().scale() <= 0) {
                int exp = exponent.intValueExact();
                if (exp >= 0 && exp <= 9999) {
                    return formatNumber(base.pow(exp, MATH_CONTEXT));
                }
                if (exp < 0 && exp >= -9999 && base.compareTo(BigDecimal.ZERO) != 0) {
                    return formatNumber(BigDecimal.ONE.divide(base.pow(-exp, MATH_CONTEXT), MATH_CONTEXT));
                }
            }
        } catch (ArithmeticException ignored) {
            // 回退到 double
        }
        double result = Math.pow(base.doubleValue(), exponent.doubleValue());
        if (Double.isNaN(result) || Double.isInfinite(result)) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_NUM, "POWER 结果无效");
        }
        return formatNumber(BigDecimal.valueOf(result));
    }

    /** 平方根；负数返回 #NUM!。 */
    public String SQRT(List<String> valueList) {
        BigDecimal number = ExcelStrUtil.getBigDecimal(firstOrZero(valueList));
        if (number.compareTo(BigDecimal.ZERO) < 0) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_NUM, "SQRT 参数不能为负");
        }
        double result = Math.sqrt(number.doubleValue());
        return formatNumber(BigDecimal.valueOf(result));
    }

    /** 符号：正数 1、负数 -1、零 0。 */
    public String SIGN(List<String> valueList) {
        return Integer.toString(ExcelStrUtil.getBigDecimal(firstOrZero(valueList)).signum());
    }

    // ---------- 文本 ----------

    /** 文本长度（按 Java 字符数）。 */
    public String LEN(List<String> valueList) {
        return Integer.toString(nullToEmpty(firstOrEmpty(valueList)).length());
    }

    /** 左侧截取；第二参缺省为 1。 */
    public String LEFT(List<String> valueList) {
        String text = nullToEmpty(firstOrEmpty(valueList));
        int count = valueList.size() > 1 ? parseNonNegativeInt(valueList.get(1), "LEFT") : 1;
        if (count >= text.length()) {
            return text;
        }
        return text.substring(0, count);
    }

    /** 右侧截取；第二参缺省为 1。 */
    public String RIGHT(List<String> valueList) {
        String text = nullToEmpty(firstOrEmpty(valueList));
        int count = valueList.size() > 1 ? parseNonNegativeInt(valueList.get(1), "RIGHT") : 1;
        if (count >= text.length()) {
            return text;
        }
        return text.substring(text.length() - count);
    }

    /**
     * 中间截取：{@code MID(text, start, length)}，start 从 1 起算。
     */
    public String MID(List<String> valueList) {
        String text = nullToEmpty(valueList.get(0));
        int start = parsePositiveInt(valueList.get(1), "MID");
        int length = parseNonNegativeInt(valueList.get(2), "MID");
        if (start > text.length() || length == 0) {
            return ExcelConstant.NULL_STR;
        }
        int from = start - 1;
        int to = Math.min(text.length(), from + length);
        return text.substring(from, to);
    }

    /** 去除首尾空白。 */
    public String TRIM(List<String> valueList) {
        return nullToEmpty(firstOrEmpty(valueList)).trim();
    }

    /** 转大写。 */
    public String UPPER(List<String> valueList) {
        return nullToEmpty(firstOrEmpty(valueList)).toUpperCase(Locale.ROOT);
    }

    /** 转小写。 */
    public String LOWER(List<String> valueList) {
        return nullToEmpty(firstOrEmpty(valueList)).toLowerCase(Locale.ROOT);
    }

    /** 拼接全部参数文本。 */
    public String CONCATENATE(List<String> valueList) {
        return valueList.stream().map(this::nullToEmpty).collect(Collectors.joining());
    }

    /** 同 {@link #CONCATENATE}。 */
    public String CONCAT(List<String> valueList) {
        return CONCATENATE(valueList);
    }

    /** 将文本转为数字字符串；无法解析则 #VALUE!。 */
    public String VALUE(List<String> valueList) {
        return formatNumber(ExcelStrUtil.getBigDecimal(firstOrEmpty(valueList)));
    }

    // ---------- 内部工具 ----------

    private String roundWithMode(List<String> valueList, RoundingMode mode) {
        BigDecimal number = ExcelStrUtil.getBigDecimal(valueList.get(0));
        int scale = Integer.parseInt(valueList.get(1));
        return ExcelMathUtil.NUMBER__FORMAT.format(number.setScale(scale, mode));
    }

    private boolean isTruthy(String flag) {
        if (flag == null) {
            return false;
        }
        if (ExcelConstant.BIG_CASE_TRUE.equals(flag)) {
            return true;
        }
        if (ExcelConstant.BIG_CASE_FALSE.equals(flag) || ExcelConstant.ZERO_STR.equals(flag)) {
            return false;
        }
        return !ExcelConstant.BIG_TRUE_AND_FALSE.contains(flag);
    }

    private List<BigDecimal> numericValues(List<String> valueList) {
        List<BigDecimal> numbers = new ArrayList<>();
        for (String value : valueList) {
            if (ExcelStrUtil.isBlank(value)) {
                continue;
            }
            if (ExcelConstant.EXCEL_ERROR_VALUE_LIST.contains(value)) {
                throw new ExcelException(value, "非法数字: " + value);
            }
            try {
                numbers.add(new BigDecimal(value));
            } catch (NumberFormatException ignored) {
                // 统计类函数跳过非数字文本
            }
        }
        return numbers;
    }

    private String formatNumber(BigDecimal number) {
        return number.stripTrailingZeros().toPlainString();
    }

    private String firstOrZero(List<String> valueList) {
        return valueList == null || valueList.isEmpty() ? ExcelConstant.ZERO_STR : valueList.get(0);
    }

    private String firstOrEmpty(List<String> valueList) {
        return valueList == null || valueList.isEmpty() ? ExcelConstant.NULL_STR : valueList.get(0);
    }

    private String nullToEmpty(String value) {
        return value == null ? ExcelConstant.NULL_STR : value;
    }

    private int parseNonNegativeInt(String raw, String functionName) {
        try {
            int value = new BigDecimal(raw).setScale(0, RoundingMode.DOWN).intValueExact();
            if (value < 0) {
                throw new ExcelException(ExcelConstant.EXCEL_ERROR_VALUE, functionName + " 长度不能为负");
            }
            return value;
        } catch (NumberFormatException | ArithmeticException e) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_VALUE, functionName + " 参数无效: " + raw, e);
        }
    }

    private int parsePositiveInt(String raw, String functionName) {
        int value = parseNonNegativeInt(raw, functionName);
        if (value < 1) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_VALUE, functionName + " 起始位置从 1 开始");
        }
        return value;
    }
}
