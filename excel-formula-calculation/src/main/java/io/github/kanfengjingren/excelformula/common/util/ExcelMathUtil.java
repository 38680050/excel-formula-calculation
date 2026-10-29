package io.github.kanfengjingren.excelformula.common.util;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.util.Locale;

/**
 * 数字与百分数互转工具。
 * <p>
 * 格式化器固定使用根区域（Locale.ROOT）且关闭千分位，保证不同系统环境下输出相同。
 */
public class ExcelMathUtil {

    /**
     * 百分数格式化器（根区域，不使用千分位）
     */
    public final static NumberFormat NUMBER_PERCENT_FORMAT = NumberFormat.getPercentInstance(Locale.ROOT);

    /**
     * 普通数字格式化器（根区域，不使用千分位）
     */
    public final static NumberFormat NUMBER__FORMAT = NumberFormat.getInstance(Locale.ROOT);

    static {
        NUMBER_PERCENT_FORMAT.setGroupingUsed(false);
        NUMBER__FORMAT.setGroupingUsed(false);
    }

    /** 百分数换算基数 */
    private final static BigDecimal HUNDRED = new BigDecimal(100);

    /**
     * 将小数转为百分数字符串。
     *
     * @param num 小数形式的数值
     * @return 百分数字符串
     */
    public static String numToPercentage(double num) {
        return NUMBER_PERCENT_FORMAT.format(num);
    }

    /**
     * 将百分数字符串转为小数。
     *
     * @param percentage 含百分号或不含百分号的字符串
     * @return 小数数值
     */
    @SuppressWarnings("BigDecimalMethodWithoutRoundingCalled")
    public static double percentageToNum(String percentage) {
        return new BigDecimal(percentage.replace("%", "")).divide(HUNDRED).doubleValue();
    }

}
