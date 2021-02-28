package com.wsbxd.excel.formula.calculation.common.util;

import java.math.BigDecimal;
import java.text.NumberFormat;

/**
 * description: Excel 数学工具类
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelMathUtil {

    /**
     * 百分位数字格式转换
     */
    public final static NumberFormat NUMBER_PERCENT_FORMAT = NumberFormat.getPercentInstance();
    /**
     * 数字格式转换
     */
    public final static NumberFormat NUMBER__FORMAT = NumberFormat.getInstance();
    static {
        //默认转换不用千分位
        NUMBER_PERCENT_FORMAT.setGroupingUsed(false);
        NUMBER__FORMAT.setGroupingUsed(false);
    }
    private final static BigDecimal HUNDRED = new BigDecimal(100);

    /**
     * 数字转百分数
     *
     * @param num 数字
     * @return 百分数
     */
    public static String numToPercentage(double num) {
        return NUMBER_PERCENT_FORMAT.format(num);
    }

    /**
     * 百分数转数字
     *
     * @param percentage 百分数
     * @return 数字
     */
    public static double percentageToNum(String percentage) {
        return new BigDecimal(percentage.replace("%", "")).divide(HUNDRED).doubleValue();
    }

}
