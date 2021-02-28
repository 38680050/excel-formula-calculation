package com.wsbxd.excel.formula.calculation.common.util;

/**
 * description: Excel 单元格类型转换工具类
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelCellTypeUtil {

    /**
     * 数字转百分数
     *
     * @param num 数字
     * @return 百分数
     */
    public static String numToPercentage(String num) {
        if (ExcelStrUtil.isBlank(num)) {
            return "0%";
        }
        return ExcelMathUtil.numToPercentage(Double.parseDouble(num));
    }

    /**
     * 百分数转数字
     *
     * @param percentage 百分数
     * @return 数字
     */
    public static String percentageToNum(String percentage) {
        if (ExcelStrUtil.isBlank(percentage)) {
            return "0";
        }
        return String.valueOf(ExcelMathUtil.percentageToNum(percentage));
    }

    /**
     * excel日期/字符串转数字字符串
     *
     * @param dateStr 日期字符串
     * @return 这是自1900年1月1日以来的天数。小数天代表小时，分钟和秒。
     */
    public static String dateSlashToNum(String dateStr) {
        return String.valueOf(ExcelUtil.dateToNum(ExcelDateTimeUtil.localDateParseDate(ExcelDateTimeUtil.dateSlashStrParseLocalDate(dateStr))));
    }

    /**
     * excel日期/字符串转数字字符串
     *
     * @param numberStr 要转换的浮点数字符串
     * @return Date
     **/
    public static String numToDateSlash(String numberStr) {
        return ExcelDateTimeUtil.localDateParseDateSlashStr(ExcelDateTimeUtil.dateParseLocalDate(ExcelUtil.numToDate(Double.parseDouble(numberStr))));
    }

    /**
     * excel日期-字符串转数字字符串
     *
     * @param dateStr 日期字符串
     * @return 这是自1900年1月1日以来的天数。小数天代表小时，分钟和秒。
     */
    public static String dateBarToNum(String dateStr) {
        return String.valueOf(ExcelUtil.dateToNum(ExcelDateTimeUtil.localDateParseDate(ExcelDateTimeUtil.dateBarStrParseLocalDate(dateStr))));
    }

    /**
     * excel日期-字符串转数字字符串
     *
     * @param numberStr 要转换的浮点数字符串
     * @return Date
     **/
    public static String numToDateBar(String numberStr) {
        return ExcelDateTimeUtil.localDateParseDateBarStr(ExcelDateTimeUtil.dateParseLocalDate(ExcelUtil.numToDate(Double.parseDouble(numberStr))));
    }

}
