package com.wsbxd.excel.formula.calculation.common.util;

import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * description: Excel 日期时间工具类
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelDateTimeUtil {

    /**
     * 横杠分割年月日格式器
     */
    private final static DateTimeFormatter YEAR_MONTH_DAY_BAR = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    /**
     * 斜杠分割年月日格式器
     */
    private final static DateTimeFormatter YEAR_MONTH_DAY_SLASH = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    /**
     * 斜杠分割年月日格式字符串转本地日期
     *
     * @param dateStr 斜杠分割年月日格式字符串
     * @return 本地日期
     */
    public static LocalDate dateSlashStrParseLocalDate(String dateStr) {
        return LocalDate.parse(dateStr, YEAR_MONTH_DAY_SLASH);
    }

    /**
     * 本地日期转斜杠分割年月日格式字符串
     *
     * @param localDate 本地日期
     * @return 斜杠分割年月日格式字符串
     */
    public static String localDateParseDateSlashStr(LocalDate localDate) {
        return localDate.format(YEAR_MONTH_DAY_SLASH);
    }

    /**
     * 横杠分割年月日格式字符串转本地日期
     *
     * @param dateStr 横杠分割年月日格式字符串
     * @return 本地日期
     */
    public static LocalDate dateBarStrParseLocalDate(String dateStr) {
        return LocalDate.parse(dateStr, YEAR_MONTH_DAY_BAR);
    }

    /**
     * 本地日期转横杠分割年月日格式字符串
     *
     * @param localDate 本地日期
     * @return 横杠分割年月日格式字符串
     */
    public static String localDateParseDateBarStr(LocalDate localDate) {
        return localDate.format(YEAR_MONTH_DAY_BAR);
    }

    /**
     * localDate 转 Date
     *
     * @param localDate 本地日期
     * @return Date
     */
    public static Date localDateParseDate(LocalDate localDate) {
        return Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    /**
     * Date 转 localDate
     *
     * @param date Date
     * @return 本地日期
     */
    public static LocalDate dateParseLocalDate(Date date) {
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

}
