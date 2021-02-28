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

    private final static DateTimeFormatter YEAR_MONTH_DAY_BAR = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private final static DateTimeFormatter YEAR_MONTH_DAY_SLASH = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    public static LocalDate dateSlashStrParseLocalDate(String dateStr) {
        return LocalDate.parse(dateStr, YEAR_MONTH_DAY_SLASH);
    }

    public static String localDateParseDateSlashStr(LocalDate localDate) {
        return localDate.format(YEAR_MONTH_DAY_SLASH);
    }

    public static LocalDate dateBarStrParseLocalDate(String dateStr) {
        return LocalDate.parse(dateStr, YEAR_MONTH_DAY_BAR);
    }

    public static String localDateParseDateBarStr(LocalDate localDate) {
        return localDate.format(YEAR_MONTH_DAY_BAR);
    }

    /**
     * localDate 转 Date
     */
    public static Date localDateParseDate(LocalDate localDate) {
        return Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    /**
     * Date 转 localDate
     */
    public static LocalDate dateParseLocalDate(Date date) {
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

}
