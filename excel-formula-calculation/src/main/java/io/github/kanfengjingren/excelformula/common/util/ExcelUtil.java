package io.github.kanfengjingren.excelformula.common.util;

import io.github.kanfengjingren.excelformula.common.calculation.arithmetic.ArithmeticEvaluator;
import io.github.kanfengjingren.excelformula.common.cell.enums.ExcelCellTypeEnum;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.exception.ExcelErrorUtil;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;

import java.util.*;
import java.util.stream.Collectors;

/**
 * Excel 通用工具类。
 * <p>
 * 提供单元格类型解析、日期序列互转、公式中单元格替换与算术求值、列字母与序号互转等能力。
 * 算术求值委托给 BigDecimal 表达式求值器，避免浮点精度问题。
 */
public class ExcelUtil {

    /**
     * 每分钟60秒
     */
    public static final int SECONDS_PER_MINUTE = 60;

    /**
     * 每小时60分钟
     */
    public static final int MINUTES_PER_HOUR = 60;

    /**
     * 每天24小时
     */
    public static final int HOURS_PER_DAY = 24;

    /**
     * 每天86,400秒
     */
    public static final int SECONDS_PER_DAY = (HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE);

    /** 日期序列无效时的哨兵值。 */
    private static final int BAD_DATE = -1;

    /** 一自然日的毫秒数，用于 Excel 日期序列与小数时间部分的换算。 */
    public static final long DAY_MILLISECONDS = SECONDS_PER_DAY * 1000L;

    /**
     * 解析行内单元格类型
     *
     * @param t                    行实体类
     * @param excelCalculateConfig Excel 计算配置
     * @param <T>                  <T>
     * @return 行内单元格类型
     */
    public static <T> Map<String, ExcelCellTypeEnum> parseColumnAndType(T t, ExcelCalculateConfig excelCalculateConfig) {
        HashMap<String, ExcelCellTypeEnum> columnAndTypeMap = new HashMap<>();
        if (excelCalculateConfig.getCellTypesField() == null) {
            return columnAndTypeMap;
        }
        String cellTypes = ExcelReflectUtil.getValue(t, excelCalculateConfig.getCellTypesField());
        if (ExcelStrUtil.isNotBlank(cellTypes)) {
            for (String columnAndType : cellTypes.split(ExcelConstant.SEMICOLON)) {
                if (ExcelStrUtil.isBlank(columnAndType)) {
                    continue;
                }
                String[] columnAndTypeId = columnAndType.split(ExcelConstant.MINUS);
                if (columnAndTypeId.length < 2) {
                    throw new ExcelException("单元格类型配置格式错误: " + columnAndType);
                }
                try {
                    columnAndTypeMap.put(columnAndTypeId[0],
                            ExcelCellTypeEnum.getById(Integer.parseInt(columnAndTypeId[1])));
                } catch (NumberFormatException e) {
                    throw new ExcelException("单元格类型配置格式错误: " + columnAndType, e);
                }
            }
        }
        return columnAndTypeMap;
    }

    /**
     * excel日期转数字
     *
     * @param date 日期
     * @return 这是自1900年1月1日以来的天数。小数天代表小时，分钟和秒。
     */
    public static double dateToNum(Date date) {
        return dateToNum(date, false);
    }

    /**
     * 给定日期，将其转换为代表其内部Excel表示形式的双精度数，这是自1900年1月1日以来的天数。小数天代表小时，分钟和秒。
     *
     * @param date             日期
     * @param use1904windowing 应该使用1900还是1904年日期窗口？
     * @return 日期的Excel表示形式（如果错误，则为-1-通过检查小于0.1来测试错误）
     */
    public static double dateToNum(Date date, boolean use1904windowing) {
        Calendar calStart = Calendar.getInstance();
        // If date includes hours, minutes, and seconds, set them to 0
        calStart.setTime(date);
        return internalGetExcelDate(calStart, use1904windowing);
    }

    private static double internalGetExcelDate(Calendar date, boolean use1904windowing) {
        if ((!use1904windowing && date.get(Calendar.YEAR) < 1900) ||
                (use1904windowing && date.get(Calendar.YEAR) < 1904)) {
            return BAD_DATE;
        }
        // 由于节省了夏时制，我们不能使用date.getTime（）-calStart.getTimeInMillis（），
        // 因为00:00和04:00之间的毫秒差可以是3、4或5小时，但Excel希望它始终为4小时。例如。
        // 2004-03-28 04:00 CEST-2004-03-28 00:00 CET是3小时，而2004-10-31 04:00 CET-2004-10-31 00:00 CEST是5小时
        double fraction = (((date.get(Calendar.HOUR_OF_DAY) * 60
                + date.get(Calendar.MINUTE)
        ) * 60 + date.get(Calendar.SECOND)
        ) * 1000 + date.get(Calendar.MILLISECOND)
        ) / (double) DAY_MILLISECONDS;
        Calendar calStart = dayStart(date);

        double value = fraction + absoluteDay(calStart, use1904windowing);

        if (!use1904windowing && value >= 60) {
            value++;
        } else if (use1904windowing) {
            value--;
        }

        return value;
    }

    /**
     * 返回自1900年以来以前年份的天数
     *
     * @param yr               a year (1900 < yr < 4000)
     * @param use1904windowing
     * @return days  number of days in years prior to yr.
     * @throws ExcelException if year is outside of range.
     */
    private static int daysInPriorYears(int yr, boolean use1904windowing) {
        if ((!use1904windowing && yr < 1900) || (use1904windowing && yr < 1904)) {
            throw new ExcelException("'year' must be 1900 or greater");
        }

        int yr1 = yr - 1;
        // plus julian leap days in prior years
        int leapDays = yr1 / 4
                - yr1 / 100 // minus prior century years
                + yr1 / 400 // plus years divisible by 400
                - 460;      // leap days in previous 1900 years

        return 365 * (yr - (use1904windowing ? 1904 : 1900)) + leapDays;
    }

    /**
     * Given a Calendar, return the number of days since 1900/12/31.
     *
     * @param cal the Calendar
     * @return days number of days since 1900/12/31
     * @throws IllegalArgumentException if date is invalid
     */
    protected static int absoluteDay(Calendar cal, boolean use1904windowing) {
        return cal.get(Calendar.DAY_OF_YEAR)
                + daysInPriorYears(cal.get(Calendar.YEAR), use1904windowing);
    }

    /**
     * 将cal的HH：MM：SS字段设置为00：00：00：000
     *
     * @param cal
     * @return
     */
    private static Calendar dayStart(final Calendar cal) {
        // 强制重新计算内部字段
        cal.get(Calendar.HOUR_OF_DAY);
        cal.set(Calendar.HOUR_OF_DAY, 0);
        cal.set(Calendar.MINUTE, 0);
        cal.set(Calendar.SECOND, 0);
        cal.set(Calendar.MILLISECOND, 0);
        // 强制重新计算内部字段
        cal.get(Calendar.HOUR_OF_DAY);
        return cal;
    }

    /**
     * 转换方法
     *
     * @param number 要转换的浮点数
     * @return Date
     **/
    public static Date numToDate(double number) {
        int wholeDays = (int) Math.floor(number);
        int millisecondsInday = (int) ((number - wholeDays) * DAY_MILLISECONDS + 0.5);
        Calendar calendar = new GregorianCalendar();
        setCalendar(calendar, wholeDays, millisecondsInday, false);
        return calendar.getTime();
    }

    private static void setCalendar(Calendar calendar, int wholeDays, int millisecondsInDay, boolean use1904windowing) {
        int startYear = 1900;
        // Excel认为2/29/1900是有效日期，但不是
        int dayAdjust = -1;
        if (use1904windowing) {
            startYear = 1904;
            // 1904年日期视窗将1/2/1904作为第一天
            dayAdjust = 1;
        } else if (wholeDays < 61) {
            // 日期早于1900年3月1日，因此请进行调整，因为Excel认为存在2/29/1900如果Excel日期== 2/29/1900，则Java表示将变为3/1/1900
            dayAdjust = 0;
        }
        calendar.set(startYear, 0, wholeDays + dayAdjust, 0, 0, 0);
        calendar.set(GregorianCalendar.MILLISECOND, millisecondsInDay);
    }

    /**
     * 函数参数场景下的求值（默认 THROW）。
     */
    public static String functionCalculate(String formula, Map<String, String> cellAndValueMap) {
        return functionCalculate(formula, cellAndValueMap, ExcelErrorHandleMode.THROW);
    }

    /**
     * 函数参数场景下的求值：若公式本身就是唯一单元格且值为 null，则返回 null（不转成 0）；
     * 其它情况走单元格替换后的算术求值。
     *
     * @param formula         参数表达式
     * @param cellAndValueMap 单元格引用到值的映射
     * @param mode            错误处理模式
     * @return 计算结果，或空单元格对应的 null
     */
    public static String functionCalculate(String formula, Map<String, String> cellAndValueMap, ExcelErrorHandleMode mode) {
        if (1 == cellAndValueMap.size()) {
            Map.Entry<String, String> onlyEntry = cellAndValueMap.entrySet().stream().findAny().orElse(null);
            if (onlyEntry != null && formula.equals(onlyEntry.getKey()) && onlyEntry.getValue() == null) {
                return null;
            }
            if (onlyEntry != null && formula.equals(onlyEntry.getKey()) && ExcelErrorUtil.isError(onlyEntry.getValue())) {
                return ExcelErrorUtil.fail(mode, onlyEntry.getValue(), "单元格错误: " + onlyEntry.getValue());
            }
        }
        return arithmeticCalculate(formula, cellAndValueMap, mode);
    }

    /**
     * 将公式中的单元格引用替换为值后做算术求值（默认 THROW）。
     */
    public static String arithmeticCalculate(String formula, Map<String, String> cellAndValueMap) {
        return arithmeticCalculate(formula, cellAndValueMap, ExcelErrorHandleMode.THROW);
    }

    /**
     * 将公式中的单元格引用替换为值后做算术求值。
     * 先按引用字符串长度从长到短替换，避免短引用误替换长引用的前缀（例如 A1 误伤 A10）。
     * 空白单元格按 0 参与运算；若单元格已是错误字面量则按模式短路。
     *
     * @param formula         含单元格引用的算术公式
     * @param cellAndValueMap 单元格引用到值的映射
     * @param mode            错误处理模式
     * @return 计算结果
     */
    public static String arithmeticCalculate(String formula, Map<String, String> cellAndValueMap, ExcelErrorHandleMode mode) {
        List<Map.Entry<String, String>> entries = cellAndValueMap.entrySet().stream()
                .sorted((a, b) -> Integer.compare(b.getKey().length(), a.getKey().length()))
                .collect(Collectors.toList());
        for (Map.Entry<String, String> cellAndValue : entries) {
            String value = cellAndValue.getValue();
            if (ExcelErrorUtil.isError(value)) {
                return ExcelErrorUtil.fail(mode, value, "单元格错误: " + value);
            }
            if (ExcelStrUtil.isBlank(value)) {
                value = ExcelConstant.ZERO_STR;
            }
            formula = formula.replace(cellAndValue.getKey(), value);
        }
        return calculate(formula, mode);
    }

    /**
     * 对已不含单元格引用的算术/比较表达式求值（默认 THROW）。
     */
    public static String calculate(String formula) {
        return calculate(formula, ExcelErrorHandleMode.THROW);
    }

    /**
     * 对已不含单元格引用的算术/比较表达式求值。
     *
     * @param formula 纯表达式字符串
     * @param mode    错误处理模式
     * @return 计算结果
     */
    public static String calculate(String formula, ExcelErrorHandleMode mode) {
        return ArithmeticEvaluator.evaluate(formula, mode);
    }

    /**
     * 获取中间所有列
     *
     * @param startColumn 开始列
     * @param endColumn   结束列
     * @return List<Column>
     */
    public static List<String> getBetweenColumnList(String startColumn, String endColumn) {
        List<String> betweenColumnList = new ArrayList<>();
        for (int i = columnToIndex(startColumn); i <= columnToIndex(endColumn); i++) {
            betweenColumnList.add(indexToColumn(i));
        }
        return betweenColumnList;
    }

    /**
     * 该方法用来将Excel中的ABCD列转换成具体的数据
     *
     * @param column 列名称
     * @return 将字母列名称转换成数字
     **/
    public static int columnToIndex(String column) {
        int num;
        int result = 0;
        int length = column.length();
        for (int i = 0; i < length; i++) {
            char ch = column.charAt(length - i - 1);
            num = ch - 'A' + 1;
            num *= Math.pow(26, i);
            result += num;
        }
        return result;
    }

    /**
     * 该方法用来将具体的数据转换成Excel中的ABCD列
     *
     * @param index 需要转换成字母的数字
     * @return column 列名称
     **/
    public static String indexToColumn(int index) {
        if (index <= 0) {
            throw new ExcelException("index 不得小于等于 0");
        }
        StringBuilder columnStr = new StringBuilder();
        index--;
        do {
            if (columnStr.length() > 0) {
                index--;
            }
            columnStr.insert(0, ((char) (index % 26 + (int) 'A')));
            index = (index - index % 26) / 26;
        } while (index > 0);
        return columnStr.toString();
    }

}
