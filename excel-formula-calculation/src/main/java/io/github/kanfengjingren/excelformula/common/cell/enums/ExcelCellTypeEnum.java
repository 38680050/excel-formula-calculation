package io.github.kanfengjingren.excelformula.common.cell.enums;

import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.util.ExcelCellTypeUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * Excel 单元格数据类型枚举。
 * <p>
 * 定义字符串、数字、百分数、两种日期格式的匹配规则，以及与通用数值之间的双向转换方法名。
 * 兼容/收敛转换通过工具类静态方法执行，失败时记录日志并抛出带 #VALUE! 的异常。
 */
public enum ExcelCellTypeEnum {

    /**
     * 字符串
     */
    STRING(1, Pattern.compile("^[\\s\\S]*$"), null, null),
    /**
     * 数字
     */
    NUMBER(2, Pattern.compile("^([-+])?\\d+(\\.\\d+)?$"), null, null),
    /**
     * 百分比
     */
    PERCENTAGE(3, Pattern.compile("^([-+])?\\d+(\\.\\d+)?%$"), "percentageToNum", "numToPercentage"),
    /**
     * 日期（yyyy-MM-dd）
     */
    DATE_BAR(4, Pattern.compile("^\\d{4}-\\d{1,2}-\\d{1,2}$"), "dateBarToNum", "numToDateBar"),
    /**
     * 日期（yyyy/MM/dd）
     */
    DATE_SLASH(5, Pattern.compile("^\\d{4}/\\d{1,2}/\\d{1,2}$"), "dateSlashToNum", "numToDateSlash");

    private final static Logger logger = LoggerFactory.getLogger(ExcelCellTypeEnum.class);

    /**
     * 唯一标识
     */
    private final Integer id;

    /**
     * 字符串匹配 Pattern
     */
    private final Pattern pattern;

    /**
     * 兼容为通用字符方法名
     */
    private final String compatible;

    /**
     * 收敛为原单元格类型方法名
     */
    private final String convergence;

    /**
     * Excel单元格类型转换工具类，方法名称和方法Map
     */
    public final static Map<String, Method> NAME_AND_METHOD_MAP;

    static {
        NAME_AND_METHOD_MAP = new HashMap<>();
        for (Method method : ExcelCellTypeUtil.class.getDeclaredMethods()) {
            NAME_AND_METHOD_MAP.put(method.getName(), method);
        }
    }

    /**
     * 将单元格原值转为计算用基础值（如百分数转小数、日期转序列号）。
     * 无转换方法时原样返回；转换失败抛 #VALUE!。
     */
    public String compatibleValue(String value) {
        return invokeConvert(this.compatible, value);
    }

    /**
     * 将计算用基础值转回单元格原显示形式（如小数转百分数、序列号转日期）。
     * 无转换方法时原样返回；转换失败抛 #VALUE!。
     */
    public String convergenceValue(String value) {
        return invokeConvert(this.convergence, value);
    }

    private static String invokeConvert(String methodName, String value) {
        Method method = NAME_AND_METHOD_MAP.get(methodName);
        if (null == method) {
            return value;
        }
        try {
            return (String) method.invoke(null, value);
        } catch (IllegalAccessException | InvocationTargetException e) {
            Throwable cause = e instanceof InvocationTargetException && e.getCause() != null ? e.getCause() : e;
            logger.error("数据格式转换错误,值:{},格式:{}", value, method.getName(), cause);
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_VALUE,
                    "数据格式转换错误,值:" + value + ",格式:" + method.getName(), cause);
        }
    }

    /**
     * 按配置中的数值 ID 查找单元格类型；不存在时抛异常。
     */
    public static ExcelCellTypeEnum getById(Integer id) {
        for (ExcelCellTypeEnum excelCellTypeEnum : values()) {
            if (excelCellTypeEnum.id.equals(id)) {
                return excelCellTypeEnum;
            }
        }
        throw new ExcelException("未发现 id = " + id + "的单元格类型!");
    }

    /** 类型在配置串中使用的数值 ID。 */
    public Integer getId() {
        return id;
    }

    /** 判断原值是否符合该类型的正则。 */
    public Pattern getPattern() {
        return pattern;
    }

    /** 原值转基础值时在工具类中调用的方法名；无转换时为 null。 */
    public String getCompatible() {
        return compatible;
    }

    /** 基础值转回原值时在工具类中调用的方法名；无转换时为 null。 */
    public String getConvergence() {
        return convergence;
    }

    ExcelCellTypeEnum(Integer id, Pattern pattern, String compatible, String convergence) {
        this.id = id;
        this.pattern = pattern;
        this.compatible = compatible;
        this.convergence = convergence;
    }
}
