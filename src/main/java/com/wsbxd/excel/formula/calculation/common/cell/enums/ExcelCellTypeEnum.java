package com.wsbxd.excel.formula.calculation.common.cell.enums;

import com.wsbxd.excel.formula.calculation.common.util.ExcelCellTypeUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * description: Excel 单元格坐标
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/26 22:01
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

    public String compatibleValue(String value) {
        Method method = NAME_AND_METHOD_MAP.get(this.compatible);
        //没方法就默认原值
        if (null == method) {
            return value;
        }
        try {
            value = (String) method.invoke(ExcelCellTypeUtil.class.newInstance(), value);
        } catch (Exception e) {
            logger.error("数据格式转换错误,值:{},格式:{}", value, method.getName());
            e.printStackTrace();
        }
        return value;
    }

    public String convergenceValue(String value) {
        Method method = NAME_AND_METHOD_MAP.get(this.convergence);
        //没方法就默认原值
        if (null == method) {
            return value;
        }
        try {
            value = (String) method.invoke(ExcelCellTypeUtil.class.newInstance(), value);
        } catch (IllegalAccessException | InvocationTargetException | InstantiationException e) {
            logger.error("数据格式转换错误,值:{},格式:{}", value, method.getName());
            e.printStackTrace();
        }
        return value;
    }

    public static ExcelCellTypeEnum getById(Integer id) {
        for (ExcelCellTypeEnum excelCellTypeEnum : values()) {
            if (excelCellTypeEnum.id.equals(id)) {
                return excelCellTypeEnum;
            }
        }
        throw new RuntimeException("未发现 id = " + id + "的单元格类型!");
    }

    public static ExcelCellTypeEnum getByPattern(String value) {
        for (ExcelCellTypeEnum excelCellTypeEnum : values()) {
            if (excelCellTypeEnum.pattern.matcher(value).matches()) {
                return excelCellTypeEnum;
            }
        }
        throw new RuntimeException("未匹配 value = " + value + "的单元格类型!");
    }

    public Integer getId() {
        return id;
    }

    public Pattern getPattern() {
        return pattern;
    }

    public String getCompatible() {
        return compatible;
    }

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
