package com.wsbxd.excel.formula.calculation.common.constant;

import java.util.ArrayList;
import java.util.List;

/**
 * description: Excel 常量
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelConstant {

    /**
     * 英文逗号
     */
    public final static String DOT = ",";

    /**
     * 分号
     */
    public final static String SEMICOLON = ";";

    /**
     * 减号
     */
    public final static String MINUS = "-";

    /**
     * char类型英文逗号
     */
    public final static char DOT_CHAT = ',';

    /**
     * 冒号
     */
    public final static String COLON = ":";

    /**
     * 等于号
     */
    public final static String EQUALS_SIGN = "=";

    /**
     * 双等于号
     */
    public final static String DOUBLE_EQUALS_SIGN = "==";

    /**
     * 空字符串
     */
    public final static String NULL_STR = "";

    /**
     * 感叹号
     */
    public final static String EXCLAMATION_MARK = "!";

    /**
     * 单引号
     */
    public final static String APOSTROPHE = "'";

    /**
     * 左英文括号
     */
    public final static String LEFT_PARENTHESIS = "(";

    /**
     * 右英文括号
     */
    public final static String RIGHT_PARENTHESIS = ")";

    /**
     * 小写 布尔true类型
     */
    public final static String LOWER_CASE_TRUE = "true";

    /**
     * 小写 布尔false类型
     */
    public final static String LOWER_CASE_FALSE = "false";

    /**
     * 大写 布尔true类型
     */
    public final static String BIG_CASE_TRUE = "TRUE";

    /**
     * 大写 布尔false类型
     */
    public final static String BIG_CASE_FALSE = "FALSE";

    /**
     * 大写布尔类型集合
     */
    public final static List<String> BIG_TRUE_AND_FALSE;

    /**
     * EXCEL 错误类型值集合
     */
    public final static List<String> EXCEL_ERROR_VALUE_LIST;

    /**
     * 字符串零
     */
    public final static String ZERO_STR = "0";

    /**
     * EXCEL 异常 #DIV/0!
     */
    public final static String EXCEL_ERROR_DIV0 = "#DIV/0!";

    /**
     * EXCEL 异常 #VALUE!
     */
    public final static String EXCEL_ERROR_VALUE = "#VALUE!";

    /**
     * EXCEL 异常 #REF!
     */
    public final static String EXCEL_ERROR_REF = "#REF!";

    /**
     * EXCEL 异常 #N/A
     */
    public final static String EXCEL_ERROR_NA = "#N/A";

    /**
     * EXCEL 异常 #NAME?
     */
    public final static String EXCEL_ERROR_NAME = "#NAME?";

    /**
     * EXCEL 异常 #NUM!
     */
    public final static String EXCEL_ERROR_NUM = "#NUM!";

    /**
     * EXCEL 异常 #NULL
     */
    public final static String EXCEL_ERROR_NULL = "#NULL";

    static {
        BIG_TRUE_AND_FALSE = new ArrayList<>();
        BIG_TRUE_AND_FALSE.add(BIG_CASE_TRUE);
        BIG_TRUE_AND_FALSE.add(BIG_CASE_FALSE);
        EXCEL_ERROR_VALUE_LIST = new ArrayList<>();
        EXCEL_ERROR_VALUE_LIST.add(EXCEL_ERROR_DIV0);
        EXCEL_ERROR_VALUE_LIST.add(EXCEL_ERROR_VALUE);
        EXCEL_ERROR_VALUE_LIST.add(EXCEL_ERROR_REF);
        EXCEL_ERROR_VALUE_LIST.add(EXCEL_ERROR_NA);
        EXCEL_ERROR_VALUE_LIST.add(EXCEL_ERROR_NAME);
        EXCEL_ERROR_VALUE_LIST.add(EXCEL_ERROR_NUM);
        EXCEL_ERROR_VALUE_LIST.add(EXCEL_ERROR_NULL);
    }

}
