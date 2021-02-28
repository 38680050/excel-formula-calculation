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

    public final static String DOT = ",";

    public final static String SEMICOLON = ";";

    public final static String MINUS = "-";

    public final static char DOT_CHAT = ',';

    public final static String COLON = ":";

    public final static String EQUALS_SIGN = "=";

    public final static String DOUBLE_EQUALS_SIGN = "==";

    public final static String TRUE = "true";

    public final static String NULL_CHAR = "";

    public final static String EXCLAMATION_MARK = "!";

    public final static String APOSTROPHE = "'";

    public final static String BAR = "-";

    public final static String LEFT_PARENTHESIS = "(";

    public final static String RIGHT_PARENTHESIS = ")";

    public final static String LOWER_CASE_TRUE = "true";

    public final static String LOWER_CASE_FALSE = "false";

    public final static String BIG_CASE_TRUE = "TRUE";

    public final static String BIG_CASE_FALSE = "FALSE";

    public final static List<String> TRUE_AND_FALSE;

    public final static List<String> EXCEL_ERROR_VALUE;

    public final static String ZERO = "0";

    static {
        TRUE_AND_FALSE = new ArrayList<>();
        TRUE_AND_FALSE.add(BIG_CASE_TRUE);
        TRUE_AND_FALSE.add(BIG_CASE_FALSE);
        EXCEL_ERROR_VALUE = new ArrayList<>();
        EXCEL_ERROR_VALUE.add("#DIV/0!");
        EXCEL_ERROR_VALUE.add("#VALUE!");
        EXCEL_ERROR_VALUE.add("#REF!");
        EXCEL_ERROR_VALUE.add("#N/A");
        EXCEL_ERROR_VALUE.add("#NAME?");
        EXCEL_ERROR_VALUE.add("#NUM!");
        EXCEL_ERROR_VALUE.add("#NULL");
    }

}
