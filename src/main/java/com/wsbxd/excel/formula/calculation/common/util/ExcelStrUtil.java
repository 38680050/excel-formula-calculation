package com.wsbxd.excel.formula.calculation.common.util;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * description: Excel 字符串工具类
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelStrUtil {

    /**
     * 字符串是否为空，空的定义如下:<br>
     * 1、为null <br>
     * 2、为""<br>
     *
     * @param str 被检测的字符串
     * @return 是否为空
     */
    public static boolean isEmpty(CharSequence str) {
        return str == null || str.length() == 0;
    }

    /**
     * 字符串是否为非空白 空白的定义如下： <br>
     * 1、不为null <br>
     * 2、不为""<br>
     *
     * @param str 被检测的字符串
     * @return 是否为非空
     */
    public static boolean isNotEmpty(CharSequence str) {
        return !isEmpty(str);
    }

    /**
     * 字符串是否为非空白 空白的定义如下： <br>
     * 1、不为null <br>
     * 2、不为不可见字符（如空格）<br>
     * 3、不为""<br>
     *
     * @param str 被检测的字符串
     * @return 是否为非空
     */
    public static boolean isNotBlank(CharSequence str) {
        return !isBlank(str);
    }

    /**
     * 字符串是否为空白 空白的定义如下： <br>
     * 1、为null <br>
     * 2、为不可见字符（如空格）<br>
     * 3、""<br>
     *
     * @param str 被检测的字符串
     * @return 是否为空
     */
    public static boolean isBlank(CharSequence str) {
        int length;
        if ((str == null) || ((length = str.length()) == 0)) {
            return true;
        }
        for (int i = 0; i < length; i++) {
            // 只要有一个非空字符即为非空字符串
            if (!isBlankChar(str.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    /**
     * 切分字符串<br>
     * a#b#c -> [a,b,c] <br>
     * a##b#c -> [a,"",b,c]
     *
     * @param str       被切分的字符串
     * @param separator 分隔符字符
     * @return 切分后的集合
     */
    public static List<String> split(String str, char separator) {
        return split(str, separator, 0);
    }

    /**
     * 切分字符串
     *
     * @param str       被切分的字符串
     * @param separator 分隔符字符
     * @param limit     限制分片数
     * @return 切分后的集合
     */
    public static List<String> split(String str, char separator, int limit) {
        if (str == null) {
            return null;
        }
        List<String> list = new ArrayList<String>(limit == 0 ? 16 : limit);
        if (limit == 1) {
            list.add(str);
            return list;
        }
        // 未结束切分的标志
        boolean isNotEnd = true;
        int strLen = str.length();
        StringBuilder sb = new StringBuilder(strLen);
        for (int i = 0; i < strLen; i++) {
            char c = str.charAt(i);
            if (isNotEnd && c == separator) {
                list.add(sb.toString());
                // 清空StringBuilder
                sb.delete(0, sb.length());

                // 当达到切分上限-1的量时，将所剩字符全部作为最后一个串
                if (limit != 0 && list.size() == limit - 1) {
                    isNotEnd = false;
                }
            } else {
                sb.append(c);
            }
        }
        // 加入尾串
        list.add(sb.toString());
        return list;
    }

    /**
     * 字符串集合合计计算
     *
     * @param numList 字符串数字集合
     * @return 合计结果
     */
    public static String sum(List<String> numList) {
        BigDecimal decimal = new BigDecimal(0);
        for (String value : numList) {
            decimal = decimal.add(getBigDecimal(value));
        }
        return decimal.toString();
    }

    /**
     * 根据正则进行字符串匹配
     *
     * @param source 需要匹配的字符串
     * @param regex  正则表达式
     * @return 匹配到的结果集
     */
    public static List<String> getMatcher(String source, String regex) {
        return getMatcher(source, Pattern.compile(regex));
    }

    /**
     * 根据正则进行字符串匹配
     *
     * @param source 需要匹配的字符串
     * @param regex  正则表达式Pattern
     * @return 匹配到的结果集
     */
    public static List<String> getMatcher(String source, Pattern regex) {
        List<String> matchStrList = new ArrayList<>();
        Matcher matcher = regex.matcher(source);
        while (matcher.find()) {
            matchStrList.add(matcher.group(0));
        }
        return matchStrList;
    }

    /**
     * 字符串转BigDecimal，如果为空或为空字符串则默认为0
     *
     * @param str 将要转换的字符串
     * @return 转换后的BigDecimal
     */
    public static BigDecimal getBigDecimal(String str) {
        return ExcelStrUtil.isBlank(str) ? BigDecimal.ZERO : new BigDecimal(str);
    }

    public static String toString(Object obj) {
        return null == obj ? "" : obj.toString();
    }

    /**
     * 是否空白符<br>
     * 空白符包括空格、制表符、全角空格和不间断空格<br>
     *
     * @param c 字符
     * @return 是否空白符
     * @see Character#isWhitespace(int)
     * @see Character#isSpaceChar(int)
     */
    public static boolean isBlankChar(char c) {
        return isBlankChar((int) c);
    }

    /**
     * 是否空白符<br>
     * 空白符包括空格、制表符、全角空格和不间断空格<br>
     *
     * @param c 字符
     * @return 是否空白符
     * @see Character#isWhitespace(int)
     * @see Character#isSpaceChar(int)
     */
    public static boolean isBlankChar(int c) {
        return Character.isWhitespace(c) || Character.isSpaceChar(c) || c == '\ufeff' || c == '\u202a';
    }

}
