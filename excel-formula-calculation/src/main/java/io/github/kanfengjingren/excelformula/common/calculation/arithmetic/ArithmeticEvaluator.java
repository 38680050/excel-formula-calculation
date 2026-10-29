package io.github.kanfengjingren.excelformula.common.calculation.arithmetic;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelErrorUtil;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;

import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;

/**
 * 算术与比较表达式求值器。
 * <p>
 * 使用 BigDecimal 按白名单语法递归下降解析，仅支持数字、加减乘除、括号、一元正负号，
 * 以及比较运算（等于、不等于、大于、小于、大于等于、小于等于）。
 * 布尔结果输出为大写 TRUE / FALSE。
 * 公式语义错误按错误处理模式抛异常或返回错误码字面量。
 */
public final class ArithmeticEvaluator {

    /** 运算精度与舍入方式 */
    private static final MathContext MATH_CONTEXT = new MathContext(16, RoundingMode.HALF_UP);

    /** 待解析的公式原文 */
    private final String source;

    /** 当前解析位置下标 */
    private int index;

    private ArithmeticEvaluator(String source) {
        this.source = source.trim();
    }

    /**
     * 对整段公式求值（默认 THROW 模式）。
     */
    public static String evaluate(String formula) {
        return evaluate(formula, ExcelErrorHandleMode.THROW);
    }

    /**
     * 对整段公式求值并返回字符串结果。
     * 空公式视为配置错误始终抛异常；错误字面量与除零等按模式处理。
     *
     * @param formula 算术或比较表达式
     * @param mode    错误处理模式，null 按 THROW
     * @return 计算结果字符串，或 RETURN_CODE 下的错误码
     */
    public static String evaluate(String formula, ExcelErrorHandleMode mode) {
        ExcelErrorHandleMode effective = ExcelErrorHandleMode.orThrow(mode);
        if (formula == null || formula.isBlank()) {
            throw ExcelErrorUtil.configError("公式不能为空");
        }
        if (ExcelErrorUtil.isError(formula)) {
            return ExcelErrorUtil.fail(effective, formula, "公式错误: " + formula);
        }
        try {
            ArithmeticEvaluator evaluator = new ArithmeticEvaluator(formula);
            Value value = evaluator.parseExpression();
            evaluator.skipWhitespace();
            if (!evaluator.isAtEnd()) {
                return ExcelErrorUtil.fail(effective, ExcelConstant.EXCEL_ERROR_VALUE,
                        "无法解析的公式片段: " + evaluator.source.substring(evaluator.index));
            }
            return value.toResultString();
        } catch (ExcelException e) {
            return ExcelErrorUtil.handleThrown(effective, e);
        }
    }

    /**
     * 解析比较表达式：先算左右算术值，再按运算符比较。
     * 优先匹配双字符运算符，避免把 &gt;= 拆成 &gt; 与 =。
     */
    private Value parseExpression() {
        Value left = parseAdditive();
        skipWhitespace();
        if (matchOperator("<>")) {
            return Value.bool(!left.asNumber().equals(parseAdditive().asNumber()));
        }
        if (matchOperator(">=")) {
            return Value.bool(left.asNumber().compareTo(parseAdditive().asNumber()) >= 0);
        }
        if (matchOperator("<=")) {
            return Value.bool(left.asNumber().compareTo(parseAdditive().asNumber()) <= 0);
        }
        if (matchOperator(">")) {
            return Value.bool(left.asNumber().compareTo(parseAdditive().asNumber()) > 0);
        }
        if (matchOperator("<")) {
            return Value.bool(left.asNumber().compareTo(parseAdditive().asNumber()) < 0);
        }
        if (matchOperator("=")) {
            return Value.bool(left.asNumber().compareTo(parseAdditive().asNumber()) == 0);
        }
        return left;
    }

    /** 解析加减运算，左结合 */
    private Value parseAdditive() {
        Value left = parseMultiplicative();
        while (true) {
            skipWhitespace();
            if (matchOperator("+")) {
                left = Value.number(left.asNumber().add(parseMultiplicative().asNumber(), MATH_CONTEXT));
            } else if (matchOperator("-")) {
                left = Value.number(left.asNumber().subtract(parseMultiplicative().asNumber(), MATH_CONTEXT));
            } else {
                return left;
            }
        }
    }

    /** 解析乘除运算；除数为零时抛出带 #DIV/0! 的异常 */
    private Value parseMultiplicative() {
        Value left = parseUnary();
        while (true) {
            skipWhitespace();
            if (matchOperator("*")) {
                left = Value.number(left.asNumber().multiply(parseUnary().asNumber(), MATH_CONTEXT));
            } else if (matchOperator("/")) {
                BigDecimal divisor = parseUnary().asNumber();
                if (divisor.compareTo(BigDecimal.ZERO) == 0) {
                    throw new ExcelException(ExcelConstant.EXCEL_ERROR_DIV0, "除零");
                }
                left = Value.number(left.asNumber().divide(divisor, MATH_CONTEXT));
            } else {
                return left;
            }
        }
    }

    /** 解析一元正负号 */
    private Value parseUnary() {
        skipWhitespace();
        if (matchOperator("+")) {
            return parseUnary();
        }
        if (matchOperator("-")) {
            return Value.number(parseUnary().asNumber().negate(MATH_CONTEXT));
        }
        return parsePrimary();
    }

    /** 解析括号、布尔关键字或数字字面量 */
    private Value parsePrimary() {
        skipWhitespace();
        if (matchOperator("(")) {
            Value value = parseExpression();
            skipWhitespace();
            if (!matchOperator(")")) {
                throw new ExcelException(ExcelConstant.EXCEL_ERROR_VALUE, "缺少右括号");
            }
            return value;
        }
        if (matchKeyword(ExcelConstant.BIG_CASE_TRUE) || matchKeyword(ExcelConstant.LOWER_CASE_TRUE)) {
            return Value.bool(true);
        }
        if (matchKeyword(ExcelConstant.BIG_CASE_FALSE) || matchKeyword(ExcelConstant.LOWER_CASE_FALSE)) {
            return Value.bool(false);
        }
        return Value.number(parseNumber());
    }

    /** 读取连续数字与小数点，构造 BigDecimal */
    private BigDecimal parseNumber() {
        skipWhitespace();
        int start = index;
        while (!isAtEnd() && (Character.isDigit(peek()) || peek() == '.')) {
            index++;
        }
        if (start == index) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_VALUE,
                    "期望数字，位置: " + index + ", 公式: " + source);
        }
        String number = source.substring(start, index);
        try {
            return new BigDecimal(number);
        } catch (NumberFormatException e) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_VALUE, "非法数字: " + number, e);
        }
    }

    /**
     * 尝试匹配运算符；单独匹配 = 时若遇到 == 则按等于处理并消费两个字符。
     */
    private boolean matchOperator(String op) {
        skipWhitespace();
        if (source.startsWith(op, index)) {
            if ("=".equals(op) && source.startsWith("==", index)) {
                index += 2;
                return true;
            }
            index += op.length();
            return true;
        }
        return false;
    }

    /** 尝试匹配布尔关键字（忽略大小写，且后面不能紧跟字母数字） */
    private boolean matchKeyword(String keyword) {
        skipWhitespace();
        if (source.regionMatches(true, index, keyword, 0, keyword.length())) {
            int end = index + keyword.length();
            if (end < source.length() && Character.isLetterOrDigit(source.charAt(end))) {
                return false;
            }
            index = end;
            return true;
        }
        return false;
    }

    private void skipWhitespace() {
        while (!isAtEnd() && Character.isWhitespace(peek())) {
            index++;
        }
    }

    private char peek() {
        return source.charAt(index);
    }

    private boolean isAtEnd() {
        return index >= source.length();
    }

    /**
     * 求值中间结果：要么是数值，要么是布尔。
     * 参与算术时布尔按 1 / 0 处理。
     */
    private static final class Value {
        private final BigDecimal number;
        private final Boolean bool;

        private Value(BigDecimal number, Boolean bool) {
            this.number = number;
            this.bool = bool;
        }

        static Value number(BigDecimal number) {
            return new Value(number, null);
        }

        static Value bool(boolean value) {
            return new Value(null, value);
        }

        /** 转为数值；布尔 true 为 1，false 为 0 */
        BigDecimal asNumber() {
            if (number != null) {
                return number;
            }
            return bool ? BigDecimal.ONE : BigDecimal.ZERO;
        }

        /**
         * 输出最终字符串：布尔为大写 TRUE/FALSE；
         * 数值去掉无意义的尾随零，使用普通十进制写法。
         */
        String toResultString() {
            if (bool != null) {
                return bool ? ExcelConstant.BIG_CASE_TRUE : ExcelConstant.BIG_CASE_FALSE;
            }
            BigDecimal normalized = number.stripTrailingZeros();
            if (normalized.scale() < 0) {
                normalized = normalized.setScale(0);
            }
            return normalized.toPlainString();
        }
    }
}
