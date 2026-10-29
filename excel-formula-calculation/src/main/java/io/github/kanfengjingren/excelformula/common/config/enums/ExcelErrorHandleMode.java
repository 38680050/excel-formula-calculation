package io.github.kanfengjingren.excelformula.common.config.enums;

/**
 * 公式语义错误的处理方式。
 * <p>
 * 配置错误、非法 API 用法、反射等基础设施故障不受本模式影响，始终抛异常。
 */
public enum ExcelErrorHandleMode {

    /**
     * 默认：公式错误抛出带错误码的 ExcelException，便于 Java 调用方 try/catch。
     */
    THROW,

    /**
     * 公式错误作为字符串结果返回（如 #DIV/0!），并在后续运算中短路传播。
     */
    RETURN_CODE;

    /**
     * null 时回退为 {@link #THROW}，避免各处重复写三元表达式。
     */
    public static ExcelErrorHandleMode orThrow(ExcelErrorHandleMode mode) {
        return mode == null ? THROW : mode;
    }
}
