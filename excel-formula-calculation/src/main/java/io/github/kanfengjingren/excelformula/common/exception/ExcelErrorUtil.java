package io.github.kanfengjingren.excelformula.common.exception;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;

/**
 * 公式语义错误统一出口。
 * <p>
 * 按错误处理模式决定抛出带错误码的异常，或返回错误字面量字符串。
 * 配置或编程错误请通过 {@link #configError} 构造异常并由调用方抛出，与 THROW/RETURN_CODE 无关。
 */
public final class ExcelErrorUtil {

    private ExcelErrorUtil() {
    }

    /**
     * 判断字符串是否为内置 Excel 错误字面量（如 {@code #N/A}）。
     */
    public static boolean isError(String value) {
        return value != null && ExcelConstant.EXCEL_ERROR_VALUE_LIST.contains(value);
    }

    /**
     * 按配置中的错误处理模式处理语义错误；模式缺省为 THROW。
     */
    public static String fail(ExcelCalculateConfig config, String code, String message) {
        return fail(config == null ? null : config.getErrorHandleMode(), code, message, null);
    }

    /**
     * 按指定错误处理模式处理语义错误；{@link ExcelErrorHandleMode#THROW} 抛带错误码的异常，
     * {@link ExcelErrorHandleMode#RETURN_CODE} 返回错误码字符串。
     */
    public static String fail(ExcelErrorHandleMode mode, String code, String message) {
        return fail(mode, code, message, null);
    }

    /**
     * 按指定错误处理模式处理语义错误；THROW 时抛 {@link ExcelException}（可附带 cause），
     * RETURN_CODE 时返回 {@code code}，不抛异常。
     */
    public static String fail(ExcelErrorHandleMode mode, String code, String message, Throwable cause) {
        ExcelErrorHandleMode effective = ExcelErrorHandleMode.orThrow(mode);
        if (effective == ExcelErrorHandleMode.RETURN_CODE) {
            return code;
        }
        String msg = message != null ? message : code;
        throw cause == null ? new ExcelException(code, msg) : new ExcelException(code, msg, cause);
    }

    /**
     * 捕获路径上的二次分发：若异常带错误码且模式为 RETURN_CODE，则返回该错误码；否则原样抛出。
     */
    public static String handleThrown(ExcelErrorHandleMode mode, ExcelException e) {
        if (e.getExcelErrorCode() != null) {
            if (ExcelErrorHandleMode.orThrow(mode) == ExcelErrorHandleMode.RETURN_CODE) {
                return e.getExcelErrorCode();
            }
        }
        throw e;
    }

    /**
     * 构造配置或编程类错误异常；不受 THROW/RETURN_CODE 影响，调用方应直接抛出返回值。
     */
    public static ExcelException configError(String message) {
        return new ExcelException(message);
    }

    /**
     * 构造带原因的配置或编程类错误异常；不受错误处理模式影响，调用方应直接抛出。
     */
    public static ExcelException configError(String message, Throwable cause) {
        return new ExcelException(message, cause);
    }
}
