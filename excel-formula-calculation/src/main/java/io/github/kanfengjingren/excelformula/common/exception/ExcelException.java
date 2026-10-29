package io.github.kanfengjingren.excelformula.common.exception;

/**
 * Excel 计算相关运行时异常。
 * <p>
 * 可携带 Excel 错误码（如 #DIV/0!），便于调用方按码分支处理；
 * 也可携带底层 cause，便于排查配置与反射问题。
 */
public class ExcelException extends RuntimeException {

    /** Excel 错误码，可为 null（纯配置/编程错误时） */
    private final String excelErrorCode;

    /**
     * @param message 异常说明
     */
    public ExcelException(String message) {
        super(message);
        this.excelErrorCode = null;
    }

    /**
     * @param message 异常说明
     * @param cause   底层原因
     */
    public ExcelException(String message, Throwable cause) {
        super(message, cause);
        this.excelErrorCode = null;
    }

    /**
     * @param excelErrorCode Excel 错误码，如 #DIV/0!
     * @param message        异常说明
     */
    public ExcelException(String excelErrorCode, String message) {
        super(message != null ? message : excelErrorCode);
        this.excelErrorCode = excelErrorCode;
    }

    /**
     * @param excelErrorCode Excel 错误码
     * @param message        异常说明
     * @param cause          底层原因
     */
    public ExcelException(String excelErrorCode, String message, Throwable cause) {
        super(message != null ? message : excelErrorCode, cause);
        this.excelErrorCode = excelErrorCode;
    }

    /**
     * @return Excel 错误码；非公式语义错误时可能为 null
     */
    public String getExcelErrorCode() {
        return excelErrorCode;
    }
}
