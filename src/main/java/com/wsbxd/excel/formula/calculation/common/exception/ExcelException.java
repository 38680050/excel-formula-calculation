package com.wsbxd.excel.formula.calculation.common.exception;

/**
 * description: Excel 异常
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/3/27 17:59
 */
public class ExcelException extends RuntimeException {

    /**
     * 定义 Excel 异常
     *
     * @param message 异常信息
     */
    public ExcelException(String message) {
        super(message);
    }

}
