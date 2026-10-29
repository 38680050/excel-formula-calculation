package io.github.kanfengjingren.excelformula.common.function;

import java.util.List;

/**
 * 单个 Excel 函数的处理器；可用于 Lambda 注册。
 */
@FunctionalInterface
public interface ExcelFunctionHandler {

    /**
     * @param args 已展开求值后的参数字符串列表
     * @return 函数结果字符串
     */
    String apply(List<String> args);
}
