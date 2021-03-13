package com.wsbxd.excel.formula.calculation.common.calculation.function;

import com.wsbxd.excel.formula.calculation.common.calculation.function.node.ExcelFunctionNode;
import com.wsbxd.excel.formula.calculation.common.function.Function;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * description: Excel 函数
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/3/13 16:42
 */
public class ExcelFunction<T> {

    /**
     * Excel 函数匹配
     */
    private final static Pattern FUNCTION_PATTERN = Pattern.compile("[A-Z]+\\(");

    /**
     * 括号匹配
     */
    private final static Pattern PARENTHESIS_PATTERN = Pattern.compile("[(|)]");

    /**
     * 函数和函数名称 Map
     */
    public Map<String, Method> nameFunctionMap;

    /**
     * 函数实现
     */
    private Function functionImpl;

    /**
     * Excel 实体类 数据属性
     */
    private ExcelEntityProperties properties;

    /**
     * 当前函数中包含的函数节点
     */
    private final List<ExcelFunctionNode<T>> excelFunctionNodeList = new ArrayList<>();

}
