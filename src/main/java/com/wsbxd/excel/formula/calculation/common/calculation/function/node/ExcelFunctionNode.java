package com.wsbxd.excel.formula.calculation.common.calculation.function.node;

import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;

import java.util.ArrayList;
import java.util.List;

/**
 * description: Excel 函数 节点
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/3/13 16:08
 */
public class ExcelFunctionNode<T> {

    /**
     * 公式字符串
     */
    private String function;

    /**
     * 公式名称
     */
    private String functionName;

    /**
     * 参数
     */
    private String parameters;

    /**
     * 值
     */
    private String value;

    /**
     * 左括号索引
     */
    private Integer leftIndex;

    /**
     * 右括号索引
     */
    private Integer rightIndex;

    /**
     * Excel 实体类 数据属性
     */
    private ExcelEntityProperties properties;

    /**
     * 当前节点中包含的函数节点
     */
    private List<ExcelFunctionNode<T>> excelFunctionNodeList = new ArrayList<>();

}
