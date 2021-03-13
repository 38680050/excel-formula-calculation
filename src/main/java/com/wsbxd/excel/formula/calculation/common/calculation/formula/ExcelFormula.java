package com.wsbxd.excel.formula.calculation.common.calculation.formula;

import com.wsbxd.excel.formula.calculation.common.calculation.function.ExcelFunction;
import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;

/**
 * description: Excel 公式
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/3/13 16:55
 */
public class ExcelFormula<T> {

    /**
     * 返回单元格
     */
    private ExcelCell returnCell;

    /**
     * 公式
     */
    private String formula;

    /**
     * 值
     */
    private String value;

    /**
     * Excel 实体类 数据属性
     */
    private ExcelEntityProperties properties;

    /**
     * Excel 函数
     */
    private ExcelFunction<T> excelFunction;

}
