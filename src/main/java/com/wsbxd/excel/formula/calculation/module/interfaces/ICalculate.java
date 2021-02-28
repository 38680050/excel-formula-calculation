package com.wsbxd.excel.formula.calculation.module.interfaces;

import java.util.List;

/**
 * description: Excel 计算 接口
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public interface ICalculate {

    /**
     * 计算公式但不修改excelCells的值
     *
     * @param formula 公式
     * @return 计算结果
     */
    String calculateNotChangeValue(String formula);

    /**
     * 计算公式但不修改excelCells的值
     *
     * @param formulaList 公式集合
     */
    void calculateListChangeValue(String... formulaList);

    /**
     * 计算公式但不修改excelCells的值
     *
     * @param formulaList 公式集合
     */
    void calculateListChangeValue(List<String> formulaList);

    /**
     * 计算公式但不修改excelCells的值
     *
     * @param formula 公式
     * @return 计算结果
     */
    String calculateChangeValue(String formula);

}
