package com.wsbxd.excel.formula.calculation.common.interfaces;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * description: Excel 计算 接口
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public interface IExcelCalculate {

    /**
     * 批量计算
     *
     * @param formulaList 公式数组
     * @return 计算结果集合
     */
    default List<String> calculateList(String... formulaList) {
        return this.calculateList(Arrays.asList(formulaList));
    }

    /**
     * 批量计算
     *
     * @param formulaList 公式集合
     * @return 计算结果集合
     */
    default List<String> calculateList(List<String> formulaList) {
        List<String> resultList = new ArrayList<>();
        formulaList.forEach(formula -> resultList.add(calculate(formula)));
        return resultList;
    }

    /**
     * 计算公式
     *
     * @param formula 公式
     * @return 计算结果
     */
    String calculate(String formula);

    /**
     * 收敛计算结果到原集合
     */
    void integrationResult();

}
