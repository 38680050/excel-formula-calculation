package com.wsbxd.excel.formula.calculation.module.interfaces.abstracts;

import com.wsbxd.excel.formula.calculation.module.interfaces.ICalculate;

import java.util.Arrays;
import java.util.List;

/**
 * description: Excel 计算 抽象
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public abstract class AbstractsCalculate implements ICalculate {

    @Override
    public void calculateListChangeValue(String... formulaList) {
        this.calculateListChangeValue(Arrays.asList(formulaList));
    }

    @Override
    public void calculateListChangeValue(List<String> formulaList) {
        formulaList.forEach(this::calculateChangeValue);
    }

}
