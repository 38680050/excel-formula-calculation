package com.wsbxd.excel.formula.calculation.common.function;

import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.interfaces.IFunction;
import com.wsbxd.excel.formula.calculation.common.util.ExcelMathUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.List;
import java.util.Objects;

/**
 * description: Excel 公式 实现
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class FunctionImpl implements IFunction {

    @Override
    public String IF(List<String> valueList) {
        String flag = valueList.get(0);
        //true 或 != 0 都走 valueList.get(1)
        if (ExcelConstant.BIG_CASE_TRUE.equals(flag) || (!ExcelConstant.TRUE_AND_FALSE.contains(flag) && !ExcelConstant.ZERO.equals(flag))) {
            return ExcelStrUtil.isNotBlank(valueList.get(1)) ? valueList.get(1) : "0";
        } else {
            return valueList.size() > 2 ? ExcelStrUtil.isNotBlank(valueList.get(2)) ? valueList.get(2) : "0" : ExcelConstant.BIG_CASE_FALSE;
        }
    }

    @Override
    public String ABS(List<String> valueList) {
        return ExcelStrUtil.getBigDecimal(valueList.stream().findAny().orElse("0")).abs().toString();
    }

    @Override
    public String SUM(List<String> valueList) {
        return ExcelStrUtil.sum(valueList);
    }

    @Override
    public String MAX(List<String> valueList) {
        return ExcelStrUtil.toString(valueList.stream().filter(Objects::nonNull).map(ExcelStrUtil::getBigDecimal).max(BigDecimal::compareTo).orElse(new BigDecimal("0")));
    }

    @Override
    public String MIN(List<String> valueList) {
        return ExcelStrUtil.toString(valueList.stream().filter(Objects::nonNull).map(ExcelStrUtil::getBigDecimal).min(BigDecimal::compareTo).orElse(new BigDecimal("0")));
    }

    @Override
    public String ROUND(List<String> valueList) {
        return ExcelMathUtil.NUMBER__FORMAT.format(new BigDecimal(valueList.get(0)).setScale(Integer.parseInt(valueList.get(1)), RoundingMode.HALF_UP));
    }

}
