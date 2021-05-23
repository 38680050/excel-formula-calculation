package com.wsbxd.excel.formula.calculation.module.row;

import com.wsbxd.excel.formula.calculation.common.calculation.formula.ExcelFormula;
import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelCalculate;
import com.wsbxd.excel.formula.calculation.common.config.ExcelCalculateConfig;
import com.wsbxd.excel.formula.calculation.common.config.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.module.row.entity.ExcelRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * description: Row Calculate
 *
 * @author chenhaoxuan
 * @date 2019/8/28
 */
public class RowCalculate<T> implements IExcelCalculate {

    private final static Logger logger = LoggerFactory.getLogger(RowCalculate.class);

    /**
     * excel row
     */
    private final ExcelRow<T> excelRow;

    /**
     * excel data properties
     */
    private final ExcelCalculateConfig excelCalculateConfig;

    @Override
    public String calculate(String formula) {
        ExcelFormula<T> rowFormula = new ExcelFormula<T>(null, formula, this.excelCalculateConfig);
        String value = rowFormula.calculate(null, this.excelRow);
        if (null != rowFormula.getReturnCell()) {
            value = this.excelRow.updateExcelCellValue(rowFormula.getReturnCell());
        }
        return value;
    }

    @Override
    public void integrationResult() {
        this.excelRow.integrationResult();
    }

    public RowCalculate(T t, ExcelCalculateConfig excelCalculateConfig) {
        this.excelCalculateConfig = excelCalculateConfig;
        this.excelCalculateConfig.setCalculateType(ExcelCalculateTypeEnum.ROW);
        this.excelRow = new ExcelRow<>(t, this.excelCalculateConfig);
    }

}
