package com.wsbxd.excel.formula.calculation.module.row;

import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelCalculate;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;
import com.wsbxd.excel.formula.calculation.common.prop.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.module.row.entity.ExcelRow;
import com.wsbxd.excel.formula.calculation.module.row.formula.RowFormula;
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
    private final ExcelEntityProperties properties;

    @Override
    public String calculate(String formula) {
        RowFormula<T> rowFormula = new RowFormula<T>(formula, this.properties);
        String value = rowFormula.calculate(this.excelRow);
        if (null != rowFormula.getReturnCell()) {
            this.excelRow.updateExcelCellValue(rowFormula.getReturnCell());
        }
        return value;
    }

    @Override
    public void integrationResult() {
        this.excelRow.integrationResult();
    }

    public RowCalculate(T t) {
        this.properties = new ExcelEntityProperties(ExcelCalculateTypeEnum.ROW, t.getClass());
        this.excelRow = new ExcelRow<>(t, this.properties);
    }

}
