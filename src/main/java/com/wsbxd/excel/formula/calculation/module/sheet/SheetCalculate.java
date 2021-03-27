package com.wsbxd.excel.formula.calculation.module.sheet;

import com.wsbxd.excel.formula.calculation.common.calculation.formula.ExcelFormula;
import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelCalculate;
import com.wsbxd.excel.formula.calculation.common.config.ExcelCalculateConfig;
import com.wsbxd.excel.formula.calculation.common.config.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.module.sheet.entity.ExcelSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * description: Sheet Calculate
 *
 * @author chenhaoxuan
 * @date 2019/8/28
 */
public class SheetCalculate<T> implements IExcelCalculate {

    private final static Logger logger = LoggerFactory.getLogger(SheetCalculate.class);

    /**
     * excel cells data
     */
    private final ExcelSheet<T> excelSheet;

    /**
     * excel data properties
     */
    private final ExcelCalculateConfig excelCalculateConfig;

    @Override
    public String calculate(String formula) {
        ExcelFormula<T> sheetFormula = new ExcelFormula<>(null, formula, this.excelCalculateConfig);
        String value = sheetFormula.calculate(null, this.excelSheet);
        if (null != sheetFormula.getReturnCell()) {
            this.excelSheet.updateExcelCellValue(sheetFormula.getReturnCell());
        }
        return value;
    }

    @Override
    public void integrationResult() {
        this.excelSheet.integrationResult();
    }

    public SheetCalculate(List<T> excelList, ExcelCalculateConfig excelCalculateConfig) {
        this.excelCalculateConfig = excelCalculateConfig;
        this.excelCalculateConfig.setCalculateType(ExcelCalculateTypeEnum.SHEET);
        this.excelSheet = new ExcelSheet<>(excelList, this.excelCalculateConfig);
    }

}
