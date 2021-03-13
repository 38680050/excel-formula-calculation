package com.wsbxd.excel.formula.calculation.module.sheet;

import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelCalculate;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;
import com.wsbxd.excel.formula.calculation.common.prop.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.module.sheet.entity.ExcelSheet;
import com.wsbxd.excel.formula.calculation.module.sheet.formula.SheetFormula;
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
    private final ExcelEntityProperties properties;

    @Override
    public String calculate(String formula) {
        SheetFormula<T> sheetFormula = new SheetFormula<>(formula, this.properties);
        String value = sheetFormula.calculate(this.excelSheet);
        if (null != sheetFormula.getReturnCell()) {
            this.excelSheet.updateExcelCellValue(sheetFormula.getReturnCell());
        }
        return value;
    }

    @Override
    public void integrationResult() {
        this.excelSheet.integrationResult();
    }

    public SheetCalculate(List<T> excelList, Class<T> tClass) {
        this.properties = new ExcelEntityProperties(ExcelCalculateTypeEnum.SHEET, tClass);
        this.excelSheet = new ExcelSheet<>(excelList, properties);
    }

}
