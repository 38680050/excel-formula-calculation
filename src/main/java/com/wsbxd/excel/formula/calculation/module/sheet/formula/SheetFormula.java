package com.wsbxd.excel.formula.calculation.module.sheet.formula;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelDataProperties;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;
import com.wsbxd.excel.formula.calculation.module.sheet.entity.ExcelSheet;

import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

/**
 * description: sheet formula
 *
 * @author chenhaoxuan
 * @date 2019/4/21 9:48
 */
public class SheetFormula<T> {

    private ExcelCell returnCell;

    /**
     * current formula
     */
    private String formula;

    /**
     * Calculated Value
     */
    private String value;

    /**
     * excel data properties
     */
    private ExcelDataProperties properties;

    /**
     * function helper
     */
    private SheetFunction<T> sheetFunction;

    public SheetFormula(String formula, ExcelDataProperties properties) {
        this.properties = properties;
        //返回单元格
        Matcher matcher = this.properties.getReturnCellPattern().matcher(formula);
        if (matcher.find()) {
            this.returnCell = new ExcelCell(this.properties.getCellStrListByFormula(matcher.group()).get(0), this.properties.getExcelIdTypeEnum());
            this.formula = formula.substring(matcher.end());
        } else {
            this.formula = formula;
        }
        this.sheetFunction = new SheetFunction<T>(formula, this.properties);
    }

    public String calculate(ExcelSheet<T> excelSheet) {
        this.sheetFunction.functionCalculate(excelSheet);
        //公式替换为值
        this.sheetFunction.getSheetFunctionNodeList().forEach(functionNode -> {
            this.formula = formula.replace(functionNode.getFunction(), functionNode.getValue());
        });
        //解析单元格获取值
        List<String> cellStrList = this.properties.getCellStrListByFormula(this.formula);
        Map<String, String> cellAndValue = excelSheet.getCellStrAndValueMap(cellStrList);
        String value = ExcelUtil.arithmeticCalculate(this.formula, cellAndValue);
        //如果有返回单元格则添加至返回单元格
        if (null != this.returnCell) {
            this.returnCell.setOriginalValue(value);
        }
        return value;
    }

    public ExcelCell getReturnCell() {
        return returnCell;
    }

    public void setReturnCell(ExcelCell returnCell) {
        this.returnCell = returnCell;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public ExcelDataProperties getProperties() {
        return properties;
    }

    public void setProperties(ExcelDataProperties properties) {
        this.properties = properties;
    }

    public SheetFunction<T> getSheetFunction() {
        return sheetFunction;
    }

    public void setSheetFunction(SheetFunction<T> sheetFunction) {
        this.sheetFunction = sheetFunction;
    }

    @Override
    public String toString() {
        return "SheetFormula{" +
                "returnCell=" + returnCell +
                ", formula='" + formula + '\'' +
                ", value='" + value + '\'' +
                ", properties=" + properties +
                ", sheetFunction=" + sheetFunction +
                '}';
    }
}
