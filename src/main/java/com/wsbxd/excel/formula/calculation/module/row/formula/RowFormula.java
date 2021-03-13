package com.wsbxd.excel.formula.calculation.module.row.formula;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;
import com.wsbxd.excel.formula.calculation.module.row.entity.ExcelRow;

import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * description: Row Formula
 *
 * @author chenhaoxuan
 * @date 2019/4/21 9:48
 */
public class RowFormula<T> {

    private final static Pattern RETURN_CELL_PATTERN = Pattern.compile("^[A-Z]+=");

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
    private ExcelEntityProperties properties;

    /**
     * function helper
     */
    private RowFunction<T> rowFunction;

    public RowFormula(String formula, ExcelEntityProperties properties) {
        this.properties = properties;
        //返回单元格
        Matcher matcher = this.properties.getReturnCellPattern().matcher(formula);
        if (matcher.find()) {
            this.returnCell = new ExcelCell(this.properties.getCellStrListByFormula(matcher.group()).get(0), this.properties.getExcelIdTypeEnum());
            this.formula = formula.substring(matcher.end());
        } else {
            this.formula = formula;
        }
        this.rowFunction = new RowFunction<T>(formula, this.properties);
    }

    public String calculate(ExcelRow<T> excelRow) {
        this.rowFunction.functionCalculate(excelRow);
        //公式替换为值
        this.rowFunction.getRowFunctionNodeList().forEach(functionNode -> {
            this.formula = formula.replace(functionNode.getFunction(), functionNode.getValue());
        });
        //解析单元格获取值
        List<String> columnStrList = this.properties.getCellStrListByFormula(this.formula);
        Map<String, String> cellAndValue = excelRow.getCellStrAndValueMap(columnStrList);
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

    public ExcelEntityProperties getProperties() {
        return properties;
    }

    public void setProperties(ExcelEntityProperties properties) {
        this.properties = properties;
    }

    public RowFunction<T> getRowFunction() {
        return rowFunction;
    }

    public void setRowFunction(RowFunction<T> rowFunction) {
        this.rowFunction = rowFunction;
    }

    @Override
    public String toString() {
        return "RowFormula{" +
                "returnCell=" + returnCell +
                ", formula='" + formula + '\'' +
                ", value='" + value + '\'' +
                ", properties=" + properties +
                ", rowFunction=" + rowFunction +
                '}';
    }
}
