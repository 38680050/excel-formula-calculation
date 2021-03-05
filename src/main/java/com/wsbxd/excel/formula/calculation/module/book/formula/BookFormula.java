package com.wsbxd.excel.formula.calculation.module.book.formula;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelDataProperties;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;
import com.wsbxd.excel.formula.calculation.module.book.entity.ExcelBook;

import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

/**
 * description: Excel 工作簿 公式
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class BookFormula<T> {

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
     * book Function
     */
    private BookFunction<T> bookFunction;

    public BookFormula(String currentSheet, String formula, ExcelDataProperties properties) {
        this.properties = properties;
        //返回单元格
        Matcher matcher = this.properties.getReturnCellPattern().matcher(formula);
        if (matcher.find()) {
            this.returnCell = new ExcelCell(this.properties.getCellStrListByFormula(matcher.group()).get(0), this.properties.getExcelIdTypeEnum(), currentSheet);
            this.formula = formula.substring(matcher.end());
        } else {
            this.formula = formula;
        }
        this.bookFunction = new BookFunction<>(formula, this.properties);
    }

    public String calculate(String currentSheet, ExcelBook<T> excelBook) {
        this.bookFunction.functionCalculate(currentSheet,excelBook);
        //公式替换为值
        this.bookFunction.getBookFunctionNodeList().forEach(functionNode -> {
            this.formula = formula.replace(functionNode.getFunction(), functionNode.getValue());
        });
        //解析单元格获取值
        List<String> cellStrList = this.properties.getCellStrListByFormula(this.formula);
        Map<String, String> cellAndValue = excelBook.getCellStrAndValueMap(cellStrList);
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

    public BookFunction<T> getBookFunction() {
        return bookFunction;
    }

    public void setBookFunction(BookFunction<T> bookFunction) {
        this.bookFunction = bookFunction;
    }

    @Override
    public String toString() {
        return "BookFormula{" +
                "returnCell=" + returnCell +
                ", formula='" + formula + '\'' +
                ", value='" + value + '\'' +
                ", properties=" + properties +
                ", bookFunction=" + bookFunction +
                '}';
    }
}
