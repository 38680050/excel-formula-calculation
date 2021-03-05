package com.wsbxd.excel.formula.calculation.module.row.formula;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelDataProperties;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;
import com.wsbxd.excel.formula.calculation.module.row.entity.ExcelRow;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * description: excel Function Node
 *
 * @author chenhaoxuan
 * @date 2019/8/20 21:54
 */
public class RowFunctionNode<T> {

    /**
     * function
     */
    private String function;

    /**
     * function name
     */
    private String functionName;

    /**
     * parameters
     */
    private String parameters;

    /**
     * value
     */
    private String value;

    /**
     * left parenthesis index
     */
    private Integer leftIndex;

    /**
     * right parenthesis index
     */
    private Integer rightIndex;

    /**
     * excel data properties
     */
    private ExcelDataProperties properties;

    /**
     * The function node contained in the current node
     */
    private List<RowFunctionNode<T>> rowFunctionNodeList = new ArrayList<>();

    public void functionCalculate(RowFunction rowFunction, ExcelRow<T> excelRow) {
        //递归计算所有公式
        this.rowFunctionNodeList.forEach(rowFunctionNode -> {
            rowFunctionNode.functionCalculate(rowFunction, excelRow);
            parameters = parameters.replace(rowFunctionNode.function, rowFunctionNode.getValue());
        });
        List<String> valueList = parseParameters(excelRow);
        Method method = rowFunction.getMethodByName(this.getFunctionName());
        try {
            this.setValue((String) method.invoke(rowFunction.getFunctionImpl(), valueList));
        } catch (IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
    }

    private List<String> parseParameters(ExcelRow<T> excelRow) {
        return new ArrayList<String>() {{
            for (String parameter : ExcelStrUtil.split(getParameters(), ExcelConstant.DOT_CHAT)) {
                if (parameter.contains(ExcelConstant.COLON)) {
                    // Colon parameter processing
                    String[] cellColon = parameter.split(ExcelConstant.COLON);
                    ExcelCell startExcelCell = new ExcelCell(cellColon[0], properties.getExcelIdTypeEnum());
                    ExcelCell endExcelCell = new ExcelCell(cellColon[1], properties.getExcelIdTypeEnum());
                    this.addAll(excelRow.getExcelCellValueList(startExcelCell, endExcelCell));
                } else {
                    // Not Colon parameter processing
                    List<String> cellStrList = properties.getCellStrListByFormula(parameter);
                    Map<String, String> columnAndValue = excelRow.getColumnStrAndValueMap(cellStrList);
                    this.add(ExcelUtil.functionCalculate(parameter, columnAndValue));
                }
            }
        }};
    }

    /**
     * add to FunctionNodeList
     *
     * @param rowFunctionNode function Node
     */
    public void setFunctionNode(RowFunctionNode<T> rowFunctionNode) {
        if (this.rowFunctionNodeList.isEmpty()) {
            this.rowFunctionNodeList.add(rowFunctionNode);
        } else {
            boolean flag = true;
            for (RowFunctionNode<T> oldRowFunctionNode : this.rowFunctionNodeList) {
                if (oldRowFunctionNode.contain(rowFunctionNode)) {
                    oldRowFunctionNode.setFunctionNode(rowFunctionNode);
                    flag = false;
                }
            }
            if (flag) {
                this.rowFunctionNodeList.add(rowFunctionNode);
            }
        }
    }

    /**
     * Whether to include
     *
     * @param rowFunctionNode function Node
     * @return true/false
     */
    public boolean contain(RowFunctionNode rowFunctionNode) {
        return this.contain(rowFunctionNode.getLeftIndex(), rowFunctionNode.getRightIndex());
    }

    /**
     * Whether to include
     *
     * @param leftIndex  left parenthesis index
     * @param rightIndex right parenthesis index
     * @return true/false
     */
    public boolean contain(Integer leftIndex, Integer rightIndex) {
        return this.leftIndex < leftIndex && this.rightIndex > rightIndex;
    }

    /**
     * is there a value?
     *
     * @return true/false
     */
    public boolean hasValue() {
        return value != null;
    }

    public RowFunctionNode(String function, String parameters, Integer leftIndex, Integer rightIndex, ExcelDataProperties properties) {
        this.properties = properties;
        this.function = function;
        this.parameters = parameters;
        this.leftIndex = leftIndex;
        this.rightIndex = rightIndex;
        this.functionName = function.substring(0, function.length() - parameters.length() - 2);
    }

    public String getFunction() {
        return function;
    }

    public void setFunction(String function) {
        this.function = function;
    }

    public String getFunctionName() {
        return functionName;
    }

    public void setFunctionName(String functionName) {
        this.functionName = functionName;
    }

    public String getParameters() {
        return parameters;
    }

    public void setParameters(String parameters) {
        this.parameters = parameters;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public Integer getLeftIndex() {
        return leftIndex;
    }

    public void setLeftIndex(Integer leftIndex) {
        this.leftIndex = leftIndex;
    }

    public Integer getRightIndex() {
        return rightIndex;
    }

    public void setRightIndex(Integer rightIndex) {
        this.rightIndex = rightIndex;
    }

    public ExcelDataProperties getProperties() {
        return properties;
    }

    public void setProperties(ExcelDataProperties properties) {
        this.properties = properties;
    }

    public List<RowFunctionNode<T>> getRowFunctionNodeList() {
        return rowFunctionNodeList;
    }

    public void setRowFunctionNodeList(List<RowFunctionNode<T>> rowFunctionNodeList) {
        this.rowFunctionNodeList = rowFunctionNodeList;
    }

    @Override
    public String toString() {
        return "RowFunctionNode{" +
                "function='" + function + '\'' +
                ", functionName='" + functionName + '\'' +
                ", parameters='" + parameters + '\'' +
                ", value='" + value + '\'' +
                ", leftIndex=" + leftIndex +
                ", rightIndex=" + rightIndex +
                ", properties=" + properties +
                ", rowFunctionNodeList=" + rowFunctionNodeList +
                '}';
    }
}
