package com.wsbxd.excel.formula.calculation.module.sheet.formula;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;
import com.wsbxd.excel.formula.calculation.module.sheet.entity.ExcelSheet;

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
public class SheetFunctionNode<T> {

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
    private ExcelEntityProperties properties;

    /**
     * The function node contained in the current node
     */
    private List<SheetFunctionNode<T>> sheetFunctionNodeList = new ArrayList<>();

    public void functionCalculate(SheetFunction<T> sheetFunction, ExcelSheet<T> excelSheet) {
        //递归计算所有公式
        this.sheetFunctionNodeList.forEach(sheetFunctionNode -> {
            sheetFunctionNode.functionCalculate(sheetFunction, excelSheet);
            parameters = parameters.replace(sheetFunctionNode.function, sheetFunctionNode.getValue());
        });
        List<String> valueList = parseParameters(excelSheet);
        Method method = sheetFunction.getMethodByName(this.getFunctionName());
        try {
            this.setValue((String) method.invoke(sheetFunction.getFunctionImpl(), valueList));
        } catch (IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
    }

    private List<String> parseParameters(ExcelSheet<T> excelSheet) {
        List<String> resultList = new ArrayList<>();
        for (String parameter : ExcelStrUtil.split(getParameters(), ExcelConstant.DOT_CHAT)) {
            if (parameter.contains(ExcelConstant.COLON)) {
                // Colon parameter processing
                String[] cellColon = parameter.split(ExcelConstant.COLON);
                ExcelCell startExcelCell = new ExcelCell(cellColon[0], properties.getExcelIdTypeEnum());
                ExcelCell endExcelCell = new ExcelCell(cellColon[1], properties.getExcelIdTypeEnum());
                resultList.addAll(excelSheet.getExcelCellValueList(startExcelCell, endExcelCell));
            } else {
                // Not Colon parameter processing
                List<String> cellStrList = properties.getCellStrListByFormula(parameter);
                Map<String, String> cellAndValue = excelSheet.getCellStrAndValueMap(cellStrList);
                resultList.add(ExcelUtil.functionCalculate(parameter, cellAndValue));
            }
        }
        return resultList;
    }

    /**
     * add to FunctionNodeList
     *
     * @param sheetFunctionNode function Node
     */
    public void setFunctionNode(SheetFunctionNode<T> sheetFunctionNode) {
        if (this.sheetFunctionNodeList.isEmpty()) {
            this.sheetFunctionNodeList.add(sheetFunctionNode);
        } else {
            boolean flag = true;
            for (SheetFunctionNode<T> oldSheetFunctionNode : this.sheetFunctionNodeList) {
                if (oldSheetFunctionNode.contain(sheetFunctionNode)) {
                    oldSheetFunctionNode.setFunctionNode(sheetFunctionNode);
                    flag = false;
                }
            }
            if (flag) {
                this.sheetFunctionNodeList.add(sheetFunctionNode);
            }
        }
    }

    /**
     * Whether to include
     *
     * @param sheetFunctionNode function Node
     * @return true/false
     */
    public boolean contain(SheetFunctionNode sheetFunctionNode) {
        return this.contain(sheetFunctionNode.getLeftIndex(), sheetFunctionNode.getRightIndex());
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

    public SheetFunctionNode(String function, String parameters, Integer leftIndex, Integer rightIndex, ExcelEntityProperties properties) {
        this.function = function;
        this.parameters = parameters;
        this.leftIndex = leftIndex;
        this.rightIndex = rightIndex;
        this.functionName = function.substring(0, function.length() - parameters.length() - 2);
        this.properties = properties;
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

    public ExcelEntityProperties getProperties() {
        return properties;
    }

    public void setProperties(ExcelEntityProperties properties) {
        this.properties = properties;
    }

    public List<SheetFunctionNode<T>> getSheetFunctionNodeList() {
        return sheetFunctionNodeList;
    }

    public void setSheetFunctionNodeList(List<SheetFunctionNode<T>> sheetFunctionNodeList) {
        this.sheetFunctionNodeList = sheetFunctionNodeList;
    }

    @Override
    public String toString() {
        return "SheetFunctionNode{" +
                "function='" + function + '\'' +
                ", functionName='" + functionName + '\'' +
                ", parameters='" + parameters + '\'' +
                ", value='" + value + '\'' +
                ", leftIndex=" + leftIndex +
                ", rightIndex=" + rightIndex +
                ", properties=" + properties +
                ", sheetFunctionNodeList=" + sheetFunctionNodeList +
                '}';
    }
}
