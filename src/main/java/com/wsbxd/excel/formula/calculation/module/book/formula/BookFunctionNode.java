package com.wsbxd.excel.formula.calculation.module.book.formula;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelDataProperties;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;
import com.wsbxd.excel.formula.calculation.module.book.entity.ExcelBook;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * description: Excel 工作簿 方法 节点
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class BookFunctionNode<T> {

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
    private List<BookFunctionNode<T>> bookFunctionNodeList = new ArrayList<>();

    public void functionCalculate(BookFunction bookFunction, ExcelBook<T> excelBook, String currentSheet) {
        //递归计算所有公式
        this.bookFunctionNodeList.forEach(bookFunctionNode -> {
            bookFunctionNode.functionCalculate(bookFunction, excelBook, currentSheet);
            parameters = parameters.replace(bookFunctionNode.function, bookFunctionNode.getValue());
        });
        List<String> valueList = parseParameters(excelBook, currentSheet);
        Method method = bookFunction.getMethodByName(this.getFunctionName());
        try {
            this.setValue((String) method.invoke(bookFunction.getFunctionImpl(), valueList));
        } catch (IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
    }

    private List<String> parseParameters(ExcelBook<T> excelBook, String currentSheet) {
        return new ArrayList<String>() {{
            for (String parameter : ExcelStrUtil.split(getParameters(), ExcelConstant.DOT_CHAT)) {
                if (parameter.contains(ExcelConstant.COLON)) {
                    // Colon parameter processing
                    String[] cellColon = parameter.split(ExcelConstant.COLON);
                    ExcelCell startExcelCell = new ExcelCell(cellColon[0], properties.getExcelIdTypeEnum(), currentSheet);
                    ExcelCell endExcelCell = new ExcelCell(cellColon[1], properties.getExcelIdTypeEnum(), currentSheet);
                    this.addAll(excelBook.getExcelCellValueList(startExcelCell, endExcelCell));
                } else {
                    // Not Colon parameter processing
                    List<String> cellStrList = properties.getCellStrListByFormula(parameter);
                    Map<String, String> cellAndValue = excelBook.getCellStrAndValueMap(cellStrList);
                    this.add(ExcelUtil.functionCalculate(parameter, cellAndValue));
                }
            }
        }};
    }

    /**
     * add to FunctionNodeList
     *
     * @param bookFunctionNode function Node
     */
    public void setFunctionNode(BookFunctionNode<T> bookFunctionNode) {
        if (this.bookFunctionNodeList.isEmpty()) {
            this.bookFunctionNodeList.add(bookFunctionNode);
        } else {
            boolean flag = true;
            for (BookFunctionNode<T> oldBookFunctionNode : this.bookFunctionNodeList) {
                if (oldBookFunctionNode.contain(bookFunctionNode)) {
                    oldBookFunctionNode.setFunctionNode(bookFunctionNode);
                    flag = false;
                }
            }
            if (flag) {
                this.bookFunctionNodeList.add(bookFunctionNode);
            }
        }
    }

    /**
     * Whether to include
     *
     * @param bookFunctionNode function Node
     * @return true/false
     */
    public boolean contain(BookFunctionNode bookFunctionNode) {
        return this.contain(bookFunctionNode.getLeftIndex(), bookFunctionNode.getRightIndex());
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

    public BookFunctionNode(String function, String parameters, Integer leftIndex, Integer rightIndex, ExcelDataProperties properties) {
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

    public ExcelDataProperties getProperties() {
        return properties;
    }

    public void setProperties(ExcelDataProperties properties) {
        this.properties = properties;
    }

    public List<BookFunctionNode<T>> getBookFunctionNodeList() {
        return bookFunctionNodeList;
    }

    public void setBookFunctionNodeList(List<BookFunctionNode<T>> bookFunctionNodeList) {
        this.bookFunctionNodeList = bookFunctionNodeList;
    }

    @Override
    public String toString() {
        return "BookFunctionNode{" +
                "function='" + function + '\'' +
                ", functionName='" + functionName + '\'' +
                ", parameters='" + parameters + '\'' +
                ", value='" + value + '\'' +
                ", leftIndex=" + leftIndex +
                ", rightIndex=" + rightIndex +
                ", properties=" + properties +
                ", bookFunctionNodeList=" + bookFunctionNodeList +
                '}';
    }
}
