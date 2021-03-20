package com.wsbxd.excel.formula.calculation.common.calculation.function;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.function.FunctionImpl;
import com.wsbxd.excel.formula.calculation.common.interfaces.IFunction;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;
import com.wsbxd.excel.formula.calculation.module.book.entity.ExcelBook;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * description: Excel 函数
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/3/13 16:42
 */
public class ExcelFunction<T> {

    /**
     * Excel 函数匹配
     */
    private final static Pattern FUNCTION_PATTERN = Pattern.compile("[A-Z]+\\(");

    /**
     * 括号匹配
     */
    private final static Pattern PARENTHESIS_PATTERN = Pattern.compile("[(|)]");

    /**
     * 函数和函数名称 Map
     */
    public Map<String, Method> nameFunctionMap;

    /**
     * 函数实现
     */
    private IFunction functionImpl;

    /**
     * Excel 实体类 数据属性
     */
    private final ExcelEntityProperties properties;

    /**
     * 当前函数中包含的函数节点
     */
    private final List<ExcelFunctionNode<T>> excelFunctionNodeList = new ArrayList<>();

    public Method getMethodByName(String name) {
        return this.nameFunctionMap.get(name);
    }

    public void functionCalculate(String currentSheet, ExcelBook<T> excelBook) {
        this.excelFunctionNodeList.forEach(ExcelFunctionNode -> {
            ExcelFunctionNode.functionCalculate(this, excelBook, currentSheet);
        });
    }

    /**
     * set FunctionNodeList
     *
     * @param formula              original formula
     * @param parenthesisIndexMap  Map<Constant.LEFT_PARENTHESIS.index, Constant.RIGHT_PARENTHESIS.index>
     * @param indexFunctionNameMap Map<Index, FunctionName>
     */
    private void setFunctionNodeList(String formula, Map<Integer, Integer> parenthesisIndexMap, Map<Integer, String> indexFunctionNameMap) {
        for (Map.Entry<Integer, Integer> parenthesisIndex : parenthesisIndexMap.entrySet()) {
            for (Map.Entry<Integer, String> indexFunctionName : indexFunctionNameMap.entrySet()) {
                if (parenthesisIndex.getKey().equals(indexFunctionName.getKey())) {
                    String function = formula.substring(indexFunctionName.getKey() - indexFunctionName.getValue().length(), parenthesisIndex.getValue());
                    String parameters = formula.substring(indexFunctionName.getKey(), parenthesisIndex.getValue() - 1);
                    addToFunctionNodeList(new ExcelFunctionNode<>(function, parameters, parenthesisIndex.getKey(), parenthesisIndex.getValue(), this.properties));
                }
            }
        }
    }

    /**
     * add to FunctionNodeList
     *
     * @param ExcelFunctionNode function Node
     */
    public void addToFunctionNodeList(ExcelFunctionNode<T> ExcelFunctionNode) {
        if (this.excelFunctionNodeList.isEmpty()) {
            this.excelFunctionNodeList.add(ExcelFunctionNode);
        } else {
            boolean flag = true;
            for (ExcelFunctionNode<T> oldExcelFunctionNode : this.excelFunctionNodeList) {
                if (oldExcelFunctionNode.contain(ExcelFunctionNode)) {
                    oldExcelFunctionNode.setFunctionNode(ExcelFunctionNode);
                    flag = false;
                }
            }
            if (flag) {
                this.excelFunctionNodeList.add(ExcelFunctionNode);
            }
        }
    }

    public ExcelFunction(String formula, ExcelEntityProperties properties) {
        this(formula, FunctionImpl.class, properties);
    }

    public ExcelFunction(String formula, Class<? extends IFunction> functionImplClass, ExcelEntityProperties properties) {
        this.properties = properties;
        Map<Integer, Integer> parenthesisIndexMap = getParenthesisIndexMapByFormula(formula);
        Map<Integer, String> indexFunctionNameMap = getIndexFunctionNameMapByFormula(formula);
        setFunctionNodeList(formula, parenthesisIndexMap, indexFunctionNameMap);
        try {
            this.functionImpl = functionImplClass.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        this.nameFunctionMap = new HashMap<>();
        for (Method method : functionImplClass.getDeclaredMethods()) {
            this.nameFunctionMap.put(method.getName(), method);
        }
    }

    /**
     * 根据公式获取的下标方法对应Map
     *
     * @param formula 公式
     * @return Map<Index, FunctionName>
     */
    private Map<Integer, String> getIndexFunctionNameMapByFormula(String formula) {
        Map<Integer, String> indexAndFunctionNameMap = new LinkedHashMap<>();
        Matcher functionMatcher = FUNCTION_PATTERN.matcher(formula);
        while (functionMatcher.find()) {
            indexAndFunctionNameMap.put(functionMatcher.end(), functionMatcher.group());
        }
        return indexAndFunctionNameMap;
    }

    /**
     * 根据公式获取对应的括号下标
     *
     * @param formula formula
     * @return Map<Constant.LEFT_PARENTHESIS.index, Constant.RIGHT_PARENTHESIS.index>
     */
    private Map<Integer, Integer> getParenthesisIndexMapByFormula(String formula) {
        Map<Integer, String> parenthesisMap = new LinkedHashMap<>();
        Matcher parenthesisMatcher = PARENTHESIS_PATTERN.matcher(formula);
        while (parenthesisMatcher.find()) {
            parenthesisMap.put(parenthesisMatcher.end(), parenthesisMatcher.group());
        }
        LinkedHashMap<Integer, Integer> parenthesisIndexMap = new LinkedHashMap<>();
        //总数除以2为对应括号数量
        int parenthesisNum = parenthesisMap.size() / 2;
        for (int i = 0; i < parenthesisNum; i++) {
            int count = 0;
            Integer left = null;
            Integer right = null;
            Iterator<Map.Entry<Integer, String>> parenthesisIterator = parenthesisMap.entrySet().iterator();
            while (parenthesisIterator.hasNext()) {
                Map.Entry<Integer, String> parenthesis = parenthesisIterator.next();
                if (null == left && ExcelConstant.LEFT_PARENTHESIS.equals(parenthesis.getValue())) {
                    left = parenthesis.getKey();
                    parenthesisIterator.remove();
                } else if (ExcelConstant.LEFT_PARENTHESIS.equals(parenthesis.getValue())) {
                    count++;
                } else if (0 == count && ExcelConstant.RIGHT_PARENTHESIS.equals(parenthesis.getValue())) {
                    right = parenthesis.getKey();
                    parenthesisIterator.remove();
                    break;
                } else {
                    count--;
                }
            }
            parenthesisIndexMap.put(left, right);
        }
        return parenthesisIndexMap;
    }

    public IFunction getFunctionImpl() {
        return functionImpl;
    }

    public List<ExcelFunctionNode<T>> getExcelFunctionNodeList() {
        return excelFunctionNodeList;
    }

    /**
     * description: Excel 函数 节点
     *
     * @author chenhaoxuan
     * @version 1.0
     * @date 2021/3/13 16:08
     */
    public static class ExcelFunctionNode<T> {

        /**
         * 公式字符串
         */
        private final String function;

        /**
         * 公式名称
         */
        private final String functionName;

        /**
         * 参数
         */
        private String parameters;

        /**
         * 值
         */
        private String value;

        /**
         * 左括号索引
         */
        private final Integer leftIndex;

        /**
         * 右括号索引
         */
        private final Integer rightIndex;

        /**
         * Excel 实体类 数据属性
         */
        private ExcelEntityProperties properties;

        /**
         * 当前节点中包含的函数节点
         */
        private final List<ExcelFunctionNode<T>> excelFunctionNodeList = new ArrayList<>();

        public void functionCalculate(ExcelFunction<T> ExcelFunction, ExcelBook<T> excelBook, String currentSheet) {
            //递归计算所有公式
            this.excelFunctionNodeList.forEach(ExcelFunctionNode -> {
                ExcelFunctionNode.functionCalculate(ExcelFunction, excelBook, currentSheet);
                parameters = parameters.replace(ExcelFunctionNode.function, ExcelFunctionNode.getValue());
            });
            List<String> valueList = parseParameters(excelBook, currentSheet);
            Method method = ExcelFunction.getMethodByName(this.getFunctionName());
            try {
                this.setValue((String) method.invoke(ExcelFunction.getFunctionImpl(), valueList));
            } catch (IllegalAccessException | InvocationTargetException e) {
                e.printStackTrace();
            }
        }

        private List<String> parseParameters(ExcelBook<T> excelBook, String currentSheet) {
            List<String> resultList = new ArrayList<>();
            for (String parameter : ExcelStrUtil.split(getParameters(), ExcelConstant.DOT_CHAT)) {
                if (parameter.contains(ExcelConstant.COLON)) {
                    // Colon parameter processing
                    String[] cellColon = parameter.split(ExcelConstant.COLON);
                    ExcelCell startExcelCell = new ExcelCell(cellColon[0], properties.getExcelIdTypeEnum(), currentSheet);
                    ExcelCell endExcelCell = new ExcelCell(cellColon[1], properties.getExcelIdTypeEnum(), currentSheet);
                    resultList.addAll(excelBook.getExcelCellValueList(startExcelCell, endExcelCell));
                } else {
                    // Not Colon parameter processing
                    List<String> cellStrList = properties.getCellStrListByFormula(parameter);
                    Map<String, String> cellAndValue = excelBook.getCellStrAndValueMap(cellStrList);
                    resultList.add(ExcelUtil.functionCalculate(parameter, cellAndValue));
                }
            }
            return resultList;
        }

        /**
         * add to FunctionNodeList
         *
         * @param ExcelFunctionNode function Node
         */
        public void setFunctionNode(ExcelFunctionNode<T> ExcelFunctionNode) {
            if (this.excelFunctionNodeList.isEmpty()) {
                this.excelFunctionNodeList.add(ExcelFunctionNode);
            } else {
                boolean flag = true;
                for (ExcelFunctionNode<T> oldExcelFunctionNode : this.excelFunctionNodeList) {
                    if (oldExcelFunctionNode.contain(ExcelFunctionNode)) {
                        oldExcelFunctionNode.setFunctionNode(ExcelFunctionNode);
                        flag = false;
                    }
                }
                if (flag) {
                    this.excelFunctionNodeList.add(ExcelFunctionNode);
                }
            }
        }

        /**
         * Whether to include
         *
         * @param ExcelFunctionNode function Node
         * @return true/false
         */
        public boolean contain(ExcelFunctionNode<T> ExcelFunctionNode) {
            return this.contain(ExcelFunctionNode.getLeftIndex(), ExcelFunctionNode.getRightIndex());
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

        public ExcelFunctionNode(String function, String parameters, Integer leftIndex, Integer rightIndex, ExcelEntityProperties properties) {
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

        public String getFunctionName() {
            return functionName;
        }

        public String getParameters() {
            return parameters;
        }

        public String getValue() {
            return value;
        }

        private void setValue(String value) {
            this.value = value;
        }

        public Integer getLeftIndex() {
            return leftIndex;
        }

        public Integer getRightIndex() {
            return rightIndex;
        }

        public ExcelEntityProperties getProperties() {
            return properties;
        }

        public void setProperties(ExcelEntityProperties properties) {
            this.properties = properties;
        }

        public List<ExcelFunctionNode<T>> getExcelFunctionNodeList() {
            return excelFunctionNodeList;
        }

        @Override
        public String toString() {
            return "ExcelFunctionNode{" +
                    "function='" + function + '\'' +
                    ", functionName='" + functionName + '\'' +
                    ", parameters='" + parameters + '\'' +
                    ", value='" + value + '\'' +
                    ", leftIndex=" + leftIndex +
                    ", rightIndex=" + rightIndex +
                    ", properties=" + properties +
                    ", excelFunctionNodeList=" + excelFunctionNodeList +
                    '}';
        }

    }

}
