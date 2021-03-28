package com.wsbxd.excel.formula.calculation.common.calculation.function;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.function.DefaultFunctionImpl;
import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelEntity;
import com.wsbxd.excel.formula.calculation.common.interfaces.IFunction;
import com.wsbxd.excel.formula.calculation.common.config.ExcelCalculateConfig;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;

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
     * Excel 实体类 数据属性
     */
    private final ExcelCalculateConfig excelCalculateConfig;

    /**
     * 当前函数中包含的函数节点
     */
    private final List<ExcelFunctionNode<T>> excelFunctionNodeList = new ArrayList<>();

    public void functionCalculate(String currentSheet, IExcelEntity<T> excelEntity) {
        this.excelFunctionNodeList.forEach(ExcelFunctionNode -> ExcelFunctionNode.functionCalculate(excelEntity, currentSheet));
    }

    /**
     * 添加函数节点集合
     *
     * @param formula              原公式
     * @param parenthesisIndexMap  括号索引Map，Map<Constant.LEFT_PARENTHESIS.index, Constant.RIGHT_PARENTHESIS.index>
     * @param indexFunctionNameMap 下标和函数名称Map，Map<Index, FunctionName>
     */
    private void setFunctionNodeList(String formula, Map<Integer, Integer> parenthesisIndexMap, Map<Integer, String> indexFunctionNameMap) {
        for (Map.Entry<Integer, Integer> parenthesisIndex : parenthesisIndexMap.entrySet()) {
            for (Map.Entry<Integer, String> indexFunctionName : indexFunctionNameMap.entrySet()) {
                if (parenthesisIndex.getKey().equals(indexFunctionName.getKey())) {
                    String function = formula.substring(indexFunctionName.getKey() - indexFunctionName.getValue().length(), parenthesisIndex.getValue());
                    String parameters = formula.substring(indexFunctionName.getKey(), parenthesisIndex.getValue() - 1);
                    addToFunctionNodeList(new ExcelFunctionNode<>(function, parameters, parenthesisIndex.getKey(), parenthesisIndex.getValue(), this.excelCalculateConfig));
                }
            }
        }
    }

    /**
     * 添加到函数节点集合
     *
     * @param ExcelFunctionNode 函数节点
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

    public ExcelFunction(String formula, ExcelCalculateConfig excelCalculateConfig) {
        this.excelCalculateConfig = excelCalculateConfig;
        Map<Integer, Integer> parenthesisIndexMap = getParenthesisIndexMapByFormula(formula);
        Map<Integer, String> indexFunctionNameMap = getIndexFunctionNameMapByFormula(formula);
        setFunctionNodeList(formula, parenthesisIndexMap, indexFunctionNameMap);
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
        private ExcelCalculateConfig excelCalculateConfig;

        /**
         * 当前节点中包含的函数节点
         */
        private final List<ExcelFunctionNode<T>> excelFunctionNodeList = new ArrayList<>();

        /**
         * 函数计算
         *
         * @param excelEntity  Excel 实体类
         * @param currentSheet 当前页签，只有工作簿计算的时候才会使用这参数
         */
        public void functionCalculate(IExcelEntity<T> excelEntity, String currentSheet) {
            //递归计算所有公式
            this.excelFunctionNodeList.forEach(ExcelFunctionNode -> {
                ExcelFunctionNode.functionCalculate(excelEntity, currentSheet);
                parameters = parameters.replace(ExcelFunctionNode.function, ExcelFunctionNode.getValue());
            });
            List<String> valueList = parseParameters(excelEntity, currentSheet);
            this.setValue(this.excelCalculateConfig.functionCalculate(this.functionName, valueList));
        }

        /**
         * 获取单元格参数值集合
         *
         * @param excelEntity  Excel 实体类
         * @param currentSheet 当前页签，只有工作簿计算的时候才会使用这参数
         * @return 单元格参数值集合
         */
        private List<String> parseParameters(IExcelEntity<T> excelEntity, String currentSheet) {
            List<String> resultList = new ArrayList<>();
            for (String parameter : ExcelStrUtil.split(getParameters(), ExcelConstant.DOT_CHAT)) {
                if (parameter.contains(ExcelConstant.COLON)) {
                    //冒号参数处理
                    String[] cellColon = parameter.split(ExcelConstant.COLON);
                    ExcelCell startExcelCell = new ExcelCell(cellColon[0], this.excelCalculateConfig.getExcelIdTypeEnum(), currentSheet);
                    ExcelCell endExcelCell = new ExcelCell(cellColon[1], this.excelCalculateConfig.getExcelIdTypeEnum(), currentSheet);
                    resultList.addAll(excelEntity.getExcelCellValueList(startExcelCell, endExcelCell));
                } else {
                    //非冒号参数处理
                    List<String> cellStrList = this.excelCalculateConfig.getCellStrListByFormula(parameter);
                    Map<String, String> cellAndValue = excelEntity.getCellStrAndValueMap(cellStrList);
                    resultList.add(ExcelUtil.functionCalculate(parameter, cellAndValue));
                }
            }
            return resultList;
        }

        /**
         * 添加到函数节点集合
         *
         * @param ExcelFunctionNode 函数节点
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
         * 当前函数节点是否包括参数中的函数节点
         *
         * @param ExcelFunctionNode 参数中的函数节点
         * @return 是否包括
         */
        public boolean contain(ExcelFunctionNode<T> ExcelFunctionNode) {
            return this.contain(ExcelFunctionNode.getLeftIndex(), ExcelFunctionNode.getRightIndex());
        }

        /**
         * 当前函数节点的括号下标是否包括参数中函数节点的括号下标
         *
         * @param leftIndex  左括号下标
         * @param rightIndex 右括号下标
         * @return 是否包括
         */
        public boolean contain(Integer leftIndex, Integer rightIndex) {
            return this.leftIndex < leftIndex && this.rightIndex > rightIndex;
        }

        public ExcelFunctionNode(String function, String parameters, Integer leftIndex, Integer rightIndex, ExcelCalculateConfig excelCalculateConfig) {
            this.function = function;
            this.parameters = parameters;
            this.leftIndex = leftIndex;
            this.rightIndex = rightIndex;
            this.functionName = function.substring(0, function.length() - parameters.length() - 2);
            this.excelCalculateConfig = excelCalculateConfig;
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

        public ExcelCalculateConfig getExcelCalculateConfig() {
            return excelCalculateConfig;
        }

        public void setExcelCalculateConfig(ExcelCalculateConfig excelCalculateConfig) {
            this.excelCalculateConfig = excelCalculateConfig;
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
                    ", properties=" + excelCalculateConfig +
                    ", excelFunctionNodeList=" + excelFunctionNodeList +
                    '}';
        }

    }

}
