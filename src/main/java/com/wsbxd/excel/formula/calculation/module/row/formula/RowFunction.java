package com.wsbxd.excel.formula.calculation.module.row.formula;

import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.function.Function;
import com.wsbxd.excel.formula.calculation.common.function.FunctionImpl;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;
import com.wsbxd.excel.formula.calculation.module.row.entity.ExcelRow;

import java.lang.reflect.Method;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * description: excel Function
 *
 * @author chenhaoxuan
 * @date 2019/8/20 16:50
 */
public class RowFunction<T> {

    /**
     * Excel function matching
     */
    private final static Pattern FUNCTION_PATTERN = Pattern.compile("[A-Z]+\\(");

    /**
     * parenthesis matching
     */
    private final static Pattern PARENTHESIS_PATTERN = Pattern.compile("[(|)]");

    /**
     * name and function map
     */
    public Map<String, Method> nameFunctionMap;

    /**
     * function implement
     */
    private Function functionImpl;

    /**
     * excel data properties
     */
    private ExcelEntityProperties properties;

    /**
     * The function node contained in the current node
     */
    private List<RowFunctionNode<T>> rowFunctionNodeList = new ArrayList<>();

    public Method getMethodByName(String name) {
        return this.nameFunctionMap.get(name);
    }

    public void functionCalculate(ExcelRow<T> excelRow) {
        this.rowFunctionNodeList.forEach(rowFunctionNode -> {
            rowFunctionNode.functionCalculate(this, excelRow);
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
                    addToFunctionNodeList(new RowFunctionNode<>(function, parameters, parenthesisIndex.getKey(), parenthesisIndex.getValue(), this.properties));
                }
            }
        }
    }

    /**
     * add to FunctionNodeList
     *
     * @param rowFunctionNode function Node
     */
    public void addToFunctionNodeList(RowFunctionNode<T> rowFunctionNode) {
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
     * 根据公式获取的下标方法对应Map
     *
     * @param formula 公式
     * @return Map<Index, FunctionName>
     */
    private Map<Integer, String> getIndexFunctionNameMapByFormula(String formula) {
        Map<Integer, String> indexAndFunctionName = new LinkedHashMap<>();
        Matcher functionMatcher = FUNCTION_PATTERN.matcher(formula);
        while (functionMatcher.find()) {
            indexAndFunctionName.put(functionMatcher.end(), functionMatcher.group());
        }
        return indexAndFunctionName;
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

    public RowFunction(String formula, ExcelEntityProperties properties) {
        this(formula, FunctionImpl.class, properties);
    }

    public RowFunction(String formula, Class<? extends Function> functionImplClass, ExcelEntityProperties properties) {
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

    public Function getFunctionImpl() {
        return functionImpl;
    }

    public List<RowFunctionNode<T>> getRowFunctionNodeList() {
        return rowFunctionNodeList;
    }


}
