package io.github.kanfengjingren.excelformula.common.calculation.function;

import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.interfaces.IExcelEntity;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel 函数解析与求值入口。
 * <p>
 * 从公式中识别函数名与括号配对，构建嵌套函数树，再从内到外依次求值。
 *
 * @param <T> 数据实体类型
 */
public class ExcelFunction<T> {

    /** 匹配函数名加左括号，如 SUM( */
    private final static Pattern FUNCTION_PATTERN = Pattern.compile("[A-Z]+\\(");

    /** 仅匹配左右圆括号 */
    private final static Pattern PARENTHESIS_PATTERN = Pattern.compile("[()]");

    /** 计算配置（含函数注册表与引用解析） */
    private final ExcelCalculateConfig excelCalculateConfig;

    /** 顶层函数节点列表 */
    private final List<ExcelFunctionNode<T>> excelFunctionNodeList = new ArrayList<>();

    /**
     * 解析公式中的函数并建树。
     *
     * @param formula              已去掉赋值前缀后的公式
     * @param excelCalculateConfig 计算配置
     */
    public ExcelFunction(String formula, ExcelCalculateConfig excelCalculateConfig) {
        this.excelCalculateConfig = excelCalculateConfig;
        Map<Integer, Integer> parenthesisIndexMap = getParenthesisIndexMapByFormula(formula);
        Map<Integer, String> indexFunctionNameMap = getIndexFunctionNameMapByFormula(formula);
        buildFunctionTree(formula, parenthesisIndexMap, indexFunctionNameMap);
    }

    /**
     * 对所有顶层函数节点求值（内部会先递归求子函数）。
     *
     * @param currentSheet 当前页签名，多页签时使用
     * @param excelEntity  单元格数据实体
     */
    public void functionCalculate(String currentSheet, IExcelEntity<T> excelEntity) {
        this.excelFunctionNodeList.forEach(node -> node.functionCalculate(excelEntity, currentSheet));
    }

    /**
     * 按相同左括号下标配对函数名与右括号，截取完整函数串与参数串并插入函数树。
     */
    private void buildFunctionTree(String formula, Map<Integer, Integer> parenthesisIndexMap, Map<Integer, String> indexFunctionNameMap) {
        for (Map.Entry<Integer, Integer> parenthesisIndex : parenthesisIndexMap.entrySet()) {
            for (Map.Entry<Integer, String> indexFunctionName : indexFunctionNameMap.entrySet()) {
                if (parenthesisIndex.getKey().equals(indexFunctionName.getKey())) {
                    String function = formula.substring(indexFunctionName.getKey() - indexFunctionName.getValue().length(), parenthesisIndex.getValue());
                    String parameters = formula.substring(indexFunctionName.getKey(), parenthesisIndex.getValue() - 1);
                    FunctionNodeTree.insert(this.excelFunctionNodeList,
                            new ExcelFunctionNode<>(function, parameters, parenthesisIndex.getKey(), parenthesisIndex.getValue(), this.excelCalculateConfig));
                }
            }
        }
    }

    /**
     * 扫描公式，得到「左括号后下标 → 函数名含左括号」的映射。
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
     * 扫描公式中的括号，按配对规则得到「左括号结束下标 → 右括号结束下标」。
     */
    private Map<Integer, Integer> getParenthesisIndexMapByFormula(String formula) {
        Map<Integer, String> parenthesisMap = new LinkedHashMap<>();
        Matcher parenthesisMatcher = PARENTHESIS_PATTERN.matcher(formula);
        while (parenthesisMatcher.find()) {
            parenthesisMap.put(parenthesisMatcher.end(), parenthesisMatcher.group());
        }
        LinkedHashMap<Integer, Integer> parenthesisIndexMap = new LinkedHashMap<>();
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
            if (left == null || right == null) {
                throw new ExcelException("括号配对失败，公式: " + formula);
            }
            parenthesisIndexMap.put(left, right);
        }
        if (!parenthesisMap.isEmpty()) {
            throw new ExcelException("括号配对失败，公式: " + formula);
        }
        return parenthesisIndexMap;
    }

    /** 获取顶层函数节点列表 */
    public List<ExcelFunctionNode<T>> getExcelFunctionNodeList() {
        return excelFunctionNodeList;
    }
}
