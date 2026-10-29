package io.github.kanfengjingren.excelformula.common.calculation.function;

import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCell;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelErrorUtil;
import io.github.kanfengjingren.excelformula.common.interfaces.IExcelEntity;
import io.github.kanfengjingren.excelformula.common.util.ExcelStrUtil;
import io.github.kanfengjingren.excelformula.common.util.ExcelUtil;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 单个 Excel 函数节点。
 * <p>
 * 保存函数原文、函数名、参数串、括号下标及嵌套子函数；
 * 求值时先算子函数并替换进参数，再解析参数（区域引用或表达式），最后调用已注册函数。
 * 子结果或参数为错误字面量时按配置模式短路；禁止将 null 直接 replace 进参数串。
 *
 * @param <T> 数据实体类型
 */
public class ExcelFunctionNode<T> {

    /** 完整函数原文，如 SUM(A1:B2) */
    private final String function;

    /** 函数名，如 SUM */
    private final String functionName;

    /** 括号内参数原文，子函数求值后会被替换为数值 */
    private String parameters;

    /** 本节点求值结果 */
    private String value;

    /** 左括号在公式中的结束下标 */
    private final Integer leftIndex;

    /** 右括号在公式中的结束下标 */
    private final Integer rightIndex;

    /** 计算配置 */
    private final ExcelCalculateConfig excelCalculateConfig;

    /** 嵌套在本函数参数中的子函数节点 */
    private final List<ExcelFunctionNode<T>> excelFunctionNodeList = new ArrayList<>();

    /**
     * @param function             完整函数串
     * @param parameters           括号内参数串
     * @param leftIndex            左括号结束下标
     * @param rightIndex           右括号结束下标
     * @param excelCalculateConfig 计算配置
     */
    public ExcelFunctionNode(String function, String parameters, Integer leftIndex, Integer rightIndex, ExcelCalculateConfig excelCalculateConfig) {
        this.function = function;
        this.parameters = parameters;
        this.leftIndex = leftIndex;
        this.rightIndex = rightIndex;
        this.functionName = function.substring(0, function.length() - parameters.length() - 2);
        this.excelCalculateConfig = excelCalculateConfig;
    }

    /**
     * 先递归计算子函数并替换参数中的子函数原文，再解析参数并调用注册函数得到本节点值。
     */
    public void functionCalculate(IExcelEntity<T> excelEntity, String currentSheet) {
        ExcelErrorHandleMode mode = this.excelCalculateConfig.getErrorHandleMode();
        for (ExcelFunctionNode<T> node : this.excelFunctionNodeList) {
            node.functionCalculate(excelEntity, currentSheet);
            String childValue = normalizeReplaceValue(node.getValue());
            if (ExcelErrorUtil.isError(childValue)) {
                this.value = ExcelErrorUtil.fail(mode, childValue, "子函数错误: " + childValue);
                return;
            }
            parameters = parameters.replace(node.function, childValue);
        }
        List<String> valueList = parseParameters(excelEntity, currentSheet);
        for (String paramValue : valueList) {
            if (ExcelErrorUtil.isError(paramValue)) {
                this.value = ExcelErrorUtil.fail(mode, paramValue, "函数参数错误: " + paramValue);
                return;
            }
        }
        this.value = this.excelCalculateConfig.functionCalculate(this.functionName, valueList);
    }

    /**
     * null 按空单元格语义规范化为空串，避免 String.replace 的 NPE。
     */
    private static String normalizeReplaceValue(String raw) {
        return raw == null ? "" : raw;
    }

    /**
     * 按逗号拆分参数：含冒号的按区域展开为多个单元格值；否则对参数做单元格替换与算术求值。
     */
    private List<String> parseParameters(IExcelEntity<T> excelEntity, String currentSheet) {
        ExcelErrorHandleMode mode = this.excelCalculateConfig.getErrorHandleMode();
        List<String> resultList = new ArrayList<>();
        for (String parameter : ExcelStrUtil.split(getParameters(), ExcelConstant.DOT_CHAT)) {
            if (parameter.contains(ExcelConstant.COLON)) {
                String[] cellColon = parameter.split(ExcelConstant.COLON);
                ExcelCell startExcelCell = new ExcelCell(cellColon[0], this.excelCalculateConfig.getRowIdCodec(), currentSheet);
                ExcelCell endExcelCell = new ExcelCell(cellColon[1], this.excelCalculateConfig.getRowIdCodec(), currentSheet);
                List<String> rangeValues = excelEntity.getExcelCellValueList(startExcelCell, endExcelCell);
                for (String rangeValue : rangeValues) {
                    if (ExcelErrorUtil.isError(rangeValue)) {
                        resultList.clear();
                        resultList.add(rangeValue);
                        return resultList;
                    }
                }
                resultList.addAll(rangeValues);
            } else {
                List<String> cellStrList = this.excelCalculateConfig.getCellStrListByFormula(parameter);
                Map<String, String> cellAndValue = excelEntity.getCellStrAndValueMap(cellStrList);
                resultList.add(ExcelUtil.functionCalculate(parameter, cellAndValue, mode));
            }
        }
        return resultList;
    }

    /**
     * 判断本节点括号范围是否严格包含另一节点（用于挂载嵌套函数）。
     */
    public boolean contain(ExcelFunctionNode<T> excelFunctionNode) {
        return this.leftIndex < excelFunctionNode.getLeftIndex() && this.rightIndex > excelFunctionNode.getRightIndex();
    }

    /** 访问器：函数原文、名称、参数串、求值结果、括号下标、配置及嵌套子节点列表。 */
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

    public Integer getLeftIndex() {
        return leftIndex;
    }

    public Integer getRightIndex() {
        return rightIndex;
    }

    public ExcelCalculateConfig getExcelCalculateConfig() {
        return excelCalculateConfig;
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
                '}';
    }
}
