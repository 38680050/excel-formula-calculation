package io.github.kanfengjingren.excelformula.common.calculation.formula;

import io.github.kanfengjingren.excelformula.common.calculation.function.ExcelFunction;
import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCell;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelErrorUtil;
import io.github.kanfengjingren.excelformula.common.interfaces.IExcelEntity;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.util.ExcelUtil;

import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * Excel 单条公式对象。
 * <p>
 * 负责识别赋值目标单元格、解析并计算函数、将单元格引用替换为值后做算术求值，
 * 若有赋值目标则把结果写入返回单元格对象。
 *
 * @param <T> 数据实体类型
 */
public class ExcelFormula<T> {

    /** 赋值目标单元格；公式以「单元格=」开头时存在 */
    private ExcelCell returnCell;

    /** 去掉赋值前缀后的待计算表达式 */
    private String formula;

    /** 本公式中的函数解析与求值器 */
    private final ExcelFunction<T> excelFunction;

    /** 计算配置 */
    private final ExcelCalculateConfig properties;

    /**
     * 构造公式：剥离赋值前缀（若有），再用剩余表达式构建函数树。
     *
     * @param currentSheet 当前页签名，多页签时用于解析相对引用
     * @param formula      完整公式，可带赋值前缀如 C1=A1+B1
     * @param properties   计算配置
     */
    public ExcelFormula(String currentSheet, String formula, ExcelCalculateConfig properties) {
        this.properties = properties;
        Optional<String> returnPrefix = this.properties.findReturnCellPrefix(formula);
        if (returnPrefix.isPresent()) {
            String prefix = returnPrefix.get();
            // 前缀含末尾 =，单元格引用为去掉等号的部分
            String cellRef = prefix.substring(0, prefix.length() - 1);
            this.returnCell = new ExcelCell(cellRef, this.properties.getRowIdCodec(), currentSheet);
            this.formula = formula.substring(prefix.length());
        } else {
            this.formula = formula;
        }
        // 必须使用剥离赋值前缀后的公式解析函数，避免把赋值段误解析进函数树
        this.excelFunction = new ExcelFunction<>(this.formula, this.properties);
    }

    /**
     * 计算流程：先算全部函数并替换回公式，再解析剩余单元格引用并做算术求值。
     *
     * @param currentSheet 当前页签名
     * @param excelEntity  单元格数据实体
     * @return 计算结果字符串
     */
    public String calculate(String currentSheet, IExcelEntity<T> excelEntity) {
        this.excelFunction.functionCalculate(currentSheet, excelEntity);
        for (var functionNode : this.excelFunction.getExcelFunctionNodeList()) {
            String nodeValue = functionNode.getValue();
            // 禁止 replace(..., null)：null 按空单元格语义规范为 0
            if (nodeValue == null) {
                nodeValue = ExcelConstant.ZERO_STR;
            }
            if (ExcelErrorUtil.isError(nodeValue)) {
                String failed = ExcelErrorUtil.fail(this.properties, nodeValue, "函数结果错误: " + nodeValue);
                if (null != this.returnCell) {
                    this.returnCell.setOriginalValue(failed);
                }
                return failed;
            }
            this.formula = formula.replace(functionNode.getFunction(), nodeValue);
        }
        List<String> cellStrList = this.properties.getCellStrListByFormula(this.formula);
        Map<String, String> cellAndValue = excelEntity.getCellStrAndValueMap(cellStrList);
        String value = ExcelUtil.arithmeticCalculate(this.formula, cellAndValue, this.properties.getErrorHandleMode());
        if (null != this.returnCell) {
            this.returnCell.setOriginalValue(value);
        }
        return value;
    }

    /** 赋值目标单元格；公式无「单元格=」前缀时为 null。 */
    public ExcelCell getReturnCell() {
        return returnCell;
    }

    /** 去掉赋值前缀后、待参与函数与算术求值的表达式。 */
    public String getFormula() {
        return formula;
    }

    /** 本公式解析出的函数树及求值器。 */
    public ExcelFunction<T> getExcelFunction() {
        return excelFunction;
    }

    /** 计算配置（引用解析、函数注册、错误模式等）。 */
    public ExcelCalculateConfig getProperties() {
        return properties;
    }

    @Override
    public String toString() {
        return "ExcelFormula{" +
                "returnCell=" + returnCell +
                ", formula='" + formula + '\'' +
                ", excelFunction=" + excelFunction +
                ", excelEntityProperties=" + properties +
                '}';
    }

}
