package com.wsbxd.excel.formula.calculation.common.calculation.formula;

import com.wsbxd.excel.formula.calculation.common.calculation.function.ExcelFunction;
import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelEntity;
import com.wsbxd.excel.formula.calculation.common.config.ExcelCalculateConfig;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;

import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

/**
 * description: Excel 公式
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/3/13 16:55
 */
public class ExcelFormula<T> {

    /**
     * 返回单元格
     */
    private ExcelCell returnCell;

    /**
     * 公式
     */
    private String formula;

    /**
     * Excel 函数
     */
    private final ExcelFunction<T> excelFunction;

    /**
     * Excel 实体类 数据属性
     */
    private final ExcelCalculateConfig properties;

    /**
     * Excel 公式
     *
     * @param currentSheet 当前页签名称，只有工作簿计算时不为空
     * @param formula      公式
     * @param properties   Excel 实体类 数据属性
     */
    public ExcelFormula(String currentSheet, String formula, ExcelCalculateConfig properties) {
        this.properties = properties;
        //返回单元格
        Matcher matcher = this.properties.getReturnCellPattern().matcher(formula);
        if (matcher.find()) {
            this.returnCell = new ExcelCell(this.properties.getCellStrListByFormula(matcher.group()).get(0), this.properties.getExcelIdTypeEnum(), currentSheet);
            this.formula = formula.substring(matcher.end());
        } else {
            this.formula = formula;
        }
        this.excelFunction = new ExcelFunction<>(formula, this.properties);
    }

    /**
     * 公式计算
     *
     * @param currentSheet 当前页签名称，只有工作簿计算时不为空
     * @param excelEntity  Excel 实体类
     * @return 计算结果
     */
    public String calculate(String currentSheet, IExcelEntity<T> excelEntity) {
        this.excelFunction.functionCalculate(currentSheet, excelEntity);
        //公式替换为值
        this.excelFunction.getExcelFunctionNodeList().forEach(functionNode -> {
            this.formula = formula.replace(functionNode.getFunction(), functionNode.getValue());
        });
        //解析单元格获取值
        List<String> cellStrList = this.properties.getCellStrListByFormula(this.formula);
        Map<String, String> cellAndValue = excelEntity.getCellStrAndValueMap(cellStrList);
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

    public String getFormula() {
        return formula;
    }

    public ExcelFunction<T> getExcelFunction() {
        return excelFunction;
    }

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
