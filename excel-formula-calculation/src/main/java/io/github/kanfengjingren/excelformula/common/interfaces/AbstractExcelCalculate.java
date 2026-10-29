package io.github.kanfengjingren.excelformula.common.interfaces;

import io.github.kanfengjingren.excelformula.common.calculation.formula.ExcelFormula;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;

/**
 * 计算器抽象基类。
 * <p>
 * 统一「解析公式 → 求值 → 若有赋值目标则写回单元格 → 可选择收敛到原 POJO」的流程，
 * 供单行、单页签、多页签计算器复用。
 *
 * @param <T> 数据实体类型
 */
public abstract class AbstractExcelCalculate<T> implements IExcelCalculate {

    /** 已绑定计算类型的配置快照 */
    protected final ExcelCalculateConfig excelCalculateConfig;

    /** 与当前计算粒度对应的内存单元格模型（行 / 页签 / 工作簿）。 */
    protected final IExcelEntity<T> excelEntity;

    /**
     * @param excelCalculateConfig 计算配置（应已绑定计算类型）
     * @param excelEntity          对应级别的数据实体
     */
    protected AbstractExcelCalculate(ExcelCalculateConfig excelCalculateConfig, IExcelEntity<T> excelEntity) {
        this.excelCalculateConfig = excelCalculateConfig;
        this.excelEntity = excelEntity;
    }

    /**
     * 按公式计算；当前页签传空（单行、单页签场景）。
     */
    @Override
    public String calculate(String formula) {
        return doCalculate(null, formula);
    }

    /**
     * 执行一次公式计算。
     * 若公式带赋值前缀（如 C1=...），将结果写入对应单元格并返回写回后的原值形式。
     *
     * @param currentSheet 当前页签名，仅多页签计算时使用
     * @param formula      完整公式字符串
     * @return 计算结果
     */
    protected String doCalculate(String currentSheet, String formula) {
        ExcelFormula<T> excelFormula = new ExcelFormula<>(currentSheet, formula, this.excelCalculateConfig);
        String value = excelFormula.calculate(currentSheet, this.excelEntity);
        if (null != excelFormula.getReturnCell()) {
            value = this.excelEntity.updateExcelCellValue(excelFormula.getReturnCell());
        }
        return value;
    }

    /**
     * 将内存中单元格结果写回原始 POJO 字段。
     */
    @Override
    public void integrationResult() {
        this.excelEntity.integrationResult();
    }

    /** 获取当前计算器使用的配置 */
    public ExcelCalculateConfig getExcelCalculateConfig() {
        return excelCalculateConfig;
    }
}
