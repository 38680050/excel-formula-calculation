package io.github.kanfengjingren.excelformula.module.row;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelCalculateTypeEnum;
import io.github.kanfengjingren.excelformula.common.interfaces.AbstractExcelCalculate;
import io.github.kanfengjingren.excelformula.module.row.entity.ExcelRow;

/**
 * 单行计算器。
 * <p>
 * 将配置绑定为单行类型，并构建单行内存模型；公式中的单元格引用为列名（如 A、B）。
 *
 * @param <T> 数据实体类型
 */
public class RowCalculate<T> extends AbstractExcelCalculate<T> {

    /**
     * @param t                    单行数据对象
     * @param excelCalculateConfig 计算配置（可与其它计算器共享，内部会生成类型快照）
     */
    public RowCalculate(T t, ExcelCalculateConfig excelCalculateConfig) {
        this(excelCalculateConfig.forType(ExcelCalculateTypeEnum.ROW), t);
    }

    /** 使用已绑定 ROW 类型的配置构建单行模型。 */
    private RowCalculate(ExcelCalculateConfig typedConfig, T t) {
        super(typedConfig, new ExcelRow<>(t, typedConfig));
    }

}
