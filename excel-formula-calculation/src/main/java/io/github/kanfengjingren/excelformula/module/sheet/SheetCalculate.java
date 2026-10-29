package io.github.kanfengjingren.excelformula.module.sheet;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelCalculateTypeEnum;
import io.github.kanfengjingren.excelformula.common.interfaces.AbstractExcelCalculate;
import io.github.kanfengjingren.excelformula.module.sheet.entity.ExcelSheet;

import java.util.List;

/**
 * 单页签计算器。
 * <p>
 * 将配置绑定为单页签类型，并构建页签内存模型。
 *
 * @param <T> 数据实体类型
 */
public class SheetCalculate<T> extends AbstractExcelCalculate<T> {

    /**
     * @param excelList            同一页签下的多行数据
     * @param excelCalculateConfig 计算配置（可与其它计算器共享，内部会生成类型快照）
     */
    public SheetCalculate(List<T> excelList, ExcelCalculateConfig excelCalculateConfig) {
        this(excelCalculateConfig.forType(ExcelCalculateTypeEnum.SHEET), excelList);
    }

    /** 使用已绑定 SHEET 类型的配置构建页签模型。 */
    private SheetCalculate(ExcelCalculateConfig typedConfig, List<T> excelList) {
        super(typedConfig, new ExcelSheet<>(excelList, typedConfig));
    }

}
