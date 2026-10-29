package io.github.kanfengjingren.excelformula.module.book;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelCalculateTypeEnum;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.interfaces.AbstractExcelCalculate;
import io.github.kanfengjingren.excelformula.module.book.entity.ExcelBook;

import java.util.List;

/**
 * 多页签（工作簿）计算器。
 * <p>
 * 将配置绑定为多页签类型，并构建工作簿内存模型；计算时必须指定当前页签名称。
 *
 * @param <T> 数据实体类型
 */
public class BookCalculate<T> extends AbstractExcelCalculate<T> {

    /**
     * @param excelList            含多个页签数据的列表
     * @param excelCalculateConfig 计算配置（可与其它计算器共享，内部会生成类型快照）
     */
    public BookCalculate(List<T> excelList, ExcelCalculateConfig excelCalculateConfig) {
        this(excelCalculateConfig.forType(ExcelCalculateTypeEnum.BOOK), excelList);
    }

    /** 使用已绑定 BOOK 类型的配置构建工作簿模型。 */
    private BookCalculate(ExcelCalculateConfig typedConfig, List<T> excelList) {
        super(typedConfig, new ExcelBook<>(excelList, typedConfig));
    }

    /**
     * 多页签场景禁止只传公式；请使用带当前页签名的重载方法。
     */
    @Override
    public String calculate(String formula) {
        throw new ExcelException("多页签计算请使用 calculate(currentSheet, formula)");
    }

    /**
     * 在指定当前页签下计算公式。
     *
     * @param currentSheet 当前页签名称
     * @param formula      公式
     * @return 计算结果
     */
    public String calculate(String currentSheet, String formula) {
        return doCalculate(currentSheet, formula);
    }

}
