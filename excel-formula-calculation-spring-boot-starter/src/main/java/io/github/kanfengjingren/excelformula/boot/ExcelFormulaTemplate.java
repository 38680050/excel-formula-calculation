package io.github.kanfengjingren.excelformula.boot;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.module.book.BookCalculate;
import io.github.kanfengjingren.excelformula.module.row.RowCalculate;
import io.github.kanfengjingren.excelformula.module.sheet.SheetCalculate;

import java.util.List;
import java.util.Objects;

/**
 * Spring 侧计算入口：按实体类型创建计算器，并自动套用 {@code excel.formula.*} 全局默认值。
 * <p>
 * 业务代码注入本类即可，无需再把 {@link ExcelCalculateConfig} 注册成 Bean。
 */
public class ExcelFormulaTemplate {

    private final ExcelFormulaSettings settings;

    public ExcelFormulaTemplate(ExcelFormulaSettings settings) {
        this.settings = Objects.requireNonNull(settings, "ExcelFormulaSettings 不能为空");
    }

    /**
     * 按实体类创建配置，并应用全局默认（如错误处理模式）。
     */
    public ExcelCalculateConfig config(Class<?> entityClass) {
        Objects.requireNonNull(entityClass, "entityClass 不能为空");
        return settings.apply(ExcelCalculateConfig.create(entityClass));
    }

    /**
     * 对已有配置（例如带自定义函数的 Builder 产物）套上全局默认，不修改入参原实例。
     */
    public ExcelCalculateConfig config(ExcelCalculateConfig config) {
        return settings.apply(config);
    }

    /** 单行计算器。 */
    public <T> RowCalculate<T> row(Class<T> entityClass, T row) {
        return new RowCalculate<>(row, config(entityClass));
    }

    /** 单页签计算器。 */
    public <T> SheetCalculate<T> sheet(Class<T> entityClass, List<T> rows) {
        return new SheetCalculate<>(rows, config(entityClass));
    }

    /** 多页签计算器。 */
    public <T> BookCalculate<T> book(Class<T> entityClass, List<T> rows) {
        return new BookCalculate<>(rows, config(entityClass));
    }

    /** 使用自定义配置创建单行计算器（仍会套用全局默认）。 */
    public <T> RowCalculate<T> row(ExcelCalculateConfig config, T row) {
        return new RowCalculate<>(row, config(config));
    }

    /** 使用自定义配置创建单页签计算器（仍会套用全局默认）。 */
    public <T> SheetCalculate<T> sheet(ExcelCalculateConfig config, List<T> rows) {
        return new SheetCalculate<>(rows, config(config));
    }

    /** 使用自定义配置创建多页签计算器（仍会套用全局默认）。 */
    public <T> BookCalculate<T> book(ExcelCalculateConfig config, List<T> rows) {
        return new BookCalculate<>(rows, config(config));
    }

    /** 当前绑定的设置器。 */
    public ExcelFormulaSettings getSettings() {
        return settings;
    }
}
