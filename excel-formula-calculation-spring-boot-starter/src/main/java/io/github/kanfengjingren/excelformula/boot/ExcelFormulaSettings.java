package io.github.kanfengjingren.excelformula.boot;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;

/**
 * 将 {@link ExcelFormulaProperties} 应用到计算配置上的辅助 Bean。
 */
public class ExcelFormulaSettings {

    private final ExcelFormulaProperties properties;

    public ExcelFormulaSettings(ExcelFormulaProperties properties) {
        this.properties = properties;
    }

    /**
     * 按全局配置生成快照：写入错误处理模式等属性，不修改入参原实例。
     */
    public ExcelCalculateConfig apply(ExcelCalculateConfig config) {
        if (config == null) {
            throw new IllegalArgumentException("ExcelCalculateConfig 不能为空");
        }
        return config.withErrorHandleMode(properties.getErrorHandleMode());
    }

    /** 当前绑定的配置属性。 */
    public ExcelFormulaProperties getProperties() {
        return properties;
    }
}
