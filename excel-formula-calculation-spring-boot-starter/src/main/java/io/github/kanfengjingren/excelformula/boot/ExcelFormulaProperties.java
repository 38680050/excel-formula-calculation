package io.github.kanfengjingren.excelformula.boot;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import org.springframework.boot.context.properties.ConfigurationProperties;

/**
 * Spring Boot 配置项，前缀 {@code excel.formula}。
 * <p>
 * 用于为自行创建的 {@code ExcelCalculateConfig} 提供全局默认值；
 * 不会自动创建配置实例（配置依赖业务实体类）。
 */
@ConfigurationProperties(prefix = "excel.formula")
public class ExcelFormulaProperties {

    /**
     * 公式语义错误处理模式，默认抛出异常。
     */
    private ExcelErrorHandleMode errorHandleMode = ExcelErrorHandleMode.THROW;

    public ExcelErrorHandleMode getErrorHandleMode() {
        return errorHandleMode;
    }

    public void setErrorHandleMode(ExcelErrorHandleMode errorHandleMode) {
        this.errorHandleMode = ExcelErrorHandleMode.orThrow(errorHandleMode);
    }
}
