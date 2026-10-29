package io.github.kanfengjingren.excelformula.boot;

import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;

/**
 * Excel 公式计算自动配置：启用 {@code excel.formula.*} 属性，
 * 并注册 {@link ExcelFormulaSettings} 与 {@link ExcelFormulaTemplate}。
 */
@AutoConfiguration
@EnableConfigurationProperties(ExcelFormulaProperties.class)
public class ExcelFormulaAutoConfiguration {

    /**
     * 提供将 application 配置应用到 {@code ExcelCalculateConfig} 的设置器。
     */
    @Bean
    @ConditionalOnMissingBean
    public ExcelFormulaSettings excelFormulaSettings(ExcelFormulaProperties properties) {
        return new ExcelFormulaSettings(properties);
    }

    /**
     * Spring 侧计算入口；业务注入后按实体类型创建 Row / Sheet / Book 计算器。
     */
    @Bean
    @ConditionalOnMissingBean
    public ExcelFormulaTemplate excelFormulaTemplate(ExcelFormulaSettings settings) {
        return new ExcelFormulaTemplate(settings);
    }
}
