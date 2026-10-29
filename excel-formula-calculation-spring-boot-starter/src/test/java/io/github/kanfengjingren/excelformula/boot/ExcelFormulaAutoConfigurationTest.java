package io.github.kanfengjingren.excelformula.boot;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.field.annotation.ExcelField;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelFieldTypeEnum;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;
import io.github.kanfengjingren.excelformula.module.sheet.SheetCalculate;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.boot.autoconfigure.AutoConfigurations;
import org.springframework.boot.test.context.runner.ApplicationContextRunner;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 验证 starter 自动配置与属性绑定。
 */
@DisplayName("ExcelFormula Spring Boot 自动配置")
class ExcelFormulaAutoConfigurationTest {

    private final ApplicationContextRunner contextRunner = new ApplicationContextRunner()
            .withConfiguration(AutoConfigurations.of(ExcelFormulaAutoConfiguration.class));

    @Test
    @DisplayName("默认注册 Settings 与 Template，错误模式为 THROW")
    void registersSettingsAndTemplateWithDefaultThrow() {
        contextRunner.run(context -> {
            assertThat(context).hasSingleBean(ExcelFormulaSettings.class);
            assertThat(context).hasSingleBean(ExcelFormulaTemplate.class);

            ExcelFormulaTemplate template = context.getBean(ExcelFormulaTemplate.class);
            assertThat(template.getSettings().getProperties().getErrorHandleMode())
                    .isEqualTo(ExcelErrorHandleMode.THROW);

            ExcelCalculateConfig applied = template.config(DemoRow.class);
            assertThat(applied.getErrorHandleMode()).isEqualTo(ExcelErrorHandleMode.THROW);
        });
    }

    @Test
    @DisplayName("可通过 excel.formula.error-handle-mode 覆盖错误模式")
    void bindsErrorHandleModeFromProperties() {
        contextRunner
                .withPropertyValues("excel.formula.error-handle-mode=RETURN_CODE")
                .run(context -> {
                    ExcelFormulaTemplate template = context.getBean(ExcelFormulaTemplate.class);
                    ExcelCalculateConfig applied = template.config(DemoRow.class);
                    assertThat(applied.getErrorHandleMode()).isEqualTo(ExcelErrorHandleMode.RETURN_CODE);
                });
    }

    @Test
    @DisplayName("Template.sheet 可按实体创建计算器并完成计算")
    void templateSheetCalculates() {
        contextRunner.run(context -> {
            ExcelFormulaTemplate template = context.getBean(ExcelFormulaTemplate.class);
            DemoRow row = new DemoRow();
            row.setId("1");
            row.setSort(1);
            row.setA("2");
            row.setB("3");

            SheetCalculate<DemoRow> calc = template.sheet(DemoRow.class, List.of(row));
            String value = calc.calculate("A1=A1+B1");
            calc.integrationResult();

            assertThat(value).isEqualTo("5");
            assertThat(row.getA()).isEqualTo("5");
        });
    }

    /** 测试用最小行实体。 */
    public static class DemoRow {
        @ExcelField(value = ExcelFieldTypeEnum.ID, idType = ExcelIdTypeEnum.NUMBER)
        private String id;
        @ExcelField
        private String A;
        @ExcelField
        private String B;
        @ExcelField(ExcelFieldTypeEnum.SORT)
        private Integer sort;

        public String getId() {
            return id;
        }

        public void setId(String id) {
            this.id = id;
        }

        public String getA() {
            return A;
        }

        public void setA(String a) {
            A = a;
        }

        public String getB() {
            return B;
        }

        public void setB(String b) {
            B = b;
        }

        public Integer getSort() {
            return sort;
        }

        public void setSort(Integer sort) {
            this.sort = sort;
        }
    }
}
