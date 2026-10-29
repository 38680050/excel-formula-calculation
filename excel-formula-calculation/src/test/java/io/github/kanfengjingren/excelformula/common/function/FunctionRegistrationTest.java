package io.github.kanfengjingren.excelformula.common.function;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.field.annotation.ExcelField;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelFieldTypeEnum;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;
import io.github.kanfengjingren.excelformula.common.interfaces.IFunction;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 自定义函数叠加与同名覆盖优先级。
 */
@DisplayName("函数注册：内置叠加与覆盖")
class FunctionRegistrationTest {

    @Test
    @DisplayName("仅默认配置包含 SUM")
    void defaultsContainSum() {
        ExcelCalculateConfig config = ExcelCalculateConfig.create(DemoRow.class);
        assertThat(config.getFunctionRegistry().contains("SUM")).isTrue();
        assertThat(config.functionCalculate("SUM", List.of("1", "2"))).isEqualTo("3");
    }

    @Test
    @DisplayName("类叠加新增 POWER，且 SUM 仍可用")
    void classAdditiveKeepsBuiltins() {
        ExcelCalculateConfig config = ExcelCalculateConfig.of(DemoRow.class)
                .functions(new ExtraPower())
                .build();
        assertThat(config.functionCalculate("SUM", List.of("1", "2"))).isEqualTo("3");
        assertThat(config.functionCalculate("POWER", List.of("2", "3"))).isEqualTo("custom-8");
    }

    @Test
    @DisplayName("Lambda 可覆盖内置 ROUND")
    void lambdaOverridesRound() {
        ExcelCalculateConfig config = ExcelCalculateConfig.of(DemoRow.class)
                .function("ROUND", args -> "rounded")
                .build();
        assertThat(config.functionCalculate("ROUND", List.of("1.23", "0"))).isEqualTo("rounded");
        assertThat(config.functionCalculate("SUM", List.of("1", "1"))).isEqualTo("2");
    }

    @Test
    @DisplayName("多源同名：后注册胜出")
    void laterRegistrationWins() {
        ExcelCalculateConfig config = ExcelCalculateConfig.of(DemoRow.class)
                .functions(new TagA())
                .function("TAG", args -> "lambda")
                .functions(new TagB())
                .build();
        assertThat(config.functionCalculate("TAG", List.of())).isEqualTo("B");
    }

    public static class DemoRow {
        @ExcelField(value = ExcelFieldTypeEnum.ID, idType = ExcelIdTypeEnum.NUMBER)
        private String id;
        @ExcelField
        private String A;
    }

    public static class ExtraPower implements IFunction {
        public String POWER(List<String> args) {
            return "custom-8";
        }
    }

    public static class TagA implements IFunction {
        public String TAG(List<String> args) {
            return "A";
        }
    }

    public static class TagB implements IFunction {
        public String TAG(List<String> args) {
            return "B";
        }
    }
}
