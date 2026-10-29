package io.github.kanfengjingren.excelformula.common.function;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * 内置函数语义单测（经 FunctionRegistry 调用，覆盖错误短路例外）。
 */
@DisplayName("DefaultFunctionImpl 内置函数")
class DefaultFunctionImplTest {

    private final FunctionRegistry registry = FunctionRegistry.withDefaults().freeze();

    private String call(String name, String... args) {
        return registry.invoke(name, Arrays.asList(args), ExcelErrorHandleMode.THROW);
    }

    @Nested
    @DisplayName("逻辑")
    class Logic {
        @Test
        void andOrNot() {
            assertThat(call("AND", "TRUE", "1")).isEqualTo("TRUE");
            assertThat(call("AND", "TRUE", "0")).isEqualTo("FALSE");
            assertThat(call("OR", "0", "FALSE")).isEqualTo("FALSE");
            assertThat(call("OR", "0", "2")).isEqualTo("TRUE");
            assertThat(call("NOT", "TRUE")).isEqualTo("FALSE");
            assertThat(call("NOT", "0")).isEqualTo("TRUE");
        }

        @Test
        void ifErrorAndIsError() {
            assertThat(call("IFERROR", ExcelConstant.EXCEL_ERROR_DIV0, "fallback")).isEqualTo("fallback");
            assertThat(call("IFERROR", "12", "fallback")).isEqualTo("12");
            assertThat(call("IFNA", ExcelConstant.EXCEL_ERROR_NA, "x")).isEqualTo("x");
            assertThat(call("IFNA", ExcelConstant.EXCEL_ERROR_VALUE, "x")).isEqualTo(ExcelConstant.EXCEL_ERROR_VALUE);
            assertThat(call("ISERROR", ExcelConstant.EXCEL_ERROR_NAME)).isEqualTo("TRUE");
            assertThat(call("ISERROR", "1")).isEqualTo("FALSE");
            assertThat(call("ISBLANK", "")).isEqualTo("TRUE");
            assertThat(call("ISBLANK", "a")).isEqualTo("FALSE");
            assertThat(call("TRUE")).isEqualTo("TRUE");
            assertThat(call("FALSE")).isEqualTo("FALSE");
        }
    }

    @Nested
    @DisplayName("数学统计")
    class MathStats {
        @Test
        void averageCountProduct() {
            assertThat(call("AVERAGE", "1", "2", "3")).isEqualTo("2");
            assertThat(call("COUNT", "1", "a", "", "3")).isEqualTo("2");
            assertThat(call("COUNTA", "1", "a", "", "3")).isEqualTo("3");
            assertThat(call("PRODUCT", "2", "3", "4")).isEqualTo("24");
        }

        @Test
        void averageEmptyThrowsDiv0() {
            assertThatThrownBy(() -> call("AVERAGE", "a", ""))
                    .isInstanceOf(ExcelException.class)
                    .extracting(ex -> ((ExcelException) ex).getExcelErrorCode())
                    .isEqualTo(ExcelConstant.EXCEL_ERROR_DIV0);
        }

        @Test
        void powerSqrtModIntRound() {
            assertThat(call("POWER", "2", "10")).isEqualTo("1024");
            assertThat(call("SQRT", "9")).isEqualTo("3");
            assertThat(call("MOD", "5", "2")).isEqualTo("1");
            assertThat(call("MOD", "-3", "2")).isEqualTo("1");
            assertThat(call("INT", "3.9")).isEqualTo("3");
            assertThat(call("INT", "-3.1")).isEqualTo("-4");
            assertThat(call("ROUNDUP", "1.21", "1")).isEqualTo("1.3");
            assertThat(call("ROUNDDOWN", "1.29", "1")).isEqualTo("1.2");
            assertThat(call("SIGN", "-8")).isEqualTo("-1");
        }
    }

    @Nested
    @DisplayName("文本")
    class Text {
        @Test
        void lenLeftRightMid() {
            assertThat(call("LEN", "abc")).isEqualTo("3");
            assertThat(call("LEFT", "abcdef", "2")).isEqualTo("ab");
            assertThat(call("RIGHT", "abcdef", "3")).isEqualTo("def");
            assertThat(call("MID", "abcdef", "2", "3")).isEqualTo("bcd");
        }

        @Test
        void trimCaseConcatValue() {
            assertThat(call("TRIM", "  a b  ")).isEqualTo("a b");
            assertThat(call("UPPER", "Ab")).isEqualTo("AB");
            assertThat(call("LOWER", "Ab")).isEqualTo("ab");
            assertThat(call("CONCATENATE", "A", "1", "B")).isEqualTo("A1B");
            assertThat(call("CONCAT", "x", "y")).isEqualTo("xy");
            assertThat(call("VALUE", "12.5")).isEqualTo("12.5");
        }
    }

    @Test
    @DisplayName("普通函数参数含错误时仍短路")
    void normalFunctionStillShortCircuitsOnError() {
        assertThatThrownBy(() -> call("SUM", "1", ExcelConstant.EXCEL_ERROR_VALUE))
                .isInstanceOf(ExcelException.class);
        assertThat(registry.invoke("SUM", List.of("1", ExcelConstant.EXCEL_ERROR_VALUE), ExcelErrorHandleMode.RETURN_CODE))
                .isEqualTo(ExcelConstant.EXCEL_ERROR_VALUE);
    }

    @Test
    @DisplayName("无参 TRUE/FALSE 可调用")
    void noArgTrueFalse() {
        assertThat(registry.invoke("TRUE", Collections.emptyList())).isEqualTo("TRUE");
        assertThat(registry.invoke("FALSE", Collections.emptyList())).isEqualTo("FALSE");
    }
}
