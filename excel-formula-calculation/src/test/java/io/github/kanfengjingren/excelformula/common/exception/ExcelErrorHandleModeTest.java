package io.github.kanfengjingren.excelformula.common.exception;

import io.github.kanfengjingren.excelformula.common.calculation.arithmetic.ArithmeticEvaluator;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.util.ExcelUtil;
import io.github.kanfengjingren.excelformula.module.row.RowCalculate;
import io.github.kanfengjingren.excelformula.support.ExcelNumber;
import io.github.kanfengjingren.excelformula.support.ExcelSampleData;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/** THROW 与 RETURN_CODE 两种错误处理模式及配置快照行为。 */
@DisplayName("错误处理模式 ExcelErrorHandleMode")
class ExcelErrorHandleModeTest {

    private final ExcelCalculateConfig baseConfig = ExcelSampleData.numberConfig();

    @Nested
    @DisplayName("THROW（默认）")
    class ThrowMode {

        @Test
        void divZeroHasErrorCode() {
            assertThatThrownBy(() -> ArithmeticEvaluator.evaluate("1/0"))
                    .isInstanceOf(ExcelException.class)
                    .satisfies(ex -> assertThat(((ExcelException) ex).getExcelErrorCode())
                            .isEqualTo(ExcelConstant.EXCEL_ERROR_DIV0));
        }

        @Test
        void unknownFunctionHasNameCode() {
            RowCalculate<ExcelNumber> calc = new RowCalculate<>(ExcelSampleData.blankNumberRow(), baseConfig);
            assertThatThrownBy(() -> calc.calculate("FOO(A)"))
                    .isInstanceOf(ExcelException.class)
                    .satisfies(ex -> {
                        ExcelException ee = (ExcelException) ex;
                        assertThat(ee.getExcelErrorCode()).isEqualTo(ExcelConstant.EXCEL_ERROR_NAME);
                        assertThat(ee.getMessage()).contains("未知函数");
                    });
        }

        @Test
        void cellErrorLiteralPropagatesAsException() {
            Map<String, String> cells = new LinkedHashMap<>();
            cells.put("A1", ExcelConstant.EXCEL_ERROR_VALUE);
            assertThatThrownBy(() -> ExcelUtil.arithmeticCalculate("A1+1", cells))
                    .isInstanceOf(ExcelException.class)
                    .satisfies(ex -> assertThat(((ExcelException) ex).getExcelErrorCode())
                            .isEqualTo(ExcelConstant.EXCEL_ERROR_VALUE));
        }

        @Test
        void configErrorAlwaysThrowsEvenInReturnCodeMode() {
            ExcelCalculateConfig returnCode = baseConfig.withErrorHandleMode(ExcelErrorHandleMode.RETURN_CODE);
            assertThatThrownBy(() -> returnCode.findReturnCellPrefix("A1=1"))
                    .isInstanceOf(ExcelException.class)
                    .hasMessageContaining("计算类型");
        }
    }

    @Nested
    @DisplayName("RETURN_CODE")
    class ReturnCodeMode {

        @Test
        void divZeroReturnsLiteral() {
            assertThat(ArithmeticEvaluator.evaluate("1/0", ExcelErrorHandleMode.RETURN_CODE))
                    .isEqualTo(ExcelConstant.EXCEL_ERROR_DIV0);
        }

        @Test
        void unknownFunctionReturnsName() {
            RowCalculate<ExcelNumber> calc = new RowCalculate<>(ExcelSampleData.blankNumberRow(),
                    baseConfig.withErrorHandleMode(ExcelErrorHandleMode.RETURN_CODE));
            assertThat(calc.calculate("FOO(A)")).isEqualTo(ExcelConstant.EXCEL_ERROR_NAME);
        }

        @Test
        void cellErrorPropagates() {
            Map<String, String> cells = new LinkedHashMap<>();
            cells.put("A1", ExcelConstant.EXCEL_ERROR_VALUE);
            assertThat(ExcelUtil.arithmeticCalculate("A1+1", cells, ExcelErrorHandleMode.RETURN_CODE))
                    .isEqualTo(ExcelConstant.EXCEL_ERROR_VALUE);
        }

        @Test
        void invalidArithmeticReturnsValue() {
            assertThat(ArithmeticEvaluator.evaluate("1+", ExcelErrorHandleMode.RETURN_CODE))
                    .isEqualTo(ExcelConstant.EXCEL_ERROR_VALUE);
        }
    }

    @Nested
    @DisplayName("配置快照")
    class ConfigSnapshot {

        @Test
        void withErrorHandleModeDoesNotMutateOriginal() {
            assertThat(baseConfig.getErrorHandleMode()).isEqualTo(ExcelErrorHandleMode.THROW);
            ExcelCalculateConfig snapshot = baseConfig.withErrorHandleMode(ExcelErrorHandleMode.RETURN_CODE);
            assertThat(baseConfig.getErrorHandleMode()).isEqualTo(ExcelErrorHandleMode.THROW);
            assertThat(snapshot.getErrorHandleMode()).isEqualTo(ExcelErrorHandleMode.RETURN_CODE);
            assertThat(snapshot).isNotSameAs(baseConfig);
        }
    }
}
