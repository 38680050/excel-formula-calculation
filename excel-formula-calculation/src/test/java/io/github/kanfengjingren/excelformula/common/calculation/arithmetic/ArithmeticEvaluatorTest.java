package io.github.kanfengjingren.excelformula.common.calculation.arithmetic;

import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.util.ExcelUtil;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/** 算术表达式求值、单元格占位替换及非法公式与除零错误。 */
@DisplayName("算术求值 ArithmeticEvaluator")
class ArithmeticEvaluatorTest {

    @Nested
    @DisplayName("四则与比较")
    class Evaluation {

        @Test
        void precedenceAndUnaryMinus() {
            assertThat(ArithmeticEvaluator.evaluate("1+2*3")).isEqualTo("7");
            assertThat(ArithmeticEvaluator.evaluate("(1+2)*3")).isEqualTo("9");
            assertThat(ArithmeticEvaluator.evaluate("10/4")).isEqualTo("2.5");
            assertThat(ArithmeticEvaluator.evaluate("-5+2")).isEqualTo("-3");
        }

        @Test
        void avoidsFloatingPointPrecisionLoss() {
            assertThat(ArithmeticEvaluator.evaluate("0.1+0.2")).isEqualTo("0.3");
            assertThat(ArithmeticEvaluator.evaluate("0.3-0.1-0.1-0.1")).isEqualTo("0");
            assertThat(ArithmeticEvaluator.evaluate("1.005*100")).isEqualTo("100.5");
        }

        @Test
        void comparisons() {
            assertThat(ArithmeticEvaluator.evaluate("3=3")).isEqualTo(ExcelConstant.BIG_CASE_TRUE);
            assertThat(ArithmeticEvaluator.evaluate("3=4")).isEqualTo(ExcelConstant.BIG_CASE_FALSE);
            assertThat(ArithmeticEvaluator.evaluate("5>=3")).isEqualTo(ExcelConstant.BIG_CASE_TRUE);
            assertThat(ArithmeticEvaluator.evaluate("2<=2")).isEqualTo(ExcelConstant.BIG_CASE_TRUE);
            assertThat(ArithmeticEvaluator.evaluate("2<>3")).isEqualTo(ExcelConstant.BIG_CASE_TRUE);
            assertThat(ArithmeticEvaluator.evaluate("5>3")).isEqualTo(ExcelConstant.BIG_CASE_TRUE);
            assertThat(ArithmeticEvaluator.evaluate("1<2")).isEqualTo(ExcelConstant.BIG_CASE_TRUE);
        }
    }

    @Nested
    @DisplayName("单元格替换（ExcelUtil）")
    class CellReplace {

        @Test
        void prefersLongerReferences() {
            Map<String, String> cells = new LinkedHashMap<>();
            cells.put("A1", "5");
            cells.put("A10", "100");
            assertThat(ExcelUtil.arithmeticCalculate("A10+A1", cells)).isEqualTo("105");
        }

        @Test
        void blankCellsBecomeZero() {
            Map<String, String> cells = new LinkedHashMap<>();
            cells.put("B1", null);
            assertThat(ExcelUtil.arithmeticCalculate("B1+2", cells)).isEqualTo("2");
        }
    }

    @Nested
    @DisplayName("错误")
    class Errors {

        @Test
        void rejectsInvalidFormula() {
            assertThatThrownBy(() -> ArithmeticEvaluator.evaluate("1+"))
                    .isInstanceOf(ExcelException.class);
        }

        @Test
        void dividesByZeroThrowsWithCode() {
            assertThatThrownBy(() -> ArithmeticEvaluator.evaluate("1/0"))
                    .isInstanceOf(ExcelException.class)
                    .satisfies(ex -> assertThat(((ExcelException) ex).getExcelErrorCode())
                            .isEqualTo(ExcelConstant.EXCEL_ERROR_DIV0));
        }
    }
}
