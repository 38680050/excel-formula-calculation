package io.github.kanfengjingren.excelformula.integration;

import io.github.kanfengjingren.excelformula.module.row.RowCalculate;
import io.github.kanfengjingren.excelformula.support.ExcelNumber;
import io.github.kanfengjingren.excelformula.support.ExcelSampleData;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * 单行 {@link RowCalculate} 的相对列引用、嵌套函数与 ROUND 写回。
 */
@DisplayName("集成：单行 RowCalculate")
class RowCalculateTest {

    @Test
    @DisplayName("数字主键：列引用与 ROUND 写回")
    void calculatesWithNumberIds() {
        ExcelNumber excel = ExcelSampleData.singleRow();
        RowCalculate<ExcelNumber> calc = new RowCalculate<>(excel, ExcelSampleData.numberConfig());
        String result1 = calc.calculate("I=A+ABS(MIN(O+2,B)*2-MAX(A:H))+MAX(K:N)");
        String result2 = calc.calculate("J=ROUND(2*I+5,0)");
        calc.integrationResult();
        assertThat(result1).isEqualTo("1000295");
        assertThat(result2).isEqualTo("2000595");
        assertThat(excel.getI()).isEqualTo("1000295");
        assertThat(excel.getJ()).isEqualTo("2000595");
    }

    @Test
    @DisplayName("UUID Codec 配置下单行计算（列引用不依赖行 ID）")
    void calculatesWithUuidCodec() {
        ExcelNumber excel = ExcelSampleData.singleRow();
        RowCalculate<ExcelNumber> calc = new RowCalculate<>(excel, ExcelSampleData.uuidConfig());
        String result1 = calc.calculate("I=A+ABS(MIN(O+2,B)*2-MAX(A:H))+MAX(K:N)");
        String result2 = calc.calculate("J=ROUND(2*I+5,0)");
        calc.integrationResult();
        assertThat(result1).isEqualTo("1000295");
        assertThat(result2).isEqualTo("2000595");
        assertThat(excel.getI()).isEqualTo("1000295");
        assertThat(excel.getJ()).isEqualTo("2000595");
    }
}
