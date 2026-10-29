package io.github.kanfengjingren.excelformula.integration;

import io.github.kanfengjingren.excelformula.module.sheet.SheetCalculate;
import io.github.kanfengjingren.excelformula.support.ExcelNumber;
import io.github.kanfengjingren.excelformula.support.ExcelSampleData;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.List;

import static io.github.kanfengjingren.excelformula.support.ExcelSampleData.UUID_1;
import static io.github.kanfengjingren.excelformula.support.ExcelSampleData.UUID_2;
import static io.github.kanfengjingren.excelformula.support.ExcelSampleData.UUID_3;
import static io.github.kanfengjingren.excelformula.support.ExcelSampleData.UUID_5;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * 单页签 {@link SheetCalculate} 的连续公式计算、类型收敛与写回。
 */
@DisplayName("集成：单页签 SheetCalculate")
class SheetCalculateTest {

    @Test
    @DisplayName("数字主键：连续计算与类型收敛写回")
    void calculatesWithNumberIds() {
        List<ExcelNumber> excelList = ExcelSampleData.sheetNumbers();
        SheetCalculate<ExcelNumber> calc = new SheetCalculate<>(excelList, ExcelSampleData.numberConfig());
        String result1 = calc.calculate("K5=A1+ABS(MIN(C5+2,A1:I5)*2-MAX(L1:N5))+MAX(A2:J2)");
        String result2 = calc.calculate("N5=ROUND(2*K5+5,0)");
        calc.integrationResult();
        assertThat(result1).isEqualTo("2110-09-01");
        assertThat(result2).isEqualTo("15389900%");
        assertThat(excelList.get(4).getK()).isEqualTo(result1);
        assertThat(excelList.get(4).getN()).isEqualTo(result2);
    }

    @Test
    @DisplayName("UUID Codec：连续计算写回")
    void calculatesWithUuidCodec() {
        List<ExcelNumber> excelList = ExcelSampleData.sheetUuids();
        SheetCalculate<ExcelNumber> calc = new SheetCalculate<>(excelList, ExcelSampleData.uuidConfig());
        String result1 = calc.calculate("K" + UUID_5 + "=A" + UUID_1
                + "+ABS(MIN(C" + UUID_5 + "+2,A" + UUID_1 + ":I" + UUID_5
                + ")*2-MAX(L" + UUID_1 + ":N" + UUID_5 + "))+MAX(A" + UUID_2 + ":J" + UUID_2 + ")");
        String result2 = calc.calculate("N" + UUID_5 + "=ROUND(2*K" + UUID_5 + "+5,0)");
        calc.integrationResult();
        assertThat(result1).isEqualTo("2110-09-01");
        assertThat(result2).isEqualTo("15389900%");
        assertThat(excelList.get(4).getK()).isEqualTo(result1);
        assertThat(excelList.get(4).getN()).isEqualTo(result2);
    }
}
