package io.github.kanfengjingren.excelformula.integration;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.module.book.BookCalculate;
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
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * 多页签 {@link BookCalculate} 的公式求值、跨页引用与结果写回。
 */
@DisplayName("集成：多页签 BookCalculate")
class BookCalculateTest {

    @Test
    @DisplayName("数字主键：跨页签引用与嵌套函数写回")
    void calculatesWithNumberIds() {
        List<ExcelNumber> excelList = ExcelSampleData.bookNumbers();
        BookCalculate<ExcelNumber> calc = new BookCalculate<>(excelList, ExcelSampleData.numberConfig());
        String result = calc.calculate("交易性-股票",
                "N3=现金!A1*银行存款!C1+ABS(MIN(O2+2,银行存款!B3:银行存款!E5)*2-MAX('交易性-股票'!A1:'交易性-股票'!D3))+MAX(现金!A2:现金!O3)");
        calc.integrationResult();
        assertThat(result).isEqualTo("86504.9");
        assertThat(excelList.get(2).getN()).isEqualTo(result);
    }

    @Test
    @DisplayName("UUID 主键工作簿跨页签公式可正确求值并写回")
    void calculatesWithUuidCodec() {
        List<ExcelNumber> excelList = ExcelSampleData.bookUuids();
        ExcelCalculateConfig config = ExcelSampleData.uuidConfig();
        BookCalculate<ExcelNumber> calc = new BookCalculate<>(excelList, config);
        String result = calc.calculate("交易性-股票",
                "N" + UUID_3 + "=现金!A" + UUID_1 + "*银行存款!C" + UUID_1
                        + "+ABS(MIN(O" + UUID_2 + "+2,银行存款!B" + UUID_3 + ":银行存款!E" + UUID_5
                        + ")*2-MAX('交易性-股票'!A" + UUID_1 + ":'交易性-股票'!D" + UUID_3
                        + "))+MAX(现金!A" + UUID_2 + ":现金!O" + UUID_3 + ")");
        calc.integrationResult();
        assertThat(result).isEqualTo("86504.9");
        assertThat(excelList.get(2).getN()).isEqualTo(result);
    }

    @Test
    @DisplayName("必须传入当前页签名")
    void requiresCurrentSheet() {
        BookCalculate<ExcelNumber> calc = new BookCalculate<>(ExcelSampleData.bookNumbers(), ExcelSampleData.numberConfig());
        assertThatThrownBy(() -> calc.calculate("A1=1"))
                .isInstanceOf(ExcelException.class)
                .hasMessageContaining("currentSheet");
    }
}
