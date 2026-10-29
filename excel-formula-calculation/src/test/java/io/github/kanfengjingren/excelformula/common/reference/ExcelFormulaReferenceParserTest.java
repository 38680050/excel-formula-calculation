package io.github.kanfengjingren.excelformula.common.reference;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelCalculateTypeEnum;
import io.github.kanfengjingren.excelformula.support.PrefixedNumericRowIdCodec;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/** 按 BOOK / SHEET / ROW 粒度从公式中提取单元格引用与赋值前缀。 */
@DisplayName("公式引用解析 ExcelFormulaReferenceParser")
class ExcelFormulaReferenceParserTest {

    @Nested
    @DisplayName("SHEET")
    class Sheet {

        @Test
        @DisplayName("抽取单元格并忽略函数名")
        void findsCellsAndIgnoresFunctionName() {
            ExcelFormulaReferenceParser parser = ExcelFormulaReferenceParser.of(
                    ExcelCalculateTypeEnum.SHEET, NumberRowIdCodec.INSTANCE);
            assertThat(parser.findCellRefs("SUM(A1)+B2")).containsExactly("A1", "B2");
            assertThat(parser.findReturnPrefix("N3=A1+B2")).contains("N3=");
        }

        @Test
        @DisplayName("嵌入 UUID 行 ID")
        void findsEmbeddedUuid() {
            String uuid = "1bbd420f-3ce3-11ea-a676-00163e048143";
            ExcelFormulaReferenceParser parser = ExcelFormulaReferenceParser.of(
                    ExcelCalculateTypeEnum.SHEET, UuidRowIdCodec.INSTANCE);
            assertThat(parser.findCellRefs("A" + uuid + "+B" + uuid))
                    .containsExactly("A" + uuid, "B" + uuid);
            assertThat(parser.findReturnPrefix("N" + uuid + "=1")).contains("N" + uuid + "=");
        }
    }

    @Nested
    @DisplayName("BOOK")
    class Book {

        @Test
        @DisplayName("裸页签与引号页签")
        void quotedAndBareSheet() {
            ExcelFormulaReferenceParser parser = ExcelFormulaReferenceParser.of(
                    ExcelCalculateTypeEnum.BOOK, NumberRowIdCodec.INSTANCE);
            assertThat(parser.findCellRefs("现金!A1+'交易性-股票'!B2+C3"))
                    .containsExactly("现金!A1", "'交易性-股票'!B2", "C3");
        }
    }

    @Nested
    @DisplayName("ROW")
    class Row {

        @Test
        @DisplayName("列引用跳过函数名")
        void columnRefsSkipFunctionNames() {
            ExcelFormulaReferenceParser parser = ExcelFormulaReferenceParser.of(
                    ExcelCalculateTypeEnum.ROW, NumberRowIdCodec.INSTANCE);
            assertThat(parser.findCellRefs("A+ABS(MIN(O+2,B))")).containsExactly("A", "O", "B");
            assertThat(parser.findReturnPrefix("I=A+B")).contains("I=");
        }
    }

    @Nested
    @DisplayName("自定义 Codec")
    class CustomCodec {

        @Test
        @DisplayName("列字母与 R+数字 ID 回溯拆分")
        void prefixedNumericIds() {
            ExcelFormulaReferenceParser parser = ExcelFormulaReferenceParser.of(
                    ExcelCalculateTypeEnum.SHEET, new PrefixedNumericRowIdCodec());
            assertThat(parser.findCellRefs("AR12+BR3")).containsExactly("AR12", "BR3");
            assertThat(parser.findReturnPrefix("CR9=1")).contains("CR9=");
        }
    }
}
