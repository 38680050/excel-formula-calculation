package io.github.kanfengjingren.excelformula.common.reference;

import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCellCoordinate;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;
import io.github.kanfengjingren.excelformula.support.PrefixedNumericRowIdCodec;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/** 单元格地址字符串解析为页签名、列字母与行主键。 */
@DisplayName("单元格坐标 ExcelCellCoordinate")
class ExcelCellCoordinateTest {

    @Test
    @DisplayName("裸页签 + 数字行 ID")
    void parsesNumberCellWithSheet() {
        ExcelCellCoordinate c = new ExcelCellCoordinate("现金!A12", ExcelIdTypeEnum.NUMBER, null);
        assertThat(c.getSheet()).isEqualTo("现金");
        assertThat(c.getColumn()).isEqualTo("A");
        assertThat(c.getId()).isEqualTo("12");
    }

    @Test
    @DisplayName("引号页签 + UUID 行 ID")
    void parsesQuotedSheetAndUuid() {
        String uuid = "1bbd420f-3ce3-11ea-a676-00163e048143";
        ExcelCellCoordinate c = new ExcelCellCoordinate("'交易性-股票'!A" + uuid, ExcelIdTypeEnum.UUID, null);
        assertThat(c.getSheet()).isEqualTo("交易性-股票");
        assertThat(c.getColumn()).isEqualTo("A");
        assertThat(c.getId()).isEqualTo(uuid);
    }

    @Test
    @DisplayName("自定义 Codec 拆列与行 ID")
    void parsesWithCustomCodec() {
        ExcelCellCoordinate c = new ExcelCellCoordinate("BR99", new PrefixedNumericRowIdCodec(), null);
        assertThat(c.getColumn()).isEqualTo("B");
        assertThat(c.getId()).isEqualTo("R99");
    }
}
