package io.github.kanfengjingren.excelformula.common.reference;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelCalculateTypeEnum;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/** 内置行 ID Codec 的识别规则及与公式引用解析的配合。 */
@DisplayName("行 ID Codec")
class ExcelRowIdCodecTest {

    @Nested
    @DisplayName("枚举映射")
    class EnumMapping {

        @Test
        void mapsBuiltinIdTypes() {
            assertThat(ExcelRowIdCodecs.fromIdType(ExcelIdTypeEnum.NUMBER)).isSameAs(NumberRowIdCodec.INSTANCE);
            assertThat(ExcelRowIdCodecs.fromIdType(ExcelIdTypeEnum.SNOWFLAKE)).isSameAs(SnowflakeRowIdCodec.INSTANCE);
            assertThat(ExcelRowIdCodecs.fromIdType(ExcelIdTypeEnum.UUID)).isSameAs(UuidRowIdCodec.INSTANCE);
            assertThat(ExcelRowIdCodecs.fromIdType(ExcelIdTypeEnum.UUID_COMPACT)).isSameAs(FixedLengthRowIdCodec.UUID_COMPACT);
            assertThat(ExcelRowIdCodecs.fromIdType(ExcelIdTypeEnum.OBJECT_ID)).isSameAs(FixedLengthRowIdCodec.OBJECT_ID);
            assertThat(ExcelRowIdCodecs.fromIdType(ExcelIdTypeEnum.ULID)).isSameAs(FixedLengthRowIdCodec.ULID);
            assertThat(ExcelRowIdCodecs.fromIdType(ExcelIdTypeEnum.NONE)).isNull();
        }
    }

    @Nested
    @DisplayName("Snowflake")
    class Snowflake {

        @Test
        void readsDigitsWithinLongRange() {
            String id = "1234567890123456789";
            assertThat(SnowflakeRowIdCodec.INSTANCE.tryRead(id, 0)).isEqualTo(id.length());
            assertThat(SnowflakeRowIdCodec.INSTANCE.matches(id)).isTrue();
            assertThat(SnowflakeRowIdCodec.INSTANCE.tryRead("99999999999999999999", 0)).isEqualTo(18);
            assertThat(SnowflakeRowIdCodec.INSTANCE.tryRead(String.valueOf(Long.MAX_VALUE), 0)).isEqualTo(19);
        }

        @Test
        void worksInSheetFormula() {
            String id = "1234567890123456789";
            assertThat(ExcelFormulaReferenceParser.of(ExcelCalculateTypeEnum.SHEET, SnowflakeRowIdCodec.INSTANCE)
                    .findCellRefs("A" + id + "+B2")).containsExactly("A" + id, "B2");
        }
    }

    @Nested
    @DisplayName("定长十六进制")
    class FixedHex {

        @Test
        void compactUuidIs32Hex() {
            String id = "1bbd420f3ce311eaa67600163e048143";
            assertThat(FixedLengthRowIdCodec.UUID_COMPACT.tryRead(id, 0)).isEqualTo(32);
            assertThat(FixedLengthRowIdCodec.UUID_COMPACT.matches(id)).isTrue();
            assertThat(FixedLengthRowIdCodec.UUID_COMPACT.matches("1bbd420f-3ce3-11ea-a676-00163e048143")).isFalse();
        }

        @Test
        void objectIdIs24Hex() {
            String id = "507f1f77bcf86cd799439011";
            assertThat(FixedLengthRowIdCodec.OBJECT_ID.tryRead(id, 0)).isEqualTo(24);
            assertThat(FixedLengthRowIdCodec.OBJECT_ID.matches(id.toUpperCase())).isTrue();
        }

        @Test
        void ulidIs26Crockford() {
            String id = "01ARZ3NDEKTSV4RRFFQ69G5FAV";
            assertThat(FixedLengthRowIdCodec.ULID.tryRead(id, 0)).isEqualTo(26);
            assertThat(FixedLengthRowIdCodec.ULID.matches(id.toLowerCase())).isTrue();
            assertThat(FixedLengthRowIdCodec.ULID.matches("01ARZ3NDEKTSV4RRFFQ69G5FAI")).isFalse();
        }
    }
}
