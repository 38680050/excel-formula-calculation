package io.github.kanfengjingren.excelformula.common.field.enums;

/**
 * 内置行主键（嵌入单元格地址中的 ID 段）语义类型。
 * 非常规格式可在字段注解上指定自定义 idCodec，无需扩展本枚举。
 */
public enum ExcelIdTypeEnum {

    /**
     * 数字行号：连续十进制数字（如 {@code A1}、{@code B12}）
     */
    NUMBER,

    /**
     * 雪花 ID：可落入 {@code long} 的十进制数字（通常 18～19 位）
     */
    SNOWFLAKE,

    /**
     * 带横线 UUID：{@code 8-4-4-4-12} 小写十六进制
     */
    UUID,

    /**
     * 无横线 UUID：连续 32 位十六进制
     */
    UUID_COMPACT,

    /**
     * MongoDB ObjectId：连续 24 位十六进制
     */
    OBJECT_ID,

    /**
     * ULID：26 位 Crockford Base32
     */
    ULID,

    /**
     * 未指定（非 ID 字段默认值）
     */
    NONE
}
