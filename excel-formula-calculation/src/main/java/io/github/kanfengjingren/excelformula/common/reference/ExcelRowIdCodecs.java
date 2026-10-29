package io.github.kanfengjingren.excelformula.common.reference;

import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;

/**
 * 行 ID Codec 解析与映射工具。
 */
public final class ExcelRowIdCodecs {

    private ExcelRowIdCodecs() {
    }

    /**
     * 将内置 {@link ExcelIdTypeEnum} 映射为 Codec；NONE 或 null 返回 null。
     */
    public static ExcelRowIdCodec fromIdType(ExcelIdTypeEnum idType) {
        if (idType == null) {
            return null;
        }
        return switch (idType) {
            case NUMBER -> NumberRowIdCodec.INSTANCE;
            case SNOWFLAKE -> SnowflakeRowIdCodec.INSTANCE;
            case UUID -> UuidRowIdCodec.INSTANCE;
            case UUID_COMPACT -> FixedLengthRowIdCodec.UUID_COMPACT;
            case OBJECT_ID -> FixedLengthRowIdCodec.OBJECT_ID;
            case ULID -> FixedLengthRowIdCodec.ULID;
            case NONE -> null;
        };
    }

    /**
     * 优先使用自定义 Codec 类，否则按 idType 映射。
     */
    public static ExcelRowIdCodec resolve(Class<? extends ExcelRowIdCodec> codecClass, ExcelIdTypeEnum idType) {
        if (codecClass != null && codecClass != ExcelRowIdCodec.None.class) {
            return instantiate(codecClass);
        }
        return fromIdType(idType);
    }

    /**
     * 通过无参构造反射创建自定义 Codec 实例；失败时抛 {@link ExcelException}。
     */
    public static ExcelRowIdCodec instantiate(Class<? extends ExcelRowIdCodec> codecClass) {
        try {
            return codecClass.getDeclaredConstructor().newInstance();
        } catch (ReflectiveOperationException e) {
            throw new ExcelException("无法实例化行 ID Codec: " + codecClass.getName(), e);
        }
    }

    /**
     * Coordinate / Cell 兼容入口：枚举映射，缺失时抛配置错误。
     */
    public static ExcelRowIdCodec requireFromIdType(ExcelIdTypeEnum idType) {
        ExcelRowIdCodec codec = fromIdType(idType);
        if (codec == null) {
            throw new ExcelException("未配置有效的行 ID 类型: " + idType);
        }
        return codec;
    }
}
