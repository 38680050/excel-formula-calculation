package io.github.kanfengjingren.excelformula.common.reference;

/**
 * 带横线的 UUID 行 ID：{@code 8-4-4-4-12} 段，每段为小写十六进制，段间以 {@code -} 分隔。
 */
public final class UuidRowIdCodec implements ExcelRowIdCodec {

    /** 无状态单例，供枚举映射与默认配置使用。 */
    public static final UuidRowIdCodec INSTANCE = new UuidRowIdCodec();

    private static final int[] GROUP_LENGTHS = {8, 4, 4, 4, 12};

    private UuidRowIdCodec() {
    }

    /**
     * 从 {@code from} 起按 UUID 分段规则读取；成功返回结束下标，格式不符返回 -1。
     */
    @Override
    public int tryRead(String text, int from) {
        if (text == null || from < 0) {
            return -1;
        }
        int i = from;
        for (int g = 0; g < GROUP_LENGTHS.length; g++) {
            if (g > 0) {
                if (i >= text.length() || text.charAt(i) != '-') {
                    return -1;
                }
                i++;
            }
            int end = i + GROUP_LENGTHS[g];
            if (end > text.length()) {
                return -1;
            }
            for (; i < end; i++) {
                if (!ExcelRowIdChars.isHexLower(text.charAt(i))) {
                    return -1;
                }
            }
        }
        return i;
    }
}
