package io.github.kanfengjingren.excelformula.common.reference;

/**
 * 雪花 ID：十进制数字，且能落入 {@code long} 范围（常见 18～19 位）。
 */
public final class SnowflakeRowIdCodec implements ExcelRowIdCodec {

    /** 无状态单例，供枚举映射与默认配置使用。 */
    public static final SnowflakeRowIdCodec INSTANCE = new SnowflakeRowIdCodec();

    private static final int MAX_DIGITS = 19;

    private SnowflakeRowIdCodec() {
    }

    /**
     * 从 {@code from} 起读取十进制数字，累计值不超过 {@code long} 上限；至少一位成功，否则 -1。
     */
    @Override
    public int tryRead(String text, int from) {
        if (text == null || from < 0 || from >= text.length() || !ExcelRowIdChars.isDigit(text.charAt(from))) {
            return -1;
        }
        long value = 0L;
        int i = from;
        int digits = 0;
        while (i < text.length() && ExcelRowIdChars.isDigit(text.charAt(i)) && digits < MAX_DIGITS) {
            int d = text.charAt(i) - '0';
            if (value > (Long.MAX_VALUE - d) / 10) {
                break;
            }
            value = value * 10 + d;
            i++;
            digits++;
        }
        return digits > 0 ? i : -1;
    }
}
