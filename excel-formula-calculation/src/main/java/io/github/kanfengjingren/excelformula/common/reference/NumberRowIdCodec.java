package io.github.kanfengjingren.excelformula.common.reference;

/**
 * 数字行 ID：连续十进制数字。
 */
public final class NumberRowIdCodec implements ExcelRowIdCodec {

    /** 无状态单例，供枚举映射与默认配置使用。 */
    public static final NumberRowIdCodec INSTANCE = new NumberRowIdCodec();

    private NumberRowIdCodec() {
    }

    /**
     * 从 {@code from} 起读取最长连续十进制数字；至少一位成功，否则 -1。
     */
    @Override
    public int tryRead(String text, int from) {
        if (text == null || from < 0 || from >= text.length()) {
            return -1;
        }
        int i = from;
        while (i < text.length() && ExcelRowIdChars.isDigit(text.charAt(i))) {
            i++;
        }
        return i > from ? i : -1;
    }
}
