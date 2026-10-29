package io.github.kanfengjingren.excelformula.common.reference;

/**
 * 定长字符类行 ID：固定长度内每位字符需通过给定字符谓词校验。
 */
public final class FixedLengthRowIdCodec implements ExcelRowIdCodec {

    /** 判断单个字符是否允许出现在该定长 ID 中。 */
    @FunctionalInterface
    interface CharPredicate {
        boolean test(char c);
    }

    /** 32 位连续十六进制（无横线 UUID）。 */
    public static final FixedLengthRowIdCodec UUID_COMPACT =
            new FixedLengthRowIdCodec(32, ExcelRowIdChars::isHex);

    /** 24 位连续十六进制（ObjectId 形态）。 */
    public static final FixedLengthRowIdCodec OBJECT_ID =
            new FixedLengthRowIdCodec(24, ExcelRowIdChars::isHex);

    /** 26 位 Crockford Base32（ULID 形态）。 */
    public static final FixedLengthRowIdCodec ULID =
            new FixedLengthRowIdCodec(26, ExcelRowIdChars::isCrockfordBase32);

    private final int length;
    private final CharPredicate allowed;

    private FixedLengthRowIdCodec(int length, CharPredicate allowed) {
        this.length = length;
        this.allowed = allowed;
    }

    /**
     * 从 {@code from} 起读取恰好 {@code length} 个允许字符；不足或含非法字符时返回 -1。
     */
    @Override
    public int tryRead(String text, int from) {
        if (text == null || from < 0 || from + length > text.length()) {
            return -1;
        }
        for (int i = from; i < from + length; i++) {
            if (!allowed.test(text.charAt(i))) {
                return -1;
            }
        }
        return from + length;
    }
}
