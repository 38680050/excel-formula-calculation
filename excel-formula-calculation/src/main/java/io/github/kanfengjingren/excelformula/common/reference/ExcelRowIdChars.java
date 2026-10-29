package io.github.kanfengjingren.excelformula.common.reference;

/**
 * 行 ID / 引用扫描共用的字符判断，避免各 Codec 重复私有方法。
 */
final class ExcelRowIdChars {

    private ExcelRowIdChars() {
    }

    /** 是否为 ASCII 十进制数字 {@code 0-9}。 */
    static boolean isDigit(char c) {
        return c >= '0' && c <= '9';
    }

    /** 是否为大写拉丁字母 {@code A-Z}（列字母）。 */
    static boolean isUpperLetter(char c) {
        return c >= 'A' && c <= 'Z';
    }

    /** 是否为十六进制字符（大小写均可）。 */
    static boolean isHex(char c) {
        return isDigit(c)
                || (c >= 'a' && c <= 'f')
                || (c >= 'A' && c <= 'F');
    }

    /** 是否为小写十六进制字符（含数字 {@code 0-9}、字母 {@code a-f}）。 */
    static boolean isHexLower(char c) {
        return isDigit(c) || (c >= 'a' && c <= 'f');
    }

    /** Crockford Base32：0-9、A-Z，排除 I/L/O/U（大小写不敏感） */
    static boolean isCrockfordBase32(char c) {
        if (isDigit(c)) {
            return true;
        }
        char u = (c >= 'a' && c <= 'z') ? (char) (c - 32) : c;
        return u >= 'A' && u <= 'Z' && u != 'I' && u != 'L' && u != 'O' && u != 'U';
    }
}
