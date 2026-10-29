package io.github.kanfengjingren.excelformula.support;

import io.github.kanfengjingren.excelformula.common.reference.ExcelRowIdCodec;

/**
 * 测试用自定义行 ID：{@code R} + 数字，如 {@code R12}。
 */
public final class PrefixedNumericRowIdCodec implements ExcelRowIdCodec {

    /**
     * 从 {@code from} 起匹配 {@code R} 后接至少一位数字；成功返回 ID 结束下标，否则 -1。
     */
    @Override
    public int tryRead(String text, int from) {
        if (text == null || from >= text.length() || text.charAt(from) != 'R') {
            return -1;
        }
        int i = from + 1;
        if (i >= text.length() || text.charAt(i) < '0' || text.charAt(i) > '9') {
            return -1;
        }
        while (i < text.length() && text.charAt(i) >= '0' && text.charAt(i) <= '9') {
            i++;
        }
        return i;
    }
}
