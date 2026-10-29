package io.github.kanfengjingren.excelformula.common.reference;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelCalculateTypeEnum;

/**
 * 公式中单元格/列引用的字符扫描（不使用正则）。
 * <p>
 * 页签前缀支持 {@code '页签'!} 与 {@code 裸页签!} 两种形式；引号内页签禁止 {@code \ / ? * [ ]}，
 * 裸页签在上述之外还禁止括号、冒号、逗号、加减、比较/逻辑运算符、空格等。
 * 行 ID 段由 {@link ExcelRowIdCodec} 识别。
 * <ul>
 *   <li>BOOK：可选页签前缀 + 本地单元格（列 + 行 ID）</li>
 *   <li>SHEET：仅本地单元格</li>
 *   <li>ROW：仅列字母，且列后不能紧跟 {@code (}（避免误判函数名）</li>
 * </ul>
 */
public final class ExcelFormulaReferenceScanner {

    private ExcelFormulaReferenceScanner() {
    }

    /**
     * 本地单元格（无页签）解析结果：列 + 可选行 ID。
     */
    public record LocalCellParts(String column, String rowId) {
    }

    /** 列+行 ID 绝对匹配：column 为 [from, columnEnd)，行 ID 为 [columnEnd, idEnd) */
    private record AbsoluteCellMatch(int columnEnd, int idEnd) {
    }

    /**
     * 从 from 起读取连续大写列字母；成功返回结束下标，失败 -1。
     */
    public static int tryReadColumn(String text, int from) {
        if (text == null || from < 0 || from >= text.length()) {
            return -1;
        }
        int i = from;
        while (i < text.length() && ExcelRowIdChars.isUpperLetter(text.charAt(i))) {
            i++;
        }
        return i > from ? i : -1;
    }

    /**
     * 解析「列 + 行 ID」（带回溯）；仅有列时 rowId 为 null（如 ROW 区域 A:H）。
     *
     * @return 无法识别列时返回 null
     */
    public static LocalCellParts parseLocalCell(String local, ExcelRowIdCodec rowIdCodec) {
        if (local == null || local.isEmpty()) {
            return null;
        }
        AbsoluteCellMatch match = tryMatchAbsoluteCell(local, 0, rowIdCodec);
        if (match != null) {
            return new LocalCellParts(local.substring(0, match.columnEnd()),
                    local.substring(match.columnEnd(), match.idEnd()));
        }
        int colEnd = tryReadColumn(local, 0);
        return colEnd < 0 ? null : new LocalCellParts(local.substring(0, colEnd), null);
    }

    /**
     * 拆出引用中的页签前缀，返回剩余本地单元格串；无页签时返回原文。
     *
     * @param sheetHolder 长度 1 的数组，用于写回解析到的页签名（可为 null）
     */
    public static String stripSheetPrefix(String cellRef, String[] sheetHolder) {
        if (cellRef == null) {
            return null;
        }
        int afterQuoted = tryReadQuotedSheetBang(cellRef, 0);
        if (afterQuoted > 0) {
            if (sheetHolder != null && sheetHolder.length > 0) {
                sheetHolder[0] = cellRef.substring(1, afterQuoted - 2);
            }
            return cellRef.substring(afterQuoted);
        }
        int afterBare = tryReadBareSheetBang(cellRef, 0);
        if (afterBare > 0) {
            if (sheetHolder != null && sheetHolder.length > 0) {
                sheetHolder[0] = cellRef.substring(0, afterBare - 1);
            }
            return cellRef.substring(afterBare);
        }
        return cellRef;
    }

    /**
     * 尝试匹配「'页签'!」；成功返回感叹号后下标，失败 -1。
     */
    public static int tryReadQuotedSheetBang(String text, int from) {
        if (text == null || from < 0 || from >= text.length() || text.charAt(from) != '\'') {
            return -1;
        }
        int i = from + 1;
        while (i < text.length()) {
            char c = text.charAt(i);
            if (c == '\'') {
                if (i + 1 < text.length() && text.charAt(i + 1) == '!') {
                    return i + 2;
                }
                return -1;
            }
            if (isForbiddenInQuotedSheet(c)) {
                return -1;
            }
            i++;
        }
        return -1;
    }

    /**
     * 尝试匹配「裸页签!」；成功返回感叹号后下标，失败 -1。
     */
    public static int tryReadBareSheetBang(String text, int from) {
        if (text == null || from < 0 || from >= text.length()) {
            return -1;
        }
        int i = from;
        while (i < text.length()) {
            char c = text.charAt(i);
            if (c == '!') {
                return i > from ? i + 1 : -1;
            }
            if (isForbiddenInBareSheet(c)) {
                return -1;
            }
            i++;
        }
        return -1;
    }

    /**
     * 在 from 处尝试匹配一条完整引用；成功返回结束下标，失败 -1。
     */
    public static int tryMatchRefAt(String text, int from, ExcelCalculateTypeEnum type, ExcelRowIdCodec rowIdCodec) {
        if (text == null || from < 0 || from >= text.length() || type == null) {
            return -1;
        }
        if (type == ExcelCalculateTypeEnum.BOOK) {
            return tryMatchBookRefAt(text, from, rowIdCodec);
        }
        if (type == ExcelCalculateTypeEnum.SHEET) {
            return tryMatchCellAt(text, from, rowIdCodec);
        }
        if (type == ExcelCalculateTypeEnum.ROW) {
            return tryMatchColumnRefAt(text, from);
        }
        return -1;
    }

    /**
     * BOOK 引用：先尝试带页签前缀的单元格，再回退为无页签的本地单元格匹配。
     */
    private static int tryMatchBookRefAt(String text, int from, ExcelRowIdCodec rowIdCodec) {
        int afterQuoted = tryReadQuotedSheetBang(text, from);
        if (afterQuoted > from) {
            int cellEnd = tryMatchCellAt(text, afterQuoted, rowIdCodec);
            if (cellEnd > afterQuoted) {
                return cellEnd;
            }
        }
        int afterBare = tryReadBareSheetBang(text, from);
        if (afterBare > from) {
            int cellEnd = tryMatchCellAt(text, afterBare, rowIdCodec);
            if (cellEnd > afterBare) {
                return cellEnd;
            }
        }
        return tryMatchCellAt(text, from, rowIdCodec);
    }

    /**
     * 列 + 行 ID（带回溯），返回结束下标。
     */
    private static int tryMatchCellAt(String text, int from, ExcelRowIdCodec rowIdCodec) {
        AbsoluteCellMatch match = tryMatchAbsoluteCell(text, from, rowIdCodec);
        return match == null ? -1 : match.idEnd();
    }

    /**
     * 从 from 起匹配「列 + 行 ID」：对列字母做最短列前缀回溯，使行 ID 段由 Codec 最长合法读取。
     *
     * @return 列结束与行 ID 结束下标；无法匹配时 null
     */
    private static AbsoluteCellMatch tryMatchAbsoluteCell(String text, int from, ExcelRowIdCodec rowIdCodec) {
        if (rowIdCodec == null) {
            return null;
        }
        int colEnd = tryReadColumn(text, from);
        if (colEnd < 0) {
            return null;
        }
        for (int c = colEnd; c > from; c--) {
            int idEnd = rowIdCodec.tryRead(text, c);
            if (idEnd > c) {
                return new AbsoluteCellMatch(c, idEnd);
            }
        }
        return null;
    }

    /** ROW：仅列，且后面不能紧跟 '('（避免把函数名当列） */
    private static int tryMatchColumnRefAt(String text, int from) {
        int colEnd = tryReadColumn(text, from);
        if (colEnd < 0) {
            return -1;
        }
        if (colEnd < text.length() && text.charAt(colEnd) == '(') {
            return -1;
        }
        return colEnd;
    }

    /** 引号页签名中不允许出现的字符：{@code \ / ? * [ ]}。 */
    private static boolean isForbiddenInQuotedSheet(char c) {
        return c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']';
    }

    /** 裸页签名中不允许出现的字符：引号页签禁止集，外加运算符、括号、冒号、逗号、空格等。 */
    private static boolean isForbiddenInBareSheet(char c) {
        return isForbiddenInQuotedSheet(c)
                || c == '(' || c == ')' || c == ':' || c == ',' || c == '+' || c == '-'
                || c == '=' || c == '<' || c == '>' || c == '^' || c == '&' || c == ' ';
    }
}
