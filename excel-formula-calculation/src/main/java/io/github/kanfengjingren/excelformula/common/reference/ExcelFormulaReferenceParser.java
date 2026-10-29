package io.github.kanfengjingren.excelformula.common.reference;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelCalculateTypeEnum;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * 按计算类型与行 ID Codec 从公式中抽取单元格/列引用或赋值前缀。
 */
public final class ExcelFormulaReferenceParser {

    private final ExcelCalculateTypeEnum calculateType;
    private final ExcelRowIdCodec rowIdCodec;

    private ExcelFormulaReferenceParser(ExcelCalculateTypeEnum calculateType, ExcelRowIdCodec rowIdCodec) {
        this.calculateType = calculateType;
        this.rowIdCodec = rowIdCodec;
    }

    /**
     * 创建解析器：绑定计算类型（BOOK/SHEET/ROW）与行 ID 识别策略。
     */
    public static ExcelFormulaReferenceParser of(ExcelCalculateTypeEnum type, ExcelRowIdCodec rowIdCodec) {
        return new ExcelFormulaReferenceParser(type, rowIdCodec);
    }

    /**
     * 自左向右扫描公式，收集全部引用子串（不重叠，优先最长匹配）。
     */
    public List<String> findCellRefs(String formula) {
        List<String> refs = new ArrayList<>();
        if (formula == null || formula.isEmpty()) {
            return refs;
        }
        int i = 0;
        while (i < formula.length()) {
            int end = ExcelFormulaReferenceScanner.tryMatchRefAt(formula, i, calculateType, rowIdCodec);
            if (end > i) {
                refs.add(formula.substring(i, end));
                i = end;
            } else {
                i++;
            }
        }
        return refs;
    }

    /**
     * 若公式以「完整引用=」开头，返回含等号的前缀；否则 empty。
     */
    public Optional<String> findReturnPrefix(String formula) {
        if (formula == null || formula.isEmpty()) {
            return Optional.empty();
        }
        int end = ExcelFormulaReferenceScanner.tryMatchRefAt(formula, 0, calculateType, rowIdCodec);
        if (end > 0 && end < formula.length() && formula.startsWith(ExcelConstant.EQUALS_SIGN, end)) {
            return Optional.of(formula.substring(0, end + 1));
        }
        return Optional.empty();
    }

    /** 当前解析器使用的计算类型（决定引用形态）。 */
    public ExcelCalculateTypeEnum getCalculateType() {
        return calculateType;
    }

    /** 当前解析器用于单元格地址中行 ID 段的 Codec。 */
    public ExcelRowIdCodec getRowIdCodec() {
        return rowIdCodec;
    }
}
