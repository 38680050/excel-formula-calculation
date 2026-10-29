package io.github.kanfengjingren.excelformula.common.cell.entity;

import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;
import io.github.kanfengjingren.excelformula.common.reference.ExcelFormulaReferenceScanner;
import io.github.kanfengjingren.excelformula.common.reference.ExcelFormulaReferenceScanner.LocalCellParts;
import io.github.kanfengjingren.excelformula.common.reference.ExcelRowIdCodec;
import io.github.kanfengjingren.excelformula.common.reference.ExcelRowIdCodecs;
import io.github.kanfengjingren.excelformula.common.util.ExcelStrUtil;

/**
 * 单元格在页签、列、行 ID 三维上的坐标，以及列与行拼接后的本地引用串。
 * 可从显式分量构造，或从带可选页签前缀的引用文本拆解。
 */
public class ExcelCellCoordinate {

    private String id;
    private String cell;
    private String sheet;
    private String column;

    /** 显式页签、列、行 ID；本地引用串为列名与行 ID 拼接。 */
    public ExcelCellCoordinate(String sheet, String column, String id) {
        this.id = id;
        this.sheet = sheet;
        this.column = column;
        this.cell = column + id;
    }

    /** 无页签：列与行 ID。 */
    public ExcelCellCoordinate(String column, String id) {
        this.id = id;
        this.column = column;
        this.cell = column + id;
    }

    /** 从引用串拆解，行 ID 类型由枚举决定。 */
    public ExcelCellCoordinate(String cell, ExcelIdTypeEnum excelIdTypeEnum) {
        this(cell, ExcelRowIdCodecs.fromIdType(excelIdTypeEnum), null);
    }

    /** 从引用串拆解，并可补全当前页签。 */
    public ExcelCellCoordinate(String cell, ExcelIdTypeEnum excelIdTypeEnum, String currentSheet) {
        this(cell, ExcelRowIdCodecs.fromIdType(excelIdTypeEnum), currentSheet);
    }

    /**
     * 从完整或本地引用串拆解页签、列与行 ID；行 ID 识别规则由 Codec 提供。
     */
    public ExcelCellCoordinate(String cell, ExcelRowIdCodec rowIdCodec, String currentSheet) {
        this.cell = cell;
        String[] sheetHolder = new String[1];
        String local = ExcelFormulaReferenceScanner.stripSheetPrefix(cell, sheetHolder);
        this.sheet = sheetHolder[0];
        if (ExcelStrUtil.isEmpty(this.sheet) && ExcelStrUtil.isNotEmpty(currentSheet)) {
            this.sheet = currentSheet;
        }
        LocalCellParts parts = ExcelFormulaReferenceScanner.parseLocalCell(local, rowIdCodec);
        if (parts != null) {
            this.column = parts.column();
            this.id = parts.rowId();
        }
    }

    /** 空构造。 */
    public ExcelCellCoordinate() {
    }

    /** 列名与行 ID 拼接的本地引用（不含页签前缀）。 */
    public String getCell() {
        return cell;
    }

    public void setCell(String cell) {
        this.cell = cell;
    }

    public String getSheet() {
        return sheet;
    }

    public void setSheet(String sheet) {
        this.sheet = sheet;
    }

    public String getColumn() {
        return column;
    }

    public void setColumn(String column) {
        this.column = column;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    @Override
    public String toString() {
        return "ExcelCellCoordinate{" +
                "cell='" + cell + '\'' +
                ", sheet='" + sheet + '\'' +
                ", column='" + column + '\'' +
                ", row='" + id + '\'' +
                '}';
    }
}
