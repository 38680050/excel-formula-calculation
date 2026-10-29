package io.github.kanfengjingren.excelformula.common.cell.entity;

import io.github.kanfengjingren.excelformula.common.cell.enums.ExcelCellTypeEnum;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;
import io.github.kanfengjingren.excelformula.common.reference.ExcelRowIdCodec;

/**
 * 单个 Excel 单元格的运行时表示：坐标、原值、按类型规范化后的基础值及单元格类型。
 * 支持从显式坐标构造，或从引用字符串解析坐标。
 */
public class ExcelCell {

    /** 页签、列、行 ID 坐标 */
    private ExcelCellCoordinate coordinate;

    /** 展示/存储用的原值字符串 */
    private String originalValue;

    /** 参与运算的基础值（懒加载或由 setBaseValue 维护） */
    private String baseValue;

    /** 单元格数据类型，决定原值与基础值互转规则 */
    private ExcelCellTypeEnum type;

    /**
     * 获取单元格原值
     *
     * @return 单元格原值
     */
    public String getOriginalValue() {
        return this.originalValue;
    }

    /**
     * 获取单元格基础值；未缓存时按类型从原值转换。
     *
     * @return 单元格基础值
     */
    public String getBaseValue() {
        if (null == this.type) {
            throw new ExcelException("单元格类型不能为空!");
        } else if (null == this.baseValue) {
            this.baseValue = this.type.compatibleValue(this.originalValue);
        }
        return this.baseValue;
    }

    /**
     * 设置单元格原值（不自动刷新基础值缓存）。
     */
    public void setOriginalValue(String originalValue) {
        this.originalValue = originalValue;
    }

    /**
     * 设置基础值，并按类型回写对应的原值形式。
     */
    public void setBaseValue(String baseValue) {
        if (null == this.type) {
            throw new ExcelException("单元格类型不能为空!");
        }
        this.baseValue = baseValue;
        this.originalValue = this.type.convergenceValue(baseValue);
    }

    /** 列名（如 A、B）。 */
    public String getColumn() {
        return this.coordinate.getColumn();
    }

    /** 行主键 ID。 */
    public String getId() {
        return this.coordinate.getId();
    }

    /** 页签名称。 */
    public String getSheet() {
        return this.coordinate.getSheet();
    }

    /**
     * 指定页签、列、行与初值构造单元格；类型为空时默认为字符串。
     */
    public ExcelCell(String sheet, String column, String row, String originalValue, ExcelCellTypeEnum type) {
        this.coordinate = new ExcelCellCoordinate(sheet, column, row);
        this.originalValue = originalValue;
        if (null == type) {
            this.type = ExcelCellTypeEnum.STRING;
        } else {
            this.type = type;
        }
    }

    /**
     * 无页签场景：列、行与初值构造单元格。
     */
    public ExcelCell(String column, String row, String originalValue, ExcelCellTypeEnum type) {
        this(null, column, row, originalValue, type);
    }

    /**
     * 从单元格引用字符串解析坐标，行 ID 类型由枚举指定。
     */
    public ExcelCell(String cell, ExcelIdTypeEnum excelIdTypeEnum) {
        this(cell, excelIdTypeEnum, null);
    }

    /**
     * 从引用字符串解析坐标；引用未带页签时使用 currentSheet 补全。
     */
    public ExcelCell(String cell, ExcelIdTypeEnum excelIdTypeEnum, String currentSheet) {
        this.coordinate = new ExcelCellCoordinate(cell, excelIdTypeEnum, currentSheet);
    }

    /**
     * 从引用字符串解析坐标，行 ID 段由自定义 Codec 识别。
     */
    public ExcelCell(String cell, ExcelRowIdCodec rowIdCodec, String currentSheet) {
        this.coordinate = new ExcelCellCoordinate(cell, rowIdCodec, currentSheet);
    }

    /** 坐标对象。 */
    public ExcelCellCoordinate getCoordinate() {
        return coordinate;
    }

    /** 替换坐标；类型与原值不变。 */
    public void setCoordinate(ExcelCellCoordinate coordinate) {
        this.coordinate = coordinate;
    }

    /** 单元格类型。 */
    public ExcelCellTypeEnum getType() {
        return type;
    }

    /** 设置类型；已存在的基础值不会自动重算。 */
    public void setType(ExcelCellTypeEnum type) {
        this.type = type;
    }

    /** 空构造，供框架或反序列化使用。 */
    public ExcelCell() {
    }

    @Override
    public String toString() {
        return "ExcelCell{" +
                "coordinate=" + coordinate +
                ", originalValue='" + originalValue + '\'' +
                ", type=" + type +
                '}';
    }
}
