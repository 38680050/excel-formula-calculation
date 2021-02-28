package com.wsbxd.excel.formula.calculation.common.cell.entity;

import com.wsbxd.excel.formula.calculation.common.cell.enums.ExcelCellTypeEnum;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelIdTypeEnum;

/**
 * description: Excel 单元格
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/26 22:01
 */
public class ExcelCell {

    private ExcelCellCoordinate coordinate;

    private String originalValue;

    private String baseValue;

    private ExcelCellTypeEnum type;

    public String getOriginalValue() {
        return this.originalValue;
    }

    public String getBaseValue() {
        if (null == this.type) {
            throw new RuntimeException("单元格类型不能为空!");
        } else if (null == this.baseValue) {
            this.baseValue = this.type.compatibleValue(this.originalValue);
        }
        return this.baseValue;
    }

    public void setOriginalValue(String originalValue) {
        this.originalValue = originalValue;
    }

    public void setBaseValue(String baseValue) {
        if (null == this.type) {
            throw new RuntimeException("单元格类型不能为空!");
        }
        this.baseValue = baseValue;
        this.originalValue = this.type.convergenceValue(baseValue);
    }

    public String getColumn() {
        return this.coordinate.getColumn();
    }

    public String getId() {
        return this.coordinate.getId();
    }

    public String getSheet() {
        return this.coordinate.getSheet();
    }

    /**
     * 用于数据
     *
     * @param sheet         工作表
     * @param column        列
     * @param row           行(id)
     * @param originalValue 值
     * @param type          类型
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
     * 用于数据
     *
     * @param column        列
     * @param row           行(id)
     * @param originalValue 值
     * @param type          类型
     */
    public ExcelCell(String column, String row, String originalValue, ExcelCellTypeEnum type) {
        this(null, column, row, originalValue, type);
    }

    /**
     * 用于确定坐标
     *
     * @param cell            单元格
     * @param excelIdTypeEnum id类型
     */
    public ExcelCell(String cell, ExcelIdTypeEnum excelIdTypeEnum) {
        this(cell, excelIdTypeEnum, null);
    }

    /**
     * 用于确定坐标
     *
     * @param cell            单元格
     * @param excelIdTypeEnum id类型
     * @param currentSheet    当前工作表
     */
    public ExcelCell(String cell, ExcelIdTypeEnum excelIdTypeEnum, String currentSheet) {
        this.coordinate = new ExcelCellCoordinate(cell, excelIdTypeEnum, currentSheet);
    }

    public ExcelCellCoordinate getCoordinate() {
        return coordinate;
    }

    public void setCoordinate(ExcelCellCoordinate coordinate) {
        this.coordinate = coordinate;
    }

    public ExcelCellTypeEnum getType() {
        return type;
    }

    public void setType(ExcelCellTypeEnum type) {
        this.type = type;
    }

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
