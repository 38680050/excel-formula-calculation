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

    /**
     * 坐标
     */
    private ExcelCellCoordinate coordinate;

    /**
     * 原值
     */
    private String originalValue;

    /**
     * 基础值
     */
    private String baseValue;

    /**
     * 单元格类型枚举
     */
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
     * 获取单元格基础值
     *
     * @return 单元格基础值
     */
    public String getBaseValue() {
        if (null == this.type) {
            throw new RuntimeException("单元格类型不能为空!");
        } else if (null == this.baseValue) {
            this.baseValue = this.type.compatibleValue(this.originalValue);
        }
        return this.baseValue;
    }

    /**
     * 设置单元格基础值
     *
     * @param originalValue 单元格基础值
     */
    public void setOriginalValue(String originalValue) {
        this.originalValue = originalValue;
    }

    /**
     * 设置单元格基础值
     *
     * @param baseValue 单元格基础值
     */
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
