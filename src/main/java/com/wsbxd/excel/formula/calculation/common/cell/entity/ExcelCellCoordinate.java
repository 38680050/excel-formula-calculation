package com.wsbxd.excel.formula.calculation.common.cell.entity;

import com.wsbxd.excel.formula.calculation.common.constant.ExcelConstant;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelIdTypeEnum;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * description: Excel 单元格坐标
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/26 22:01
 */
public class ExcelCellCoordinate {

    public final static Pattern CELL = Pattern.compile("[A-Z]+\\d+");

    public final static Pattern COLUMN = Pattern.compile("^[A-Z]+");

    public final static Pattern ROW = Pattern.compile("\\d+$");

    public static final Pattern UUID = Pattern.compile("[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$");

    private String id;

    private String cell;

    private String sheet;

    private String column;

    public ExcelCellCoordinate(String sheet, String column, String id) {
        this.id = id;
        this.sheet = sheet;
        this.column = column;
        this.cell = column + id;
    }

    public ExcelCellCoordinate(String column, String id) {
        this.id = id;
        this.column = column;
        this.cell = column + id;
    }

    /**
     * 用户确定坐标
     *
     * @param cell            单元格
     * @param excelIdTypeEnum id类型
     */
    public ExcelCellCoordinate(String cell, ExcelIdTypeEnum excelIdTypeEnum) {
        this(cell, excelIdTypeEnum, null);
    }

    /**
     * 用户确定坐标
     *
     * @param cell            单元格
     * @param excelIdTypeEnum id类型
     * @param currentSheet    当前工作表
     */
    public ExcelCellCoordinate(String cell, ExcelIdTypeEnum excelIdTypeEnum, String currentSheet) {
        this.cell = cell;
        //页签名获取
        if (cell.startsWith(ExcelConstant.APOSTROPHE)) {
            this.sheet = cell.substring(1, cell.indexOf(ExcelConstant.APOSTROPHE, 1));
            cell = cell.substring(cell.indexOf(ExcelConstant.APOSTROPHE, 1) + 2);
        } else if (cell.contains(ExcelConstant.EXCLAMATION_MARK)) {
            String[] sheetAndCell = cell.split(ExcelConstant.EXCLAMATION_MARK);
            this.sheet = sheetAndCell[0];
            cell = sheetAndCell[1];
        }
        //如果没获取到工作表名,则使用currentSheet,页签可以为空字符串
        if (ExcelStrUtil.isEmpty(this.sheet) && ExcelStrUtil.isNotEmpty(currentSheet)) {
            this.sheet = currentSheet;
        }
        //单元格获取
        Matcher matcherColumn = COLUMN.matcher(cell);
        if (matcherColumn.find()) {
            this.column = matcherColumn.group();
        }
        Matcher matcherRow = null;
        if (ExcelIdTypeEnum.NUMBER.equals(excelIdTypeEnum)) {
            matcherRow = ROW.matcher(cell);
        } else if (ExcelIdTypeEnum.UUID.equals(excelIdTypeEnum)) {
            matcherRow = UUID.matcher(cell);
        }
        if (matcherRow != null && matcherRow.find()) {
            this.id = matcherRow.group();
        }
    }

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

    public ExcelCellCoordinate() {
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
