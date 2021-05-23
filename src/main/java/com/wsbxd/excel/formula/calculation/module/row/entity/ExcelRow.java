package com.wsbxd.excel.formula.calculation.module.row.entity;

import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelEntity;
import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.cell.enums.ExcelCellTypeEnum;
import com.wsbxd.excel.formula.calculation.common.config.ExcelCalculateConfig;
import com.wsbxd.excel.formula.calculation.common.util.ExcelReflectUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * description: Excel 行
 *
 * @author chenhaoxuan
 * @date 2019/8/25
 */
public class ExcelRow<T> implements IExcelEntity<T> {

    private ExcelCalculateConfig properties;

    private Map<String, ExcelCell> columnAndCellListMap;

    private T excel;

    @Override
    public void integrationResult() {
        this.columnAndCellListMap.forEach((column, value) -> {
            if (value.getBaseValue() != ExcelReflectUtil.getValue(excel, column)) {
                ExcelReflectUtil.setValue(excel, column, value.getOriginalValue());
            }
        });
    }

    @Override
    public String getExcelCellValue(String cell) {
        if (columnAndCellListMap.containsKey(cell)) {
            return columnAndCellListMap.get(cell).getBaseValue();
        }
        return null;
    }

    @Override
    public ExcelCell getExcelCell(ExcelCell excelCell) {
        if (columnAndCellListMap.containsKey(excelCell.getColumn())) {
            return columnAndCellListMap.get(excelCell.getColumn());
        }
        return null;
    }

    @Override
    public String updateExcelCellValue(ExcelCell excelCell) {
        ExcelCell oldExcelCell = this.columnAndCellListMap.get(excelCell.getColumn());
        oldExcelCell.setBaseValue(excelCell.getOriginalValue());
        return oldExcelCell.getOriginalValue();
    }

    @Override
    public List<String> getExcelCellValueList(ExcelCell startCell, ExcelCell endCell) {
        return getExcelCellList(startCell, endCell)
                .stream().map(ExcelCell::getBaseValue).collect(Collectors.toList());
    }

    @Override
    public List<ExcelCell> getExcelCellList(ExcelCell startCell, ExcelCell endCell) {
        List<ExcelCell> excelCellList = new ArrayList<>();
        //idAndSortMap中未获取到的话默认获取边界
        ExcelUtil.getBetweenColumnList(startCell.getColumn(), endCell.getColumn()).forEach(column -> {
            //没有的列不用管
            if (columnAndCellListMap.containsKey(column)) {
                excelCellList.add(columnAndCellListMap.get(column));
            }
        });
        return excelCellList;
    }

    public ExcelRow(T excel, ExcelCalculateConfig properties) {
        this.excel = excel;
        this.properties = properties;
        //列和单元格类型
        Map<String, ExcelCellTypeEnum> columnAndType = ExcelUtil.parseColumnAndType(excel, properties);
        this.columnAndCellListMap = this.properties.getColumnFieldList().stream()
                .collect(Collectors.toMap(Field::getName, column -> new ExcelCell(column.getName(),
                        ExcelStrUtil.toString(ExcelReflectUtil.getValue(excel, properties.getIdField())), ExcelReflectUtil.getValue(this.excel, column), columnAndType.getOrDefault(column.getName(), ExcelCellTypeEnum.STRING))));
    }

    public ExcelCalculateConfig getProperties() {
        return properties;
    }

    public void setProperties(ExcelCalculateConfig properties) {
        this.properties = properties;
    }

    public Map<String, ExcelCell> getColumnAndCellListMap() {
        return columnAndCellListMap;
    }

    public void setColumnAndCellListMap(Map<String, ExcelCell> columnAndCellListMap) {
        this.columnAndCellListMap = columnAndCellListMap;
    }

    public T getExcel() {
        return excel;
    }

    public void setExcel(T excel) {
        this.excel = excel;
    }

    public ExcelRow(ExcelCalculateConfig properties, Map<String, ExcelCell> columnAndCellListMap, T excel) {
        this.properties = properties;
        this.columnAndCellListMap = columnAndCellListMap;
        this.excel = excel;
    }

    public ExcelRow() {
    }

    @Override
    public String toString() {
        return "ExcelRow{" +
                "properties=" + properties +
                ", columnAndCellListMap=" + columnAndCellListMap +
                ", excel=" + excel +
                '}';
    }
}
