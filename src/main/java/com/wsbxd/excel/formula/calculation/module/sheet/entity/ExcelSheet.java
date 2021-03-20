package com.wsbxd.excel.formula.calculation.module.sheet.entity;

import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelEntity;
import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.cell.enums.ExcelCellTypeEnum;
import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;
import com.wsbxd.excel.formula.calculation.common.prop.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.common.util.ExcelReflectUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * description: Excel 页签
 *
 * @author chenhaoxuan
 * @date 2019/8/25
 */
public class ExcelSheet<T> implements IExcelEntity<T> {

    private ExcelEntityProperties properties;

    private Map<String, Integer> idAndSortMap;

    private Map<String, Map<String, ExcelCell>> idAndCellListMap;

    private List<T> excelList;

    private Integer maxSort;

    private Integer minSort;

    @Override
    public void integrationResult() {
        this.excelList.forEach(excel -> {
            String idValue = ExcelStrUtil.toString(ExcelReflectUtil.getValue(excel, this.properties.getIdField()));
            idAndCellListMap.get(idValue).forEach((column, value) -> {
                if (value.getBaseValue() != ExcelReflectUtil.getValue(excel, column)) {
                    ExcelReflectUtil.setValue(excel, column, value.getOriginalValue());
                }
            });
        });
    }

    @Override
    public String getExcelCellValue(String cell) {
        ExcelCell excelCell = this.getExcelCell(new ExcelCell(cell, this.properties.getExcelIdTypeEnum()));
        if (null != excelCell) {
            return excelCell.getBaseValue();
        }
        return null;
    }

    @Override
    public ExcelCell getExcelCell(ExcelCell excelCell) {
        if (this.idAndCellListMap.containsKey(excelCell.getId())) {
            Map<String, ExcelCell> columnAndValue = this.idAndCellListMap.get(excelCell.getId());
            if (columnAndValue.containsKey(excelCell.getColumn())) {
                return columnAndValue.get(excelCell.getColumn());
            }
        }
        return null;
    }

    @Override
    public void updateExcelCellValue(ExcelCell excelCell) {
        this.idAndCellListMap.get(excelCell.getId()).get(excelCell.getColumn()).setBaseValue(excelCell.getOriginalValue());
    }

    @Override
    public List<String> getExcelCellValueList(ExcelCell startCell, ExcelCell endCell) {
        return getExcelCellList(startCell, endCell)
                .stream().map(ExcelCell::getBaseValue).collect(Collectors.toList());
    }

    @Override
    public List<ExcelCell> getExcelCellList(ExcelCell startCell, ExcelCell endCell) {
        return getExcelCellList(startCell.getId(), startCell.getColumn(), endCell.getId(), endCell.getColumn());
    }

    public ExcelCell getExcelCell(String row, String column) {
        return idAndCellListMap.get(row).get(column);
    }

    public String getExcelCellStrValue(String row, String column) {
        return idAndCellListMap.get(row).get(column).getBaseValue();
    }

    public List<ExcelCell> getExcelCellList(String startRow, String startColumn, String endRow, String endColumn) {
        List<ExcelCell> excelCellList = new ArrayList<>();
        //idAndSortMap中未获取到的话默认获取边界
        Integer startRowNum = idAndSortMap.containsKey(startRow) ? idAndSortMap.get(startRow) : minSort;
        Integer endRowNum = idAndSortMap.containsKey(endRow) ? idAndSortMap.get(endRow) : maxSort;
        List<String> idList = idAndSortMap.entrySet().stream().filter(e -> (e.getValue() >= startRowNum && e.getValue() <= endRowNum))
                .map(Map.Entry::getKey).collect(Collectors.toList());
        idList.forEach(id -> {
            ExcelUtil.getBetweenColumnList(startColumn, endColumn).forEach(column -> {
                //没有的列不用管
                if (idAndCellListMap.get(id).containsKey(column)) {
                    excelCellList.add(idAndCellListMap.get(id).get(column));
                }
            });
        });
        return excelCellList;
    }

    public ExcelSheet(List<T> excelList, ExcelEntityProperties properties) {
        this.excelList = excelList;
        this.properties = properties;
        List<Integer> excelSort = excelList.stream().map(e -> (Integer) ExcelReflectUtil.getValue(e, this.properties.getSortField())).collect(Collectors.toList());
        this.maxSort = Collections.max(excelSort);
        this.minSort = Collections.min(excelSort);
        this.idAndSortMap = excelList.stream().collect(Collectors.toMap(e -> ExcelStrUtil.toString(ExcelReflectUtil.getValue(e, this.properties.getIdField())),
                e -> ExcelReflectUtil.getValue(e, this.properties.getSortField())));
        this.idAndCellListMap = excelList.stream().collect(Collectors.toMap(e -> ExcelStrUtil.toString(ExcelReflectUtil.getValue(e, this.properties.getIdField())),
                e -> {
                    Map<String, ExcelCell> cellStrAndCellMap = new HashMap<>();
                    Map<String, ExcelCellTypeEnum> columnAndType = ExcelUtil.parseColumnAndType(e, properties);
                    for (Field field : properties.getColumnFieldList()) {
                        String sheet = null;
                        //只有工作簿计算才需要工作表名称
                        if (ExcelCalculateTypeEnum.BOOK.equals(properties.getCalculateType())) {
                            sheet = ExcelReflectUtil.getValue(e, properties.getSheetField());
                        }
                        cellStrAndCellMap.put(field.getName(), new ExcelCell(sheet, field.getName(),
                                ExcelStrUtil.toString(ExcelReflectUtil.getValue(e, properties.getIdField())),
                                ExcelReflectUtil.getValue(e, field), columnAndType.getOrDefault(field.getName(), ExcelCellTypeEnum.STRING)));
                    }
                    return cellStrAndCellMap;
                }));
    }

    public ExcelEntityProperties getProperties() {
        return properties;
    }

    public void setProperties(ExcelEntityProperties properties) {
        this.properties = properties;
    }

    public Map<String, Integer> getIdAndSortMap() {
        return idAndSortMap;
    }

    public void setIdAndSortMap(Map<String, Integer> idAndSortMap) {
        this.idAndSortMap = idAndSortMap;
    }

    public Map<String, Map<String, ExcelCell>> getIdAndCellListMap() {
        return idAndCellListMap;
    }

    public void setIdAndCellListMap(Map<String, Map<String, ExcelCell>> idAndCellListMap) {
        this.idAndCellListMap = idAndCellListMap;
    }

    public List<T> getExcelList() {
        return excelList;
    }

    public void setExcelList(List<T> excelList) {
        this.excelList = excelList;
    }

    public ExcelSheet(ExcelEntityProperties properties, Map<String, Integer> idAndSortMap, Map<String, Map<String, ExcelCell>> idAndCellListMap, List<T> excelList) {
        this.properties = properties;
        this.idAndSortMap = idAndSortMap;
        this.idAndCellListMap = idAndCellListMap;
        this.excelList = excelList;
    }

    public ExcelSheet() {
    }

    @Override
    public String toString() {
        return "ExcelSheet{" +
                "properties=" + properties +
                ", idAndSortMap=" + idAndSortMap +
                ", idAndCellListMap=" + idAndCellListMap +
                ", excelList=" + excelList +
                '}';
    }
}
