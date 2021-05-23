package com.wsbxd.excel.formula.calculation.module.book.entity;

import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelEntity;
import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;
import com.wsbxd.excel.formula.calculation.common.cell.enums.ExcelCellTypeEnum;
import com.wsbxd.excel.formula.calculation.common.config.ExcelCalculateConfig;
import com.wsbxd.excel.formula.calculation.common.util.ExcelReflectUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelUtil;
import com.wsbxd.excel.formula.calculation.common.util.ExcelObjectUtil;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * description: Excel 工作簿
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelBook<T> implements IExcelEntity<T> {

    private ExcelCalculateConfig properties;

    private Map<String, Map<String, Map<String, ExcelCell>>> sheetAndSheetDataMap;

    private Map<String, Map<String, Integer>> sheetAndSheetDataSortMap;

    private Map<String, Integer> maxSheetAndSortMap;

    private Map<String, Integer> minSheetAndSortMap;

    private List<T> excelList;

    @Override
    public void integrationResult() {
        Map<String, T> idAndList = this.excelList.stream().collect(Collectors.toMap(e -> ExcelReflectUtil.getValue(e, this.properties.getIdField()), e -> e));
        this.sheetAndSheetDataMap.forEach((sheetName, excelSheet) -> {
            excelSheet.forEach((id, columnAndValue) -> {
                T t = idAndList.get(id);
                columnAndValue.forEach((column, excelCell) -> {
                    ExcelReflectUtil.setValue(t, column, excelCell.getOriginalValue());
                });
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
        if (this.sheetAndSheetDataMap.containsKey(excelCell.getSheet())) {
            Map<String, Map<String, ExcelCell>> idAndColumnMap = this.sheetAndSheetDataMap.get(excelCell.getSheet());
            if (idAndColumnMap.containsKey(excelCell.getId())) {
                Map<String, ExcelCell> columnAndCell = idAndColumnMap.get(excelCell.getId());
                if (columnAndCell.containsKey(excelCell.getColumn())) {
                    return columnAndCell.get(excelCell.getColumn());
                }
            }
        }
        return null;
    }

    @Override
    public String updateExcelCellValue(ExcelCell excelCell) {
        ExcelCell oldExcelCell = this.sheetAndSheetDataMap.get(excelCell.getSheet()).get(excelCell.getId()).get(excelCell.getColumn());
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
        //cell1:cell2必须在一个页签中
        if (!startCell.getSheet().equals(endCell.getSheet())) {
            throw new RuntimeException("startCell 与 endCell 不在同一个页签！");
        }
        return getExcelCellList(startCell.getSheet(), startCell.getId(), startCell.getColumn(), endCell.getId(), endCell.getColumn());
    }

    private List<ExcelCell> getExcelCellList(String sheet, String startRow, String startColumn, String endRow, String endColumn) {
        Map<String, Map<String, ExcelCell>> idAndCellListMap = sheetAndSheetDataMap.get(sheet);
        Map<String, Integer> idAndSortMap = sheetAndSheetDataSortMap.get(sheet);
        List<ExcelCell> excelCellList = new ArrayList<>();
        //idAndSortMap中未获取到的话默认获取边界
        Integer startRowNum = idAndSortMap.containsKey(startRow) ? idAndSortMap.get(startRow) : minSheetAndSortMap.get(sheet);
        Integer endRowNum = idAndSortMap.containsKey(endRow) ? idAndSortMap.get(endRow) : maxSheetAndSortMap.get(sheet);
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

    public ExcelBook(List<T> excelList, ExcelCalculateConfig properties) {
        this.excelList = excelList;
        this.properties = properties;
        Map<String, List<T>> nameAndDataListMap = ExcelObjectUtil.classifiedByField(this.excelList, this.properties.getSheetField());
        this.maxSheetAndSortMap = new HashMap<>();
        this.minSheetAndSortMap = new HashMap<>();
        nameAndDataListMap.forEach((sheet, dataList) -> {
            List<Integer> excelSort = dataList.stream().map(e -> (Integer) ExcelReflectUtil.getValue(e, properties.getSortField())).collect(Collectors.toList());
            maxSheetAndSortMap.put(sheet, Collections.max(excelSort));
            minSheetAndSortMap.put(sheet, Collections.min(excelSort));
        });
        this.sheetAndSheetDataSortMap = nameAndDataListMap.entrySet().stream().collect(Collectors.toMap(Map.Entry::getKey,
                e -> e.getValue().stream().collect(Collectors.toMap(e1 -> ExcelStrUtil.toString(ExcelReflectUtil.getValue(e1, this.properties.getIdField())),
                        e1 -> ExcelReflectUtil.getValue(e1, this.properties.getSortField())))));
        this.sheetAndSheetDataMap = nameAndDataListMap.entrySet().stream().collect(Collectors.toMap(Map.Entry::getKey,
                e -> e.getValue().stream().collect(Collectors.toMap(e1 -> ExcelStrUtil.toString(ExcelReflectUtil.getValue(e1, this.properties.getIdField())),
                        e1 -> {
                            Map<String, ExcelCell> cellStrAndCellMap = new HashMap<>();
                            Map<String, ExcelCellTypeEnum> columnAndType = ExcelUtil.parseColumnAndType(e1, properties);
                            for (Field field : properties.getColumnFieldList()) {
                                cellStrAndCellMap.put(field.getName(), new ExcelCell(ExcelReflectUtil.getValue(e1, properties.getSheetField()), field.getName(),
                                        ExcelStrUtil.toString(ExcelReflectUtil.getValue(e1, properties.getIdField())),
                                        ExcelReflectUtil.getValue(e1, field), columnAndType.getOrDefault(field.getName(), ExcelCellTypeEnum.STRING)));
                            }
                            return cellStrAndCellMap;
                        }))));
    }

    public ExcelCalculateConfig getProperties() {
        return properties;
    }

    public void setProperties(ExcelCalculateConfig properties) {
        this.properties = properties;
    }

    public Map<String, Map<String, Map<String, ExcelCell>>> getSheetAndSheetDataMap() {
        return sheetAndSheetDataMap;
    }

    public void setSheetAndSheetDataMap(Map<String, Map<String, Map<String, ExcelCell>>> sheetAndSheetDataMap) {
        this.sheetAndSheetDataMap = sheetAndSheetDataMap;
    }

    public Map<String, Map<String, Integer>> getSheetAndSheetDataSortMap() {
        return sheetAndSheetDataSortMap;
    }

    public void setSheetAndSheetDataSortMap(Map<String, Map<String, Integer>> sheetAndSheetDataSortMap) {
        this.sheetAndSheetDataSortMap = sheetAndSheetDataSortMap;
    }

    public List<T> getExcelList() {
        return excelList;
    }

    public void setExcelList(List<T> excelList) {
        this.excelList = excelList;
    }

    public ExcelBook() {
    }

    @Override
    public String toString() {
        return "ExcelBook{" +
                "properties=" + properties +
                ", sheetAndSheetDataMap=" + sheetAndSheetDataMap +
                ", sheetAndSheetDataSortMap=" + sheetAndSheetDataSortMap +
                ", excelList=" + excelList +
                '}';
    }
}
