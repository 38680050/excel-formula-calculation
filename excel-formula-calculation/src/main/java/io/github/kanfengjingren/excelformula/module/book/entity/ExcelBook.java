package io.github.kanfengjingren.excelformula.module.book.entity;

import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCell;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.interfaces.IExcelEntity;
import io.github.kanfengjingren.excelformula.common.util.ExcelEntityBuilder;
import io.github.kanfengjingren.excelformula.common.util.ExcelObjectUtil;
import io.github.kanfengjingren.excelformula.common.util.ExcelRangeHelper;
import io.github.kanfengjingren.excelformula.common.util.ExcelReflectUtil;
import io.github.kanfengjingren.excelformula.common.util.ExcelStrUtil;

import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 多页签（工作簿）内存模型。
 * <p>
 * 按页签名称分组，每页签内再按行主键组织列单元格；维护各页签的排序边界，
 * 支持跨页签单元格读写、同页签区域展开与结果写回。
 *
 * @param <T> 数据实体类型
 */
public class ExcelBook<T> implements IExcelEntity<T> {

    private final ExcelCalculateConfig properties;

    /** 页签 → 行主键 → 列 → 单元格 */
    private final Map<String, Map<String, Map<String, ExcelCell>>> sheetAndSheetDataMap;

    /** 页签 → 行主键 → 排序号 */
    private final Map<String, Map<String, Integer>> sheetAndSheetDataSortMap;

    private final Map<String, Integer> maxSheetAndSortMap;

    private final Map<String, Integer> minSheetAndSortMap;

    private final List<T> excelList;

    /**
     * 按页签字段分组构建工作簿索引与单元格映射。
     */
    public ExcelBook(List<T> excelList, ExcelCalculateConfig properties) {
        this.excelList = excelList;
        this.properties = properties;
        Map<String, List<T>> nameAndDataListMap = ExcelObjectUtil.classifiedByField(this.excelList, this.properties.getSheetField());
        this.maxSheetAndSortMap = new HashMap<>();
        this.minSheetAndSortMap = new HashMap<>();
        nameAndDataListMap.forEach((sheet, dataList) -> {
            List<Integer> excelSort = dataList.stream()
                    .map(e -> (Integer) ExcelReflectUtil.getValue(e, properties.getSortField()))
                    .collect(Collectors.toList());
            maxSheetAndSortMap.put(sheet, Collections.max(excelSort));
            minSheetAndSortMap.put(sheet, Collections.min(excelSort));
        });
        this.sheetAndSheetDataSortMap = nameAndDataListMap.entrySet().stream().collect(Collectors.toMap(
                Map.Entry::getKey,
                e -> e.getValue().stream().collect(Collectors.toMap(
                        e1 -> ExcelStrUtil.toString(ExcelReflectUtil.getValue(e1, this.properties.getIdField())),
                        e1 -> ExcelReflectUtil.getValue(e1, this.properties.getSortField())))));
        this.sheetAndSheetDataMap = nameAndDataListMap.entrySet().stream().collect(Collectors.toMap(
                Map.Entry::getKey,
                e -> e.getValue().stream().collect(Collectors.toMap(
                        e1 -> ExcelStrUtil.toString(ExcelReflectUtil.getValue(e1, this.properties.getIdField())),
                        e1 -> ExcelEntityBuilder.buildColumnCells(
                                e1,
                                properties,
                                ExcelReflectUtil.getValue(e1, properties.getSheetField()))))));
    }

    /**
     * 将各页签单元格原值写回对应 POJO 字段。
     */
    @Override
    public void integrationResult() {
        Map<String, T> idAndList = this.excelList.stream().collect(Collectors.toMap(
                e -> ExcelReflectUtil.getValue(e, this.properties.getIdField()),
                e -> e));
        this.sheetAndSheetDataMap.forEach((sheetName, excelSheet) ->
                excelSheet.forEach((id, columnAndValue) -> {
                    T t = idAndList.get(id);
                    columnAndValue.forEach((column, excelCell) ->
                            ExcelReflectUtil.setValue(t, column, excelCell.getOriginalValue()));
                }));
    }

    /** 按引用字符串取单元格基础值；页签、行或列不存在时返回 null。 */
    @Override
    public String getExcelCellValue(String cell) {
        ExcelCell excelCell = this.getExcelCell(new ExcelCell(cell, this.properties.getRowIdCodec(), null));
        return excelCell == null ? null : excelCell.getBaseValue();
    }

    /** 按坐标在工作簿中定位单元格对象。 */
    @Override
    public ExcelCell getExcelCell(ExcelCell excelCell) {
        Map<String, Map<String, ExcelCell>> idAndColumnMap = this.sheetAndSheetDataMap.get(excelCell.getSheet());
        if (idAndColumnMap == null) {
            return null;
        }
        Map<String, ExcelCell> columnAndCell = idAndColumnMap.get(excelCell.getId());
        return columnAndCell == null ? null : columnAndCell.get(excelCell.getColumn());
    }

    /**
     * 更新指定页签、行、列上的单元格；任一维度缺失时抛出异常。
     */
    @Override
    public String updateExcelCellValue(ExcelCell excelCell) {
        Map<String, Map<String, ExcelCell>> idMap = this.sheetAndSheetDataMap.get(excelCell.getSheet());
        if (idMap == null) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_REF, "页签不存在: " + excelCell.getSheet());
        }
        Map<String, ExcelCell> columnMap = idMap.get(excelCell.getId());
        if (columnMap == null) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_REF, "行不存在: " + excelCell.getId());
        }
        ExcelCell oldExcelCell = columnMap.get(excelCell.getColumn());
        if (oldExcelCell == null) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_REF, "单元格不存在: " + excelCell.getColumn());
        }
        oldExcelCell.setBaseValue(excelCell.getOriginalValue());
        return oldExcelCell.getOriginalValue();
    }

    /**
     * 展开同页签内起止单元格形成的矩形区域；起止不在同一页签时抛出异常。
     */
    @Override
    public List<ExcelCell> getExcelCellList(ExcelCell startCell, ExcelCell endCell) {
        if (!startCell.getSheet().equals(endCell.getSheet())) {
            throw new ExcelException("startCell 与 endCell 不在同一个页签！");
        }
        String sheet = startCell.getSheet();
        return ExcelRangeHelper.collectCells(
                sheetAndSheetDataSortMap.get(sheet),
                sheetAndSheetDataMap.get(sheet),
                startCell.getId(), startCell.getColumn(),
                endCell.getId(), endCell.getColumn(),
                minSheetAndSortMap.get(sheet), maxSheetAndSortMap.get(sheet));
    }

    public ExcelCalculateConfig getProperties() {
        return properties;
    }

    public Map<String, Map<String, Map<String, ExcelCell>>> getSheetAndSheetDataMap() {
        return sheetAndSheetDataMap;
    }

    public Map<String, Map<String, Integer>> getSheetAndSheetDataSortMap() {
        return sheetAndSheetDataSortMap;
    }

    public List<T> getExcelList() {
        return excelList;
    }

    @Override
    public String toString() {
        return "ExcelBook{sheetAndSheetDataMap=" + sheetAndSheetDataMap + ", excelList=" + excelList + '}';
    }
}
