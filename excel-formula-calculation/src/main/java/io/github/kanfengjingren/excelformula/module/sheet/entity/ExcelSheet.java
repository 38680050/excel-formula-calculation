package io.github.kanfengjingren.excelformula.module.sheet.entity;

import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCell;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.interfaces.IExcelEntity;
import io.github.kanfengjingren.excelformula.common.util.ExcelEntityBuilder;
import io.github.kanfengjingren.excelformula.common.util.ExcelRangeHelper;
import io.github.kanfengjingren.excelformula.common.util.ExcelReflectUtil;
import io.github.kanfengjingren.excelformula.common.util.ExcelStrUtil;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 单页签内存模型。
 * <p>
 * 按行主键组织「列→单元格」映射，并维护行排序范围，用于区域引用展开与结果写回。
 *
 * @param <T> 数据实体类型
 */
public class ExcelSheet<T> implements IExcelEntity<T> {

    private final ExcelCalculateConfig properties;

    /** 行主键到排序号 */
    private final Map<String, Integer> idAndSortMap;

    /** 行主键到列单元格映射 */
    private final Map<String, Map<String, ExcelCell>> idAndCellListMap;

    private final List<T> excelList;

    private final Integer maxSort;

    private final Integer minSort;

    /**
     * 从多行 POJO 构建页签索引与单元格映射。
     */
    public ExcelSheet(List<T> excelList, ExcelCalculateConfig properties) {
        this.excelList = excelList;
        this.properties = properties;
        List<Integer> excelSort = excelList.stream()
                .map(e -> (Integer) ExcelReflectUtil.getValue(e, this.properties.getSortField()))
                .collect(Collectors.toList());
        this.maxSort = Collections.max(excelSort);
        this.minSort = Collections.min(excelSort);
        this.idAndSortMap = excelList.stream().collect(Collectors.toMap(
                e -> ExcelStrUtil.toString(ExcelReflectUtil.getValue(e, this.properties.getIdField())),
                e -> ExcelReflectUtil.getValue(e, this.properties.getSortField())));
        this.idAndCellListMap = excelList.stream().collect(Collectors.toMap(
                e -> ExcelStrUtil.toString(ExcelReflectUtil.getValue(e, this.properties.getIdField())),
                e -> ExcelEntityBuilder.buildColumnCells(e, properties, null)));
    }

    /**
     * 按行将单元格原值写回对应 POJO 字段。
     */
    @Override
    public void integrationResult() {
        this.excelList.forEach(excel -> {
            String idValue = ExcelStrUtil.toString(ExcelReflectUtil.getValue(excel, this.properties.getIdField()));
            idAndCellListMap.get(idValue).forEach((column, cell) ->
                    ExcelReflectUtil.setValue(excel, column, cell.getOriginalValue()));
        });
    }

    /** 按引用字符串取单元格基础值；行或列不存在时返回 null。 */
    @Override
    public String getExcelCellValue(String cell) {
        ExcelCell excelCell = this.getExcelCell(new ExcelCell(cell, this.properties.getRowIdCodec(), null));
        return excelCell == null ? null : excelCell.getBaseValue();
    }

    /** 按行 ID 与列名在页签内定位单元格。 */
    @Override
    public ExcelCell getExcelCell(ExcelCell excelCell) {
        Map<String, ExcelCell> columnAndValue = this.idAndCellListMap.get(excelCell.getId());
        return columnAndValue == null ? null : columnAndValue.get(excelCell.getColumn());
    }

    /**
     * 更新指定行列单元格；行或列不存在时抛出异常。
     */
    @Override
    public String updateExcelCellValue(ExcelCell excelCell) {
        Map<String, ExcelCell> columnMap = this.idAndCellListMap.get(excelCell.getId());
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
     * 按起止单元格展开矩形区域内的单元格。
     */
    @Override
    public List<ExcelCell> getExcelCellList(ExcelCell startCell, ExcelCell endCell) {
        return ExcelRangeHelper.collectCells(idAndSortMap, idAndCellListMap,
                startCell.getId(), startCell.getColumn(), endCell.getId(), endCell.getColumn(), minSort, maxSort);
    }

    public ExcelCalculateConfig getProperties() {
        return properties;
    }

    public Map<String, Integer> getIdAndSortMap() {
        return idAndSortMap;
    }

    public Map<String, Map<String, ExcelCell>> getIdAndCellListMap() {
        return idAndCellListMap;
    }

    public List<T> getExcelList() {
        return excelList;
    }

    @Override
    public String toString() {
        return "ExcelSheet{idAndCellListMap=" + idAndCellListMap + ", excelList=" + excelList + '}';
    }
}
