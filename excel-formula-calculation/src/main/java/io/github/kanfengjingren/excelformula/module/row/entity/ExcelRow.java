package io.github.kanfengjingren.excelformula.module.row.entity;

import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCell;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.interfaces.IExcelEntity;
import io.github.kanfengjingren.excelformula.common.util.ExcelEntityBuilder;
import io.github.kanfengjingren.excelformula.common.util.ExcelRangeHelper;
import io.github.kanfengjingren.excelformula.common.util.ExcelReflectUtil;

import java.util.List;
import java.util.Map;

/**
 * 单行内存模型。
 * <p>
 * 以列名为键保存单元格；支持按列区域取值、更新单元格，以及将结果写回原 POJO。
 *
 * @param <T> 数据实体类型
 */
public class ExcelRow<T> implements IExcelEntity<T> {

    /** 计算配置 */
    private final ExcelCalculateConfig properties;

    /** 列名到单元格的映射 */
    private final Map<String, ExcelCell> columnAndCellListMap;

    /** 原始行对象 */
    private final T excel;

    /**
     * 从单行 POJO 构建列单元格映射。
     */
    public ExcelRow(T excel, ExcelCalculateConfig properties) {
        this.excel = excel;
        this.properties = properties;
        this.columnAndCellListMap = ExcelEntityBuilder.buildColumnCells(excel, properties, null);
    }

    /**
     * 将所有单元格的原值形式写回 POJO 对应字段。
     */
    @Override
    public void integrationResult() {
        columnAndCellListMap.forEach((column, cell) ->
                ExcelReflectUtil.setValue(excel, column, cell.getOriginalValue()));
    }

    /** 按列名取单元格基础值；列不存在时返回 null。 */
    @Override
    public String getExcelCellValue(String cell) {
        ExcelCell excelCell = columnAndCellListMap.get(cell);
        return excelCell == null ? null : excelCell.getBaseValue();
    }

    /** 按列名取本行单元格对象。 */
    @Override
    public ExcelCell getExcelCell(ExcelCell excelCell) {
        return columnAndCellListMap.get(excelCell.getColumn());
    }

    /**
     * 更新指定列单元格的基础值，并返回转换后的原值形式。
     */
    @Override
    public String updateExcelCellValue(ExcelCell excelCell) {
        ExcelCell oldExcelCell = this.columnAndCellListMap.get(excelCell.getColumn());
        if (oldExcelCell == null) {
            throw new ExcelException(ExcelConstant.EXCEL_ERROR_REF, "单元格不存在: " + excelCell.getColumn());
        }
        oldExcelCell.setBaseValue(excelCell.getOriginalValue());
        return oldExcelCell.getOriginalValue();
    }

    /**
     * 收集起止列之间存在的单元格。
     */
    @Override
    public List<ExcelCell> getExcelCellList(ExcelCell startCell, ExcelCell endCell) {
        return ExcelRangeHelper.collectColumns(columnAndCellListMap, startCell.getColumn(), endCell.getColumn());
    }

    public ExcelCalculateConfig getProperties() {
        return properties;
    }

    public Map<String, ExcelCell> getColumnAndCellListMap() {
        return columnAndCellListMap;
    }

    public T getExcel() {
        return excel;
    }

    @Override
    public String toString() {
        return "ExcelRow{columnAndCellListMap=" + columnAndCellListMap + ", excel=" + excel + '}';
    }
}
