package io.github.kanfengjingren.excelformula.common.interfaces;

import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCell;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Excel 内存数据实体接口。
 * <p>
 * 定义单元格读取、区域展开、赋值更新以及结果写回原 POJO 的能力；
 * 单行、单页签、多页签模型均实现本接口。
 *
 * @param <T> 原始数据实体类型
 */
public interface IExcelEntity<T> {

    /**
     * 将内存中单元格结果收敛写回原始数据对象。
     */
    void integrationResult();

    /**
     * 批量查询单元格字符串对应的基础值，组成映射表。
     *
     * @param cellStrList 单元格引用字符串列表
     * @return 引用字符串到基础值的映射
     */
    default Map<String, String> getCellStrAndValueMap(List<String> cellStrList) {
        HashMap<String, String> cellStrAndValue = new HashMap<>(16);
        cellStrList.forEach(cellStr -> cellStrAndValue.put(cellStr, getExcelCellValue(cellStr)));
        return cellStrAndValue;
    }

    /**
     * 按单元格引用字符串取基础值。
     */
    String getExcelCellValue(String cell);

    /**
     * 按坐标对象取单元格。
     */
    ExcelCell getExcelCell(ExcelCell excelCell);

    /**
     * 用传入单元格的值更新内存中对应单元格，并返回更新后的原值形式。
     */
    String updateExcelCellValue(ExcelCell excelCell);

    /**
     * 取起止单元格矩形区域内所有单元格的基础值列表。
     */
    default List<String> getExcelCellValueList(ExcelCell startCell, ExcelCell endCell) {
        return getExcelCellList(startCell, endCell).stream()
                .map(ExcelCell::getBaseValue)
                .collect(Collectors.toList());
    }

    /**
     * 取起止单元格矩形区域内所有单元格对象。
     */
    List<ExcelCell> getExcelCellList(ExcelCell startCell, ExcelCell endCell);

}
