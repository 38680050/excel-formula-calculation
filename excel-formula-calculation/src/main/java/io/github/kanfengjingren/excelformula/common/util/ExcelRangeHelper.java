package io.github.kanfengjingren.excelformula.common.util;

import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCell;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 单元格区域展开工具。
 * <p>
 * 根据起止列或起止行列，从内存映射中收集矩形区域内存在的单元格。
 */
public final class ExcelRangeHelper {

    private ExcelRangeHelper() {
    }

    /**
     * 在单行（仅有列维度）中，收集起止列之间存在的单元格。
     *
     * @param columnMap   列名到单元格的映射
     * @param startColumn 起始列（含）
     * @param endColumn   结束列（含）
     * @return 区域内单元格列表（跳过不存在的列）
     */
    public static List<ExcelCell> collectColumns(Map<String, ExcelCell> columnMap, String startColumn, String endColumn) {
        List<ExcelCell> excelCellList = new ArrayList<>();
        for (String column : ExcelUtil.getBetweenColumnList(startColumn, endColumn)) {
            ExcelCell cell = columnMap.get(column);
            if (cell != null) {
                excelCellList.add(cell);
            }
        }
        return excelCellList;
    }

    /**
     * 在页签数据中，按排序号确定行范围，再按列范围收集矩形区域内的单元格。
     * 起止行若不在排序表中，则分别回退到最小/最大排序边界。
     *
     * @param idAndSortMap     行主键到排序号的映射
     * @param idAndCellListMap 行主键到「列→单元格」的映射
     * @param startRow         起始行主键
     * @param startColumn      起始列
     * @param endRow           结束行主键
     * @param endColumn        结束列
     * @param minSort          该页签最小排序号
     * @param maxSort          该页签最大排序号
     * @return 区域内单元格列表
     */
    public static List<ExcelCell> collectCells(Map<String, Integer> idAndSortMap,
                                               Map<String, Map<String, ExcelCell>> idAndCellListMap,
                                               String startRow,
                                               String startColumn,
                                               String endRow,
                                               String endColumn,
                                               Integer minSort,
                                               Integer maxSort) {
        List<ExcelCell> excelCellList = new ArrayList<>();
        Integer startRowNum = idAndSortMap.containsKey(startRow) ? idAndSortMap.get(startRow) : minSort;
        Integer endRowNum = idAndSortMap.containsKey(endRow) ? idAndSortMap.get(endRow) : maxSort;
        List<String> idList = idAndSortMap.entrySet().stream()
                .filter(e -> e.getValue() >= startRowNum && e.getValue() <= endRowNum)
                .map(Map.Entry::getKey)
                .collect(Collectors.toList());
        for (String id : idList) {
            Map<String, ExcelCell> columnMap = idAndCellListMap.get(id);
            if (columnMap == null) {
                continue;
            }
            excelCellList.addAll(collectColumns(columnMap, startColumn, endColumn));
        }
        return excelCellList;
    }
}
