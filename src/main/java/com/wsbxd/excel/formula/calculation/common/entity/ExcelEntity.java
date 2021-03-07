package com.wsbxd.excel.formula.calculation.common.entity;

import com.wsbxd.excel.formula.calculation.common.cell.entity.ExcelCell;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * description: Excel 实体类
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/3/7 16:38
 */
public interface ExcelEntity {

    /**
     * 收敛结果到原集合
     */
    void integrationResult();

    /**
     * 根据单元格字符串集合获取 Map<cellStr, value>
     *
     * @param cellStrList 单元格字符串集合
     * @return Map<cellStr, value>
     */
    default Map<String, String> getCellStrAndValueMap(List<String> cellStrList) {
        HashMap<String, String> cellStrAndValue = new HashMap<>(16);
        cellStrList.forEach(cellStr -> cellStrAndValue.put(cellStr, getExcelCellValue(cellStr)));
        return cellStrAndValue;
    }

    /**
     * 根据单元格字符串获取单元格字符串值
     *
     * @param cell 单元格字符串
     * @return 单元格字符串值
     */
    String getExcelCellValue(String cell);

    /**
     * 根据单元格获取单元格信息
     *
     * @param excelCell 单元格
     * @return 单元格信息
     */
    ExcelCell getExcelCell(ExcelCell excelCell);

    /**
     * 修改单元格值
     *
     * @param excelCell 单元格
     */
    void updateExcelCellValue(ExcelCell excelCell);

}
