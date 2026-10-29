package io.github.kanfengjingren.excelformula.common.util;

import io.github.kanfengjingren.excelformula.common.cell.entity.ExcelCell;
import io.github.kanfengjingren.excelformula.common.cell.enums.ExcelCellTypeEnum;
import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;

import java.lang.reflect.Field;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 从 POJO 行数据构建内存单元格映射。
 * <p>
 * 按配置中的列字段列表反射取值，并结合单元格类型字段生成列名到 ExcelCell 的映射，
 * 供单行、单页签、多页签实体构造时复用。
 */
public final class ExcelEntityBuilder {

    private ExcelEntityBuilder() {
    }

    /**
     * 将一行 POJO 转为「列名 → 单元格」映射。
     *
     * @param row       一行数据对象
     * @param config    计算配置（含列字段、主键字段、类型字段等）
     * @param sheetName 页签名称；多页签时传入，单行/单页签传 null
     * @param <T>       实体类型
     * @return 有序的列单元格映射
     */
    public static <T> Map<String, ExcelCell> buildColumnCells(T row, ExcelCalculateConfig config, String sheetName) {
        Map<String, ExcelCellTypeEnum> columnAndType = ExcelUtil.parseColumnAndType(row, config);
        String id = ExcelStrUtil.toString(ExcelReflectUtil.getValue(row, config.getIdField()));
        Map<String, ExcelCell> cells = new LinkedHashMap<>();
        for (Field field : config.getColumnFieldList()) {
            cells.put(field.getName(), new ExcelCell(
                    sheetName,
                    field.getName(),
                    id,
                    ExcelReflectUtil.getValue(row, field),
                    columnAndType.getOrDefault(field.getName(), ExcelCellTypeEnum.STRING)));
        }
        return cells;
    }
}
