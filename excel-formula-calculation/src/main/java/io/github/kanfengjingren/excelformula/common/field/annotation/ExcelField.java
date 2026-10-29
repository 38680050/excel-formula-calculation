package io.github.kanfengjingren.excelformula.common.field.annotation;

import io.github.kanfengjingren.excelformula.common.field.enums.ExcelFieldTypeEnum;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;
import io.github.kanfengjingren.excelformula.common.reference.ExcelRowIdCodec;

import java.lang.annotation.*;

/**
 * 标记实体类字段在 Excel 计算中的角色：主键、数据列、页签、排序、单元格类型等。
 * 主键字段可额外指定 idType 或 idCodec，用于解析公式中的行 ID 段。
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface ExcelField {

    /**
     * 字段角色，默认为普通数据列。
     */
    ExcelFieldTypeEnum value() default ExcelFieldTypeEnum.CELL;

    /**
     * 主键字段的行 ID 语义类型；非主键字段保持默认 NONE。
     * 若 idCodec 已指定有效实现类，则以 Codec 为准。
     */
    ExcelIdTypeEnum idType() default ExcelIdTypeEnum.NONE;

    /**
     * 自定义行 ID 解析 Codec；指定非占位类时优先于 idType 的内置映射。
     */
    Class<? extends ExcelRowIdCodec> idCodec() default ExcelRowIdCodec.None.class;

}
