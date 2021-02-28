package com.wsbxd.excel.formula.calculation.common.field.annotation;

import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelFieldTypeEnum;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelIdTypeEnum;

import java.lang.annotation.*;

/**
 * description: Excel 字段注解
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface ExcelField {

    /**
     * 实体类字段类型，默认为单元格类型
     *
     * @return 实体类字段类型
     */
    ExcelFieldTypeEnum value() default ExcelFieldTypeEnum.CELL;

    /**
     * 实体类唯一标识字段值类型，除唯一标识字段外均不需要此值，因此默认为无类型
     *
     * @return 实体类唯一标识字段值类型
     */
    ExcelIdTypeEnum idType() default ExcelIdTypeEnum.NONE;

}
