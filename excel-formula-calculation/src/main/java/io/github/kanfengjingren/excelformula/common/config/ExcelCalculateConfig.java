package io.github.kanfengjingren.excelformula.common.config;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelCalculateTypeEnum;
import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.field.annotation.ExcelField;
import io.github.kanfengjingren.excelformula.common.field.enums.ExcelIdTypeEnum;
import io.github.kanfengjingren.excelformula.common.function.ExcelFunctionHandler;
import io.github.kanfengjingren.excelformula.common.function.FunctionRegistry;
import io.github.kanfengjingren.excelformula.common.interfaces.IFunction;
import io.github.kanfengjingren.excelformula.common.reference.ExcelFormulaReferenceParser;
import io.github.kanfengjingren.excelformula.common.reference.ExcelRowIdCodec;
import io.github.kanfengjingren.excelformula.common.reference.ExcelRowIdCodecs;
import io.github.kanfengjingren.excelformula.common.util.ExcelStrUtil;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Excel 计算配置。
 * <p>
 * 保存实体字段元数据、函数注册表、行 ID Codec、错误处理模式与引用解析器。
 * 实例创建后不可变；通过 forType / with* 生成快照。
 * <p>
 * 注解实体请用 {@link #of(Class)}：默认内置函数，再用 {@code function}/{@code functions} 叠加自定义（同名后覆盖先）。
 */
public class ExcelCalculateConfig {

    private final Field idField;
    private final ExcelIdTypeEnum excelIdTypeEnum;
    private final ExcelRowIdCodec rowIdCodec;
    private final Field cellTypesField;
    private final Field sheetField;
    private final Field sortField;
    private final List<Field> columnFieldList;
    private final ExcelCalculateTypeEnum calculateType;
    private final FunctionRegistry functionRegistry;
    private final ExcelErrorHandleMode errorHandleMode;
    private final ExcelFormulaReferenceParser referenceParser;

    /**
     * 按实体注解扫描字段，进入 Builder（已含内置函数）。
     */
    public static Builder of(Class<?> entityClass) {
        return new Builder(scanEntity(entityClass));
    }

    /**
     * 注解实体 + 仅内置函数，等价于 {@code of(clazz).build()}。
     */
    public static ExcelCalculateConfig create(Class<?> entityClass) {
        return of(entityClass).build();
    }

    /**
     * 按函数名调用已注册函数。
     */
    public String functionCalculate(String functionName, List<String> valueList) {
        return functionRegistry.invoke(functionName, valueList, this.errorHandleMode);
    }

    /**
     * 生成绑定指定计算类型的配置快照。
     */
    public ExcelCalculateConfig forType(ExcelCalculateTypeEnum type) {
        if (type == null) {
            throw new ExcelException("请设置 Excel 计算类型");
        }
        if (type.equals(this.calculateType)) {
            return this;
        }
        return copy(type, this.rowIdCodec, this.errorHandleMode);
    }

    /**
     * 生成指定错误处理模式的配置快照。
     */
    public ExcelCalculateConfig withErrorHandleMode(ExcelErrorHandleMode mode) {
        ExcelErrorHandleMode effective = ExcelErrorHandleMode.orThrow(mode);
        if (effective.equals(this.errorHandleMode)) {
            return this;
        }
        return copy(this.calculateType, this.rowIdCodec, effective);
    }

    /**
     * 生成指定行 ID Codec 的配置快照。
     */
    public ExcelCalculateConfig withRowIdCodec(ExcelRowIdCodec codec) {
        if (codec == this.rowIdCodec || (codec != null && codec.equals(this.rowIdCodec))) {
            return this;
        }
        return copy(this.calculateType, codec, this.errorHandleMode);
    }

    private ExcelCalculateConfig copy(ExcelCalculateTypeEnum type, ExcelRowIdCodec codec, ExcelErrorHandleMode mode) {
        return new ExcelCalculateConfig(idField, excelIdTypeEnum, codec, cellTypesField, sheetField, sortField,
                columnFieldList, type, functionRegistry, mode);
    }

    /**
     * 按字段名手动指定主键、页签、排序与列（无注解场景），仅内置函数。
     * 需要自定义函数时请给实体加 {@code @ExcelField} 并使用 {@link #of(Class)}。
     */
    public ExcelCalculateConfig(Class<?> clazz, String idFieldName, ExcelIdTypeEnum excelIdTypeEnum,
                                String sheetFieldName, String sortFieldName, List<String> columnFieldNameList) {
        this(clazz, idFieldName, excelIdTypeEnum, null, sheetFieldName, sortFieldName, columnFieldNameList);
    }

    /**
     * 按字段名手动配置，可指定行 ID Codec；仅内置函数。
     */
    public ExcelCalculateConfig(Class<?> clazz, String idFieldName, ExcelIdTypeEnum excelIdTypeEnum,
                                ExcelRowIdCodec rowIdCodec,
                                String sheetFieldName, String sortFieldName, List<String> columnFieldNameList) {
        try {
            this.idField = accessible(clazz.getDeclaredField(idFieldName));
            this.excelIdTypeEnum = excelIdTypeEnum;
            this.rowIdCodec = rowIdCodec != null ? rowIdCodec : ExcelRowIdCodecs.fromIdType(excelIdTypeEnum);
            this.sortField = accessible(clazz.getDeclaredField(sortFieldName));
            this.sheetField = ExcelStrUtil.isNotBlank(sheetFieldName)
                    ? accessible(clazz.getDeclaredField(sheetFieldName)) : null;
            this.cellTypesField = null;
            List<Field> columns = new ArrayList<>();
            for (String column : columnFieldNameList) {
                columns.add(accessible(clazz.getDeclaredField(column)));
            }
            this.columnFieldList = Collections.unmodifiableList(columns);
        } catch (NoSuchFieldException e) {
            throw new ExcelException("创建计算配置失败: " + e.getMessage(), e);
        }
        this.calculateType = null;
        this.functionRegistry = FunctionRegistry.withDefaults().freeze();
        this.errorHandleMode = ExcelErrorHandleMode.THROW;
        this.referenceParser = null;
    }

    private ExcelCalculateConfig(Field idField, ExcelIdTypeEnum excelIdTypeEnum, ExcelRowIdCodec rowIdCodec,
                                 Field cellTypesField, Field sheetField, Field sortField, List<Field> columnFieldList,
                                 ExcelCalculateTypeEnum calculateType, FunctionRegistry functionRegistry,
                                 ExcelErrorHandleMode errorHandleMode) {
        this.idField = idField;
        this.excelIdTypeEnum = excelIdTypeEnum;
        this.rowIdCodec = rowIdCodec;
        this.cellTypesField = cellTypesField;
        this.sheetField = sheetField;
        this.sortField = sortField;
        this.columnFieldList = columnFieldList;
        this.calculateType = calculateType;
        this.functionRegistry = functionRegistry;
        this.errorHandleMode = ExcelErrorHandleMode.orThrow(errorHandleMode);
        this.referenceParser = calculateType == null ? null
                : ExcelFormulaReferenceParser.of(calculateType, rowIdCodec);
    }

    private static EntityMeta scanEntity(Class<?> clazz) {
        Objects.requireNonNull(clazz, "实体类不能为空");
        Field id = null;
        ExcelIdTypeEnum idType = null;
        Class<? extends ExcelRowIdCodec> codecClass = ExcelRowIdCodec.None.class;
        Field cellTypes = null;
        Field sheet = null;
        Field sort = null;
        List<Field> columns = new ArrayList<>();
        for (Field field : clazz.getDeclaredFields()) {
            field.setAccessible(true);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (null != excelField) {
                switch (excelField.value()) {
                    case CELL -> columns.add(field);
                    case ID -> {
                        id = field;
                        idType = excelField.idType();
                        codecClass = excelField.idCodec();
                    }
                    case CELL_TYPES -> cellTypes = field;
                    case SORT -> sort = field;
                    case SHEET -> sheet = field;
                    default -> {
                    }
                }
            }
        }
        return new EntityMeta(id, idType, ExcelRowIdCodecs.resolve(codecClass, idType),
                cellTypes, sheet, sort, Collections.unmodifiableList(columns));
    }

    private static Field accessible(Field field) {
        field.setAccessible(true);
        return field;
    }

    public Optional<String> findReturnCellPrefix(String formula) {
        return requireParser().findReturnPrefix(formula);
    }

    public List<String> getCellStrListByFormula(String formula) {
        return requireParser().findCellRefs(formula);
    }

    public FunctionRegistry getFunctionRegistry() {
        return functionRegistry;
    }

    private ExcelFormulaReferenceParser requireParser() {
        requireCalculateType();
        return this.referenceParser;
    }

    private void requireCalculateType() {
        if (null == this.calculateType) {
            throw new ExcelException("请设置 Excel 计算类型");
        }
    }

    public Field getIdField() {
        return idField;
    }

    public ExcelIdTypeEnum getExcelIdTypeEnum() {
        return excelIdTypeEnum;
    }

    public ExcelRowIdCodec getRowIdCodec() {
        return rowIdCodec;
    }

    public Field getCellTypesField() {
        return cellTypesField;
    }

    public Field getSheetField() {
        return sheetField;
    }

    public Field getSortField() {
        return sortField;
    }

    public List<Field> getColumnFieldList() {
        return columnFieldList;
    }

    public ExcelCalculateTypeEnum getCalculateType() {
        return calculateType;
    }

    public ExcelErrorHandleMode getErrorHandleMode() {
        return errorHandleMode;
    }

    @Override
    public String toString() {
        return "ExcelCalculateConfig{" +
                "idField=" + idField +
                ", excelIdTypeEnum=" + excelIdTypeEnum +
                ", rowIdCodec=" + (rowIdCodec == null ? null : rowIdCodec.getClass().getSimpleName()) +
                ", calculateType=" + calculateType +
                ", errorHandleMode=" + errorHandleMode +
                '}';
    }

    /**
     * 配置构建器：内置函数之上按调用顺序叠加类或 Lambda，同名后覆盖先。
     */
    public static final class Builder {

        private final EntityMeta meta;
        private final FunctionRegistry registry = FunctionRegistry.withDefaults();
        private ExcelErrorHandleMode errorHandleMode = ExcelErrorHandleMode.THROW;
        private ExcelRowIdCodec rowIdCodec;

        private Builder(EntityMeta meta) {
            this.meta = meta;
            this.rowIdCodec = meta.rowIdCodec;
        }

        public Builder function(String name, ExcelFunctionHandler handler) {
            registry.register(name, handler);
            return this;
        }

        public Builder functions(IFunction... functionImpls) {
            if (functionImpls != null) {
                for (IFunction functionImpl : functionImpls) {
                    if (functionImpl != null) {
                        registry.register(functionImpl);
                    }
                }
            }
            return this;
        }

        @SafeVarargs
        public final Builder functions(Class<? extends IFunction>... functionImplClasses) {
            if (functionImplClasses != null) {
                for (Class<? extends IFunction> functionImplClass : functionImplClasses) {
                    if (functionImplClass != null) {
                        registry.register(functionImplClass);
                    }
                }
            }
            return this;
        }

        public Builder withErrorHandleMode(ExcelErrorHandleMode mode) {
            this.errorHandleMode = ExcelErrorHandleMode.orThrow(mode);
            return this;
        }

        public Builder withRowIdCodec(ExcelRowIdCodec codec) {
            this.rowIdCodec = codec;
            return this;
        }

        public ExcelCalculateConfig build() {
            registry.freeze();
            return new ExcelCalculateConfig(
                    meta.idField,
                    meta.excelIdTypeEnum,
                    rowIdCodec,
                    meta.cellTypesField,
                    meta.sheetField,
                    meta.sortField,
                    meta.columnFieldList,
                    null,
                    registry,
                    errorHandleMode
            );
        }
    }

    private static final class EntityMeta {
        private final Field idField;
        private final ExcelIdTypeEnum excelIdTypeEnum;
        private final ExcelRowIdCodec rowIdCodec;
        private final Field cellTypesField;
        private final Field sheetField;
        private final Field sortField;
        private final List<Field> columnFieldList;

        private EntityMeta(Field idField, ExcelIdTypeEnum excelIdTypeEnum, ExcelRowIdCodec rowIdCodec,
                           Field cellTypesField, Field sheetField, Field sortField, List<Field> columnFieldList) {
            this.idField = idField;
            this.excelIdTypeEnum = excelIdTypeEnum;
            this.rowIdCodec = rowIdCodec;
            this.cellTypesField = cellTypesField;
            this.sheetField = sheetField;
            this.sortField = sortField;
            this.columnFieldList = columnFieldList;
        }
    }
}
