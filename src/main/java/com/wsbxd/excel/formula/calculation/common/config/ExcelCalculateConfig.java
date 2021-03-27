package com.wsbxd.excel.formula.calculation.common.config;

import com.wsbxd.excel.formula.calculation.common.exception.ExcelException;
import com.wsbxd.excel.formula.calculation.common.field.annotation.ExcelField;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelFieldTypeEnum;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelIdTypeEnum;
import com.wsbxd.excel.formula.calculation.common.config.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.common.function.DefaultFunctionImpl;
import com.wsbxd.excel.formula.calculation.common.interfaces.IFunction;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * description: Excel 计算配置
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelCalculateConfig {

    /**
     * 不带页签数字行单元格匹配
     */
    private final static Pattern CELL_NUMBER_PATTERN = Pattern.compile("[A-Z]+\\d+");

    /**
     * 不带页签UUID行单元格匹配
     */
    private final static Pattern CELL_UUID_PATTERN = Pattern.compile("[A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}");

    /**
     * 带页签数字行单元格匹配
     */
    public final static Pattern SHEET_CELL_NUMBER_PATTERN = Pattern.compile("('[^\\\\/?*\\[\\]]+?'![A-Z]+\\d+|[^\\\\/?*\\[\\]():,+-]+?![A-Z]+\\d+|[A-Z]+\\d+)");

    /**
     * 带页签UUID行单元格匹配
     */
    public final static Pattern SHEET_CELL_UUID_PATTERN = Pattern.compile("('[^\\\\/?*\\[\\]]+?'![A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|[^\\\\/?*\\[\\]():,+-]+?![A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|[A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12})");

    /**
     * 单行列单元格匹配
     */
    private final static Pattern COLUMN_CELL_PATTERN = Pattern.compile("(?![A-Z]+\\()([A-Z]+)");

    /**
     * 返回不带页签数字行单元格匹配
     */
    private final static Pattern RETURN_CELL_NUMBER_PATTERN = Pattern.compile("^[A-Z]+\\d+=");

    /**
     * 返回不带页签UUID行单元格匹配
     */
    private final static Pattern RETURN_CELL_UUID_PATTERN = Pattern.compile("^[A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}=");

    /**
     * 返回带页签数字行单元格匹配
     */
    public final static Pattern RETURN_SHEET_CELL_NUMBER_PATTERN = Pattern.compile("^('[^\\\\/?*\\[\\]]+?'![A-Z]+\\d+|[^\\\\/?*\\[\\]():,+-]+?![A-Z]+\\d+|[A-Z]+\\d+)=");

    /**
     * 返回带页签UUID行单元格匹配
     */
    public final static Pattern RETURN_SHEET_CELL_UUID_PATTERN = Pattern.compile("^('[^\\\\/?*\\[\\]]+?'![A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|[^\\\\/?*\\[\\]():,+-]+?![A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|[A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12})=");


    /**
     * 返回单行列单元格匹配
     */
    public final static Pattern RETURN_COLUMN_PATTERN = Pattern.compile("^(?![A-Z]+\\()([A-Z]+)=");

    /**
     * 唯一标识字段
     */
    private Field idField;

    /**
     * 唯一标识类型枚举
     */
    private ExcelIdTypeEnum excelIdTypeEnum;

    /**
     * 单元格类型字符集合字段
     */
    private Field cellTypesField;

    /**
     * 页签字段
     */
    private Field sheetField;

    /**
     * 排序字段
     */
    private Field sortField;

    /**
     * 列字段集合
     */
    private List<Field> columnFieldList;

    /**
     * Excel 计算类型
     */
    private ExcelCalculateTypeEnum calculateType;

    /**
     * 函数实现
     */
    private IFunction functionImpl;

    /**
     * 函数和函数名称 Map
     */
    private Map<String, Method> nameFunctionMap;

    /**
     * 函数调用
     *
     * @param functionName 函数名称
     * @param valueList    函数参数
     * @return 计算结果
     */
    public String functionCalculate(String functionName, List<String> valueList) {
        Method method = nameFunctionMap.get(functionName);
        try {
            return (String) method.invoke(this.functionImpl, valueList);
        } catch (IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
        throw new ExcelException("函数调用失败");
    }

    /**
     * 处理方法实现类
     *
     * @param functionImplClass 方法实现类
     */
    private void handleFunctionImpl(Class<? extends IFunction> functionImplClass) {
        try {
            this.functionImpl = functionImplClass.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        this.nameFunctionMap = new HashMap<>();
        for (Method method : functionImplClass.getDeclaredMethods()) {
            this.nameFunctionMap.put(method.getName(), method);
        }
    }

    /**
     * 不用注解创建 Excel 实体类 数据属性
     *
     * @param clazz               excel
     * @param idFieldName         id字段名称
     * @param excelIdTypeEnum     id类型
     * @param sheetFieldName      工作表字段名称
     * @param sortFieldName       排序字段名称
     * @param columnFieldNameList 列字段名称集合
     */
    public ExcelCalculateConfig(Class<?> clazz, String idFieldName, ExcelIdTypeEnum excelIdTypeEnum, String sheetFieldName, String sortFieldName, List<String> columnFieldNameList) {
        this(clazz, idFieldName, excelIdTypeEnum, sheetFieldName, sortFieldName, columnFieldNameList, DefaultFunctionImpl.class);
    }

    /**
     * 不用注解创建 Excel 实体类 数据属性
     *
     * @param clazz               excel
     * @param idFieldName         id字段名称
     * @param excelIdTypeEnum     id类型
     * @param sheetFieldName      工作表字段名称
     * @param sortFieldName       排序字段名称
     * @param columnFieldNameList 列字段名称集合
     * @param functionImplClass   函数实现类
     */
    public ExcelCalculateConfig(Class<?> clazz, String idFieldName, ExcelIdTypeEnum excelIdTypeEnum, String sheetFieldName, String sortFieldName, List<String> columnFieldNameList, Class<? extends IFunction> functionImplClass) {
        try {
            this.idField = clazz.getDeclaredField(idFieldName);
            this.excelIdTypeEnum = excelIdTypeEnum;
            this.sortField = clazz.getDeclaredField(sortFieldName);
            if (ExcelStrUtil.isNotBlank(sheetFieldName)) {
                //如果不是工作簿计算是用不到工作表名称的
                this.sheetField = clazz.getDeclaredField(sheetFieldName);
            }
            this.columnFieldList = new ArrayList<>();
            for (String column : columnFieldNameList) {
                this.columnFieldList.add(clazz.getDeclaredField(column));
            }
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        }
        handleFunctionImpl(functionImplClass);
    }

    /**
     * 用注释创建 Excel 实体类 数据属性
     *
     * @param clazz 数据类
     */
    public ExcelCalculateConfig(Class<?> clazz) {
        this(clazz, DefaultFunctionImpl.class);
    }

    /**
     * 用注释创建 Excel 实体类 数据属性
     *
     * @param clazz             数据类
     * @param functionImplClass 函数实现类
     */
    public ExcelCalculateConfig(Class<?> clazz, Class<? extends IFunction> functionImplClass) {
        for (Field field : clazz.getDeclaredFields()) {
            field.setAccessible(true);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (null != excelField) {
                if (ExcelFieldTypeEnum.CELL.name().equals(excelField.value().name())) {
                    if (null == this.columnFieldList) {
                        this.columnFieldList = new ArrayList<>();
                    }
                    this.columnFieldList.add(field);
                } else if (ExcelFieldTypeEnum.ID.name().equals(excelField.value().name())) {
                    this.idField = field;
                    this.excelIdTypeEnum = excelField.idType();
                } else if (ExcelFieldTypeEnum.CELL_TYPES.name().equals(excelField.value().name())) {
                    this.cellTypesField = field;
                } else if (ExcelFieldTypeEnum.SORT.name().equals(excelField.value().name())) {
                    this.sortField = field;
                } else if (ExcelFieldTypeEnum.SHEET.name().equals(excelField.value().name())) {
                    this.sheetField = field;
                }
            }
        }
        handleFunctionImpl(functionImplClass);
    }

    /**
     * 获取返回单元格匹配
     *
     * @return 单元格匹配 Pattern
     */
    public Pattern getReturnCellPattern() {
        if (null == this.calculateType) {
            throw new ExcelException("请设置 Excel 计算类型");
        }
        if (ExcelCalculateTypeEnum.BOOK.equals(this.calculateType)) {
            if (ExcelIdTypeEnum.NUMBER.equals(excelIdTypeEnum)) {
                return RETURN_SHEET_CELL_NUMBER_PATTERN;
            } else if (ExcelIdTypeEnum.UUID.equals(this.excelIdTypeEnum)) {
                return RETURN_SHEET_CELL_UUID_PATTERN;
            }
        } else if (ExcelCalculateTypeEnum.SHEET.equals(calculateType)) {
            if (ExcelIdTypeEnum.NUMBER.equals(this.excelIdTypeEnum)) {
                return RETURN_CELL_NUMBER_PATTERN;
            } else if (ExcelIdTypeEnum.UUID.equals(this.excelIdTypeEnum)) {
                return RETURN_CELL_UUID_PATTERN;
            }
        } else if (ExcelCalculateTypeEnum.ROW.equals(this.calculateType)) {
            return RETURN_COLUMN_PATTERN;
        }
        return null;
    }

    /**
     * 根据公式获取单元格字符串集合
     *
     * @param formula 公式
     * @return 单元格字符串集合
     */
    public List<String> getCellStrListByFormula(String formula) {
        if (null == this.calculateType) {
            throw new ExcelException("请设置 Excel 计算类型");
        }
        List<String> cellStrList = new ArrayList<>();
        Matcher matcher = null;
        if (ExcelCalculateTypeEnum.BOOK.equals(this.calculateType)) {
            if (ExcelIdTypeEnum.NUMBER.equals(excelIdTypeEnum)) {
                matcher = SHEET_CELL_NUMBER_PATTERN.matcher(formula);
            } else if (ExcelIdTypeEnum.UUID.equals(excelIdTypeEnum)) {
                matcher = SHEET_CELL_UUID_PATTERN.matcher(formula);
            }
        } else if (ExcelCalculateTypeEnum.SHEET.equals(this.calculateType)) {
            if (ExcelIdTypeEnum.NUMBER.equals(excelIdTypeEnum)) {
                matcher = CELL_NUMBER_PATTERN.matcher(formula);
            } else if (ExcelIdTypeEnum.UUID.equals(excelIdTypeEnum)) {
                matcher = CELL_UUID_PATTERN.matcher(formula);
            }
        } else if (ExcelCalculateTypeEnum.ROW.equals(this.calculateType)) {
            matcher = COLUMN_CELL_PATTERN.matcher(formula);
        }
        while (null != matcher && matcher.find()) {
            cellStrList.add(matcher.group());
        }
        return cellStrList;
    }

    public Field getIdField() {
        return idField;
    }

    public void setIdField(Field idField) {
        this.idField = idField;
    }

    public ExcelIdTypeEnum getExcelIdTypeEnum() {
        return excelIdTypeEnum;
    }

    public void setExcelIdTypeEnum(ExcelIdTypeEnum excelIdTypeEnum) {
        this.excelIdTypeEnum = excelIdTypeEnum;
    }

    public Field getCellTypesField() {
        return cellTypesField;
    }

    public void setCellTypesField(Field cellTypesField) {
        this.cellTypesField = cellTypesField;
    }

    public Field getSheetField() {
        return sheetField;
    }

    public void setSheetField(Field sheetField) {
        this.sheetField = sheetField;
    }

    public Field getSortField() {
        return sortField;
    }

    public void setSortField(Field sortField) {
        this.sortField = sortField;
    }

    public List<Field> getColumnFieldList() {
        return columnFieldList;
    }

    public void setColumnFieldList(List<Field> columnFieldList) {
        this.columnFieldList = columnFieldList;
    }

    public ExcelCalculateTypeEnum getCalculateType() {
        return calculateType;
    }

    public void setCalculateType(ExcelCalculateTypeEnum calculateType) {
        this.calculateType = calculateType;
    }

    public Map<String, Method> getNameFunctionMap() {
        return nameFunctionMap;
    }

    public void setNameFunctionMap(Map<String, Method> nameFunctionMap) {
        this.nameFunctionMap = nameFunctionMap;
    }

    public IFunction getFunctionImpl() {
        return functionImpl;
    }

    public void setFunctionImpl(IFunction functionImpl) {
        this.functionImpl = functionImpl;
    }

    public ExcelCalculateConfig(Field idField, ExcelIdTypeEnum excelIdTypeEnum, Field cellTypesField, Field sheetField, Field sortField, List<Field> columnFieldList, ExcelCalculateTypeEnum calculateType, IFunction functionImpl) {
        this.idField = idField;
        this.excelIdTypeEnum = excelIdTypeEnum;
        this.cellTypesField = cellTypesField;
        this.sheetField = sheetField;
        this.sortField = sortField;
        this.columnFieldList = columnFieldList;
        this.calculateType = calculateType;
        this.functionImpl = functionImpl;
    }

    @Override
    public String toString() {
        return "ExcelDataProperties{" +
                "idField=" + idField +
                ", excelIdTypeEnum=" + excelIdTypeEnum +
                ", cellTypesField=" + cellTypesField +
                ", sheetField=" + sheetField +
                ", sortField=" + sortField +
                ", columnFieldList=" + columnFieldList +
                ", calculateType=" + calculateType +
                '}';
    }
}
