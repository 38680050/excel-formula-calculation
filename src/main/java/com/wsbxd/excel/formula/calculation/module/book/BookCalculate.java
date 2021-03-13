package com.wsbxd.excel.formula.calculation.module.book;

import com.wsbxd.excel.formula.calculation.common.prop.ExcelEntityProperties;
import com.wsbxd.excel.formula.calculation.common.prop.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.module.book.entity.ExcelBook;
import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelCalculate;
import com.wsbxd.excel.formula.calculation.module.book.formula.BookFormula;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * description: Excel 工作簿 计算器
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class BookCalculate<T> implements IExcelCalculate {

    private final static Logger logger = LoggerFactory.getLogger(BookCalculate.class);

    /**
     * excel book
     */
    private final ExcelBook<T> excelBook;

    /**
     * excel data properties
     */
    private final ExcelEntityProperties properties;

    @Override
    public String calculate(String formula) {
        return calculate(null, formula);
    }

    public String calculate(String currentSheet, String formula) {
        BookFormula<T> bookFormula = new BookFormula<>(currentSheet, formula, this.properties);
        String value = bookFormula.calculate(currentSheet, this.excelBook);
        if (null != bookFormula.getReturnCell()) {
            this.excelBook.updateExcelCellValue(bookFormula.getReturnCell());
        }
        return value;
    }

    @Override
    public void integrationResult() {
        this.excelBook.integrationResult();
    }

    public BookCalculate(List<T> excelList, Class<T> tClass) {
        this.properties = new ExcelEntityProperties(ExcelCalculateTypeEnum.BOOK, tClass);
        this.excelBook = new ExcelBook<>(excelList, this.properties);
    }

}
