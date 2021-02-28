package com.wsbxd.excel.formula.calculation;

import com.wsbxd.excel.formula.calculation.module.book.BookCalculate;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;

/**
 * description: Excel 计算器 测试类
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
@SpringBootTest
class ExcelFormulaCalculationApplicationTests {

    /**
     * 多页签计算(数字id)
     */
    @Test
    public void bookNum() {
        List<ExcelNumber> excels = new ArrayList<ExcelNumber>() {{
            this.add(new ExcelNumber("1", "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", "现金", 1));
            this.add(new ExcelNumber("2", "4545", null, "74182", null, "0.0111", null, null, "61", null, "10086", null, "1010%", null, null, "L-3", "现金", 2));
            this.add(new ExcelNumber("3", null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", "交易性-股票", 1));
            this.add(new ExcelNumber("4", null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", "银行存款", 1));
            this.add(new ExcelNumber("5", "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, "1900-01-01", null, "777", null, "A-5;B-3;F-5;I-5;K-4", "银行存款", 2));
        }};
        long start = System.currentTimeMillis();
        BookCalculate<ExcelNumber> bookCalculate = new BookCalculate<>(excels, ExcelNumber.class);
        bookCalculate.calculateChangeValue("交易性-股票", "N3=现金!A1*银行存款!C1+ABS(MIN(O2+2,银行存款!B3:银行存款!E5)*2-MAX('交易性-股票'!A1:'交易性-股票'!D3))+MAX(现金!A2:现金!O3)");
        bookCalculate.integrationResult();
        System.out.println("-------------------------------------------------");
        System.out.println(System.currentTimeMillis() - start + "ms");
    }

    /**
     * 多页签计算(uuid)
     */
    @Test
    public void bookUUID() {
        List<ExcelUUID> excels = new ArrayList<ExcelUUID>() {{
            this.add(new ExcelUUID("1bbd420f-3ce3-11ea-a676-00163e048143", "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", "现金", 1));
            this.add(new ExcelUUID("2bbd420f-3ce3-11ea-a676-00163e048143", "4545", null, "74182", null, "0.0111", null, null, "61", null, "10086", null, "1010%", null, null, "L-3", "现金", 2));
            this.add(new ExcelUUID("3bbd420f-3ce3-11ea-a676-00163e048143", null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", "交易性-股票", 1));
            this.add(new ExcelUUID("4bbd420f-3ce3-11ea-a676-00163e048143", null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", "银行存款", 1));
            this.add(new ExcelUUID("5bbd420f-3ce3-11ea-a676-00163e048143", "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, "1900-01-01", null, "777", null, "A-5;B-3;F-5;I-5;K-4", "银行存款", 2));
        }};
        long start = System.currentTimeMillis();
        BookCalculate<ExcelUUID> bookCalculate = new BookCalculate<>(excels, ExcelUUID.class);
        bookCalculate.calculateChangeValue("交易性-股票", "N3bbd420f-3ce3-11ea-a676-00163e048143=现金!A1bbd420f-3ce3-11ea-a676-00163e048143*银行存款!C1bbd420f-3ce3-11ea-a676-00163e048143+ABS(MIN(O2bbd420f-3ce3-11ea-a676-00163e048143+2,银行存款!B3bbd420f-3ce3-11ea-a676-00163e048143:银行存款!E5bbd420f-3ce3-11ea-a676-00163e048143)*2-MAX('交易性-股票'!A1bbd420f-3ce3-11ea-a676-00163e048143:'交易性-股票'!D3bbd420f-3ce3-11ea-a676-00163e048143))+MAX(现金!A2bbd420f-3ce3-11ea-a676-00163e048143:现金!O3bbd420f-3ce3-11ea-a676-00163e048143)");
        bookCalculate.integrationResult();
        System.out.println("-------------------------------------------------");
        System.out.println(System.currentTimeMillis() - start + "ms");
    }

//    /**
//     * 单页签计算(数字id)
//     */
//    @Test
//    public void sheetNum() {
//        List<ExcelNumber> excels = new ArrayList<ExcelNumber>() {{
//            this.add(new ExcelNumber("1", "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, 1));
//            this.add(new ExcelNumber("2", "4545", null, "74182", null, "0.01", null, null, "6341", null, null, null, "1010%", null, null, "L-3", null, 2));
//            this.add(new ExcelNumber("3", null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", null, 3));
//            this.add(new ExcelNumber("4", null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", null, 4));
//            this.add(new ExcelNumber("5", "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, null, null, "777", null, "A-5;B-3;F-5;I-5;K-4;N-3", null, 5));
//        }};
//        long start = System.currentTimeMillis();
//        SheetCalculate<ExcelNumber> sheetCalculate = new SheetCalculate<>(excels, ExcelNumber.class);
//        sheetCalculate.calculateChangeValue("K5=A1+ABS(MIN(C5+2,A1:I5)*2-MAX(L1:N5))+MAX(A2:J2)");
//        sheetCalculate.calculateChangeValue("N5=ROUND(2*K5+5,0)");
//        sheetCalculate.integrationResult();
//        System.out.println("-------------------------------------------------");
//        System.out.println(System.currentTimeMillis() - start + "ms");
//    }
//
//    /**
//     * 单页签计算(uuid)
//     */
//    @Test
//    public void sheetUUID() {
//        List<ExcelUUID> excels = new ArrayList<ExcelUUID>() {{
//            this.add(new ExcelUUID("1bbd420f-3ce3-11ea-a676-00163e048143", "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, 1));
//            this.add(new ExcelUUID("2bbd420f-3ce3-11ea-a676-00163e048143", "4545", null, "74182", null, "0.01", null, null, "6341", null, null, null, "1010%", null, null, "L-3", null, 2));
//            this.add(new ExcelUUID("3bbd420f-3ce3-11ea-a676-00163e048143", null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", null, 3));
//            this.add(new ExcelUUID("4bbd420f-3ce3-11ea-a676-00163e048143", null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", null, 4));
//            this.add(new ExcelUUID("5bbd420f-3ce3-11ea-a676-00163e048143", "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, null, null, "777", null, "A-5;B-3;F-5;I-5;K-4;N-3", null, 5));
//        }};
//        long start = System.currentTimeMillis();
//        SheetCalculate<ExcelUUID> sheetCalculate = new SheetCalculate<>(excels, ExcelUUID.class);
//        sheetCalculate.calculateChangeValue("K5bbd420f-3ce3-11ea-a676-00163e048143=A1bbd420f-3ce3-11ea-a676-00163e048143+ABS(MIN(C5bbd420f-3ce3-11ea-a676-00163e048143+2,A1bbd420f-3ce3-11ea-a676-00163e048143:I5bbd420f-3ce3-11ea-a676-00163e048143)*2-MAX(L1bbd420f-3ce3-11ea-a676-00163e048143:N5bbd420f-3ce3-11ea-a676-00163e048143))+MAX(A2bbd420f-3ce3-11ea-a676-00163e048143:J2bbd420f-3ce3-11ea-a676-00163e048143)");
//        sheetCalculate.calculateChangeValue("N5bbd420f-3ce3-11ea-a676-00163e048143=ROUND(2*K5bbd420f-3ce3-11ea-a676-00163e048143+5,0)");
//        sheetCalculate.integrationResult();
//        System.out.println("-------------------------------------------------");
//        System.out.println(System.currentTimeMillis() - start + "ms");
//    }
//
//    /**
//     * 单行计算(数字id)
//     */
//    @Test
//    public void rowNum() {
//        ExcelNumber excel = new ExcelNumber(null, "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, null);
//        long start = System.currentTimeMillis();
//        RowCalculate<ExcelNumber> calculateHelper = new RowCalculate<>(excel);
//        calculateHelper.calculateChangeValue("I=A+ABS(MIN(O+2,B)*2-MAX(A:H))+MAX(K:N)");
//        calculateHelper.calculateChangeValue("J=ROUND(2*I+5,0)");
//        calculateHelper.integrationResult();
//        System.out.println("-------------------------------------------------");
//        System.out.println(System.currentTimeMillis() - start + "ms");
//    }
//
//    /**
//     * 单行计算(uuid)
//     */
//    @Test
//    public void rowUUID() {
//        ExcelUUID excel = new ExcelUUID(null, "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, null);
//        long start = System.currentTimeMillis();
//        RowCalculate<ExcelUUID> calculateHelper = new RowCalculate<>(excel);
//        calculateHelper.calculateChangeValue("I=A+ABS(MIN(O+2,B)*2-MAX(A:H))+MAX(K:N)");
//        calculateHelper.calculateChangeValue("J=ROUND(2*I+5,0)");
//        calculateHelper.integrationResult();
//        System.out.println("-------------------------------------------------");
//        System.out.println(System.currentTimeMillis() - start + "ms");
//    }

}
