package io.github.kanfengjingren.excelformula.support;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.common.reference.UuidRowIdCodec;

import java.util.ArrayList;
import java.util.List;

/**
 * 集成测试样例数据与常用配置。
 */
public final class ExcelSampleData {

    /** 第 1 行样例数据的 UUID 行主键。 */
    public static final String UUID_1 = "1bbd420f-3ce3-11ea-a676-00163e048143";
    /** 第 2 行样例数据的 UUID 行主键。 */
    public static final String UUID_2 = "2bbd420f-3ce3-11ea-a676-00163e048143";
    /** 第 3 行样例数据的 UUID 行主键。 */
    public static final String UUID_3 = "3bbd420f-3ce3-11ea-a676-00163e048143";
    /** 第 4 行样例数据的 UUID 行主键。 */
    public static final String UUID_4 = "4bbd420f-3ce3-11ea-a676-00163e048143";
    /** 第 5 行样例数据的 UUID 行主键。 */
    public static final String UUID_5 = "5bbd420f-3ce3-11ea-a676-00163e048143";

    private ExcelSampleData() {
    }

    /** 以 {@link ExcelNumber} 为实体、数字行主键的默认计算配置。 */
    public static ExcelCalculateConfig numberConfig() {
        return ExcelCalculateConfig.create(ExcelNumber.class);
    }

    /** 数字实体字段 + UUID 行 ID Codec 的计算配置。 */
    public static ExcelCalculateConfig uuidConfig() {
        return numberConfig().withRowIdCodec(UuidRowIdCodec.INSTANCE);
    }

    /** 仅保留行号占位、其余列为空的单行实体，便于测公式与错误路径。 */
    public static ExcelNumber blankNumberRow() {
        return row(null, "1", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
    }

    /** 单行完整样例，覆盖百分比、日期、单元格类型标记等字段。 */
    public static ExcelNumber singleRow() {
        return row(null, "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, null);
    }

    /** 多页签工作簿集成测试用五行数据，行主键为数字字符串。 */
    public static List<ExcelNumber> bookNumbers() {
        return bookRows("1", "2", "3", "4", "5");
    }

    /** 多页签工作簿集成测试用五行数据，行主键为 {@link #UUID_1}～{@link #UUID_5}。 */
    public static List<ExcelNumber> bookUuids() {
        return bookRows(UUID_1, UUID_2, UUID_3, UUID_4, UUID_5);
    }

    /** 单页签集成测试用五行数据，行主键为数字字符串。 */
    public static List<ExcelNumber> sheetNumbers() {
        return sheetRows("1", "2", "3", "4", "5");
    }

    /** 单页签集成测试用五行数据，行主键为 {@link #UUID_1}～{@link #UUID_5}。 */
    public static List<ExcelNumber> sheetUuids() {
        return sheetRows(UUID_1, UUID_2, UUID_3, UUID_4, UUID_5);
    }

    /** 组装带页签名与排序号的五行工作簿样例。 */
    private static List<ExcelNumber> bookRows(String id1, String id2, String id3, String id4, String id5) {
        List<ExcelNumber> list = new ArrayList<>();
        list.add(row(id1, "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", "现金", 1));
        list.add(row(id2, "4545", null, "74182", null, "0.0111", null, null, "61", null, "10086", null, "1010%", null, null, "L-3", "现金", 2));
        list.add(row(id3, null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", "交易性-股票", 1));
        list.add(row(id4, null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", "银行存款", 1));
        list.add(row(id5, "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, "1900-01-01", null, "777", null, "A-5;B-3;F-5;I-5;K-4", "银行存款", 2));
        return list;
    }

    /** 组装无页签名、仅含排序号的五行页签样例。 */
    private static List<ExcelNumber> sheetRows(String id1, String id2, String id3, String id4, String id5) {
        List<ExcelNumber> list = new ArrayList<>();
        list.add(row(id1, "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, 1));
        list.add(row(id2, "4545", null, "74182", null, "0.01", null, null, "6341", null, null, null, "1010%", null, null, "L-3", null, 2));
        list.add(row(id3, null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", null, 3));
        list.add(row(id4, null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", null, 4));
        list.add(row(id5, "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, null, null, "777", null, "A-5;B-3;F-5;I-5;K-4;N-3", null, 5));
        return list;
    }

    /** 按列值构造一行 {@link ExcelNumber}。 */
    private static ExcelNumber row(String id, String a, String b, String c, String d, String e, String f, String g, String h,
                                  String i, String j, String k, String l, String m, String n, String cellTypes,
                                  String sheet, Integer sort) {
        return new ExcelNumber(id, a, b, c, d, e, f, g, h, i, j, k, l, m, n, cellTypes, sheet, sort);
    }
}
