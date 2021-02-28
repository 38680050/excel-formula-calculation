package com.wsbxd.excel.formula.calculation;

import com.wsbxd.excel.formula.calculation.common.field.annotation.ExcelField;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelFieldTypeEnum;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelIdTypeEnum;

/**
 * description: excel实体类（Number）
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2018/10/18 15:40
 */
public class ExcelNumber {

    @ExcelField(value = ExcelFieldTypeEnum.ID, idType = ExcelIdTypeEnum.NUMBER)
    private String id;
    @ExcelField
    private String A;
    @ExcelField
    private String B;
    @ExcelField
    private String C;
    @ExcelField
    private String D;
    @ExcelField
    private String E;
    @ExcelField
    private String F;
    @ExcelField
    private String G;
    @ExcelField
    private String H;
    @ExcelField
    private String I;
    @ExcelField
    private String J;
    @ExcelField
    private String K;
    @ExcelField
    private String L;
    @ExcelField
    private String M;
    @ExcelField
    private String N;
    @ExcelField(ExcelFieldTypeEnum.CELL_TYPES)
    private String cellTypes;
    @ExcelField(ExcelFieldTypeEnum.SHEET)
    private String sheet;
    @ExcelField(ExcelFieldTypeEnum.SORT)
    private Integer sort;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getA() {
        return A;
    }

    public void setA(String a) {
        A = a;
    }

    public String getB() {
        return B;
    }

    public void setB(String b) {
        B = b;
    }

    public String getC() {
        return C;
    }

    public void setC(String c) {
        C = c;
    }

    public String getD() {
        return D;
    }

    public void setD(String d) {
        D = d;
    }

    public String getE() {
        return E;
    }

    public void setE(String e) {
        E = e;
    }

    public String getF() {
        return F;
    }

    public void setF(String f) {
        F = f;
    }

    public String getG() {
        return G;
    }

    public void setG(String g) {
        G = g;
    }

    public String getH() {
        return H;
    }

    public void setH(String h) {
        H = h;
    }

    public String getI() {
        return I;
    }

    public void setI(String i) {
        I = i;
    }

    public String getJ() {
        return J;
    }

    public void setJ(String j) {
        J = j;
    }

    public String getK() {
        return K;
    }

    public void setK(String k) {
        K = k;
    }

    public String getL() {
        return L;
    }

    public void setL(String l) {
        L = l;
    }

    public String getM() {
        return M;
    }

    public void setM(String m) {
        M = m;
    }

    public String getN() {
        return N;
    }

    public void setN(String n) {
        N = n;
    }

    public String getCellTypes() {
        return cellTypes;
    }

    public void setCellTypes(String cellTypes) {
        this.cellTypes = cellTypes;
    }

    public String getSheet() {
        return sheet;
    }

    public void setSheet(String sheet) {
        this.sheet = sheet;
    }

    public Integer getSort() {
        return sort;
    }

    public void setSort(Integer sort) {
        this.sort = sort;
    }

    public ExcelNumber(String id, String a, String b, String c, String d, String e, String f, String g, String h, String i, String j, String k, String l, String m, String n, String cellTypes, String sheet, Integer sort) {
        this.id = id;
        A = a;
        B = b;
        C = c;
        D = d;
        E = e;
        F = f;
        G = g;
        H = h;
        I = i;
        J = j;
        K = k;
        L = l;
        M = m;
        N = n;
        this.cellTypes = cellTypes;
        this.sheet = sheet;
        this.sort = sort;
    }

    public ExcelNumber() {
    }

    @Override
    public String toString() {
        return "Excel{" +
                "id='" + id + '\'' +
                ", A='" + A + '\'' +
                ", B='" + B + '\'' +
                ", C='" + C + '\'' +
                ", D='" + D + '\'' +
                ", E='" + E + '\'' +
                ", F='" + F + '\'' +
                ", G='" + G + '\'' +
                ", H='" + H + '\'' +
                ", I='" + I + '\'' +
                ", J='" + J + '\'' +
                ", K='" + K + '\'' +
                ", L='" + L + '\'' +
                ", M='" + M + '\'' +
                ", N='" + N + '\'' +
                ", cellTypes='" + cellTypes + '\'' +
                ", sheet='" + sheet + '\'' +
                ", sort=" + sort +
                '}';
    }
}
