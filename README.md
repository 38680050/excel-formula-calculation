# excel-formula-calculation-old
Java 仿 excel 公式计算,目前有单行公式计算,单页签公式计算和多页签公式计算
## 使用说明
>下面所有东西都可以在 src/test/java 文件夹中找到,并有[excel公式计算.xlsx](src/test/java/com/wsbxd/excel/formula/calculation/excel公式计算.xlsx)辅助验算
### 注解@ExcelField
```java
public @interface ExcelField {

    ExcelFieldTypeEnum value() default ExcelFieldTypeEnum.COLUMN;

    ExcelIdTypeEnum idType() default ExcelIdTypeEnum.NONE;

}
```
`value`为单元格类型([ExcelFieldTypeEnum.java](src/main/java/com/wsbxd/excel/formula/calculation/annotation/ExcelFieldTypeEnum.java)),默认为`COLUMN` ,对应为excel中的A,B,C,D列
![excel列](https://raw.githubusercontent.com/38680050/image/master/excel列.png)
除`COLUMN`外,还有`ID(行数据唯一标识)`,`SORT(行数据排序)`,`SHEET(行数据页签名称)`,`CELL_TYPES(行数据单元格类型)`,
当前`CELL_TYPES`数据格式为`列名-单元格类型id`,如果有多个,则使用`;`拼接,举例`A-1;F-2;N-5`,单元格类型id可以在[ExcelCellTypeEnum.java](src/main/java/com/wsbxd/excel/formula/calculation/constant/ExcelCellTypeEnum.java)查看<br/>
<br/>
`idType`为id类型([ExcelIdTypeEnum.java](src/main/java/com/wsbxd/excel/formula/calculation/annotation/ExcelIdTypeEnum.java)),只有`value`为`ExcelFieldTypeEnum.id`才可设置,`idType`目前分为两种,id为数字(`ExcelIdTypeEnum.NUMBER`)或UUID(`ExcelIdTypeEnum.UUID`)
### EXCEL函数实现
在 [FunctionImpl.java](src/main/java/com/wsbxd/excel/formula/calculation/function/FunctionImpl.java) 中实现excel函数,目前实现的函数有`IF,ABS,SUM,MAX,MIN,ROUND`,如有需要可自行实现
### 单行公式计算
![单行公式计算](https://raw.githubusercontent.com/38680050/image/master/单行公式计算.png)
>可在 [ExcelFormulaCalculationApplicationTests.java](src/test/java/com/wsbxd/excel/formula/calculation/ExcelFormulaCalculationApplicationTests.java) 找到对应案例

因为是单行公式计算,因此`id,sheet,sort`都可以不用填,公式计算后可在`excels`中查看结果,如果要即时获取计算结果,也可使用`calculateHelper.calculateChangeValue(String)`的返回值查看当前计算结果
### 单页签公式计算
![单页签公式计算](https://raw.githubusercontent.com/38680050/image/master/单页签公式计算.png)
>可在 [ExcelFormulaCalculationApplicationTests.java](src/test/java/com/wsbxd/excel/formula/calculation/ExcelFormulaCalculationApplicationTests.java) 找到对应案例

因为是单页签公式计算,因此`sheet`可以不用填
### 多页签公式计算
![多页签公式计算](https://raw.githubusercontent.com/38680050/image/master/多页签公式计算.png)
>可在 [ExcelFormulaCalculationApplicationTests.java](src/test/java/com/wsbxd/excel/formula/calculation/ExcelFormulaCalculationApplicationTests.java) 找到对应案例