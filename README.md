# excel-formula-calculation
Java 1.8 仿 Excel 公式计算,目前有单行公式计算,单页签公式计算和多页签公式计算

# 开源代码许可

![MIT开源代码许可证](https://img.shields.io/badge/license-MIT-green)

# 使用说明

## 一、创建计算配置
通过 [ExcelCalculateConfig.java](src/main/java/com/wsbxd/excel/formula/calculation/common/config/ExcelCalculateConfig.java) 中`ExcelCalculateConfig(Class<?> clazz, Class<? extends IFunction> functionImplClass)`和`ExcelCalculateConfig(Class<?> clazz)`
构造器 new 出对象，clazz 参数是将要计算对象的类，而 functionImplClass 参数是实现函数的类，如果不传functionImplClass参数，则默认使用 [DefaultFunctionImpl](src/main/java/com/wsbxd/excel/formula/calculation/common/function/DefaultFunctionImpl.java) 来实现函数计算。

## 二、创建计算器并计算
目前有三种计算器，[工作簿计算器 BookCalculate](src/main/java/com/wsbxd/excel/formula/calculation/module/book/BookCalculate.java)、[工作表计算器 SheetCalculate](src/main/java/com/wsbxd/excel/formula/calculation/module/sheet/SheetCalculate.java)、[单行计算器 RowCalculate](src/main/java/com/wsbxd/excel/formula/calculation/module/row/RowCalculate.java)，
分别针对多页签、单页签、单行数据进行计算。

使用时新建构造器，excelList 参数为将要计算的数据的集合，excelCalculateConfig 参数为第一步创建出的计算配置；

再使用`calculate(String formula)`函数进行公式计算，参数为将要计算的公式，**公式的格式与Excel公式的格式一致**，
可连续使用此函数就算，第二次的公式计算将依赖第一次公式计算的结果，以此类推，如果需要立马获取结果，可使用此函数的返回值；

如果需要将结果收敛至原计算数据的集合的对象中，可使用`integrationResult()`函数，这样就可以从原计算数据中获取数据。

## 三、自定义公式
[IFunction](src/main/java/com/wsbxd/excel/formula/calculation/common/interfaces/IFunction.java) 是Excel函数定义的接口，[DefaultFunctionImpl](src/main/java/com/wsbxd/excel/formula/calculation/common/function/DefaultFunctionImpl.java) 是默认的Excel函数实现类，目前里面自定义的函数较少，如需自定义Excel函数，
也可以新建类继承 [IFunction](src/main/java/com/wsbxd/excel/formula/calculation/common/interfaces/IFunction.java)，实现需要的参数后在第一步创建 [ExcelCalculateConfig.java](src/main/java/com/wsbxd/excel/formula/calculation/common/config/ExcelCalculateConfig.java) 构造器时添加到第二个参数内。

# Excel公式计算实现

下面所有东西都可以在 src/test/java 文件夹中找到,并有[excel公式计算.xlsx](src/test/java/com/wsbxd/excel/formula/calculation/excel公式计算.xlsx)辅助验算

## 单行公式计算
可在 [ExcelFormulaCalculationTests.java](src/test/java/com/wsbxd/excel/formula/calculation/ExcelFormulaCalculationTests.java) 37、57行找到对应案例

## 单页签公式计算
可在 [ExcelFormulaCalculationTests.java](src/test/java/com/wsbxd/excel/formula/calculation/ExcelFormulaCalculationTests.java) 77、99行找到对应案例

## 多页签公式计算
可在 [ExcelFormulaCalculationTests.java](src/test/java/com/wsbxd/excel/formula/calculation/ExcelFormulaCalculationTests.java) 121、138行找到对应案例