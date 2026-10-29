# excel-formula-calculation

在 Java 中对 POJO 数据执行仿 Excel 公式计算，支持**单行 / 单页签 / 多页签**，语法与 Excel 一致。

[![Java](https://img.shields.io/badge/Java-17+-orange)](pom.xml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Release](https://img.shields.io/github/v/release/38680050/excel-formula-calculation)](https://github.com/38680050/excel-formula-calculation/releases)

## 模块

| Artifact | 说明 |
|----------|------|
| `excel-formula-calculation` | 核心库，无 Spring 依赖（仅需 `slf4j-api`） |
| `excel-formula-calculation-spring-boot-starter` | 可选 Spring Boot 3 自动配置 |

版本：**1.0.0**

## 特性

- **嵌入式库**：核心 jar 不绑 Spring Boot
- **安全求值**：白名单算术引擎，不依赖 JS `ScriptEngine`
- **三级计算**：`RowCalculate` / `SheetCalculate` / `BookCalculate`
- **可扩展函数**：内置常用逻辑 / 统计 / 数学 / 文本函数（见下），可自定义
- **结果写回**：`integrationResult()` 将计算结果收敛到原 POJO
- **可选 Starter**：`excel.formula.*` 配置 + 自动注册 `ExcelFormulaTemplate`

## 快速开始

### 1. 定义实体（`@ExcelField`）

```java
public class RowData {
    @ExcelField(value = ExcelFieldTypeEnum.ID, idType = ExcelIdTypeEnum.NUMBER)
    private String id;

    @ExcelField  // 列名即字段名：A、B、C...
    private String A;
    @ExcelField
    private String B;
    @ExcelField
    private String C;

    @ExcelField(ExcelFieldTypeEnum.SORT)
    private Integer sort;

    // getter / setter ...
}
```

页签场景再增加 `@ExcelField(ExcelFieldTypeEnum.SHEET)`；单元格类型（百分比、日期等）可用 `CELL_TYPES` 字段描述。

行主键类型（嵌入公式地址，如 `A1` / `A{uuid}`）通过 `idType` 选择；也可用 `idCodec` 挂自定义实现：

| `idType` | 说明 | 公式中 ID 示例 |
|----------|------|----------------|
| `NUMBER` | 数字行号 | `1`、`12` |
| `SNOWFLAKE` | 雪花（可落入 `long` 的数字） | `1234567890123456789` |
| `UUID` | 带横线 UUID（小写） | `1bbd420f-3ce3-11ea-a676-00163e048143` |
| `UUID_COMPACT` | 无横线 UUID（32 位十六进制） | `1bbd420f3ce311eaa67600163e048143` |
| `OBJECT_ID` | MongoDB ObjectId（24 位十六进制） | `507f1f77bcf86cd799439011` |
| `ULID` | ULID（26 位 Crockford Base32） | `01ARZ3NDEKTSV4RRFFQ69G5FAV` |

### 2. 配置 + 计算

```java
ExcelCalculateConfig config = ExcelCalculateConfig.create(RowData.class);

// 单页签
SheetCalculate<RowData> calc = new SheetCalculate<>(rows, config);
String value = calc.calculate("C1=A1+B1*2");
calc.integrationResult();   // 写回 rows 中的 POJO

// 单行
RowCalculate<RowData> rowCalc = new RowCalculate<>(row, config);
rowCalc.calculate("C=A+ABS(B)");

// 多页签（必须指定当前页签；单参数 calculate 会抛异常）
BookCalculate<RowData> bookCalc = new BookCalculate<>(rows, config);
bookCalc.calculate("现金", "A2=银行存款!B1+SUM(A1:A10)");
```

同一份 `config` 可被多种 Calculator 复用；内部通过 `forType(...)` 生成快照，不会改写原配置。

### 3. 公式约定

| 场景 | 示例 |
|------|------|
| 赋值写回 | `C1=A1+B1` |
| 区域引用 | `SUM(A1:B5)` |
| 跨页签 | `现金!A1`、`'交易性-股票'!B2` |
| 比较 | `IF(A1>=B1,1,0)` |

可连续调用 `calculate`，后一次依赖前一次写回结果；需要即时结果时用返回值。

## 自定义函数

在**内置函数之上叠加**自定义实现；同名时**后注册覆盖先注册**（可覆盖内置）。

推荐 Builder：

```java
// Lambda
ExcelCalculateConfig config = ExcelCalculateConfig.of(RowData.class)
    .function("POWER", args -> String.valueOf(
        Math.pow(Double.parseDouble(args.get(0)), Double.parseDouble(args.get(1)))))
    .build();

// 类：方法名 = 函数名，签名 public String XXX(List<String> args)
public class MyFunctions implements IFunction {
    public String POWER(List<String> args) {
        return "...";
    }
}

ExcelCalculateConfig config = ExcelCalculateConfig.of(RowData.class)
    .functions(new MyFunctions())
    .build();

// 混合：BatchB 同名 > ROUND lambda > BatchA > 内置
ExcelCalculateConfig config = ExcelCalculateConfig.of(RowData.class)
    .functions(new BatchA())
    .function("ROUND", args -> "...")
    .functions(new BatchB())
    .build();
```

默认实现见 [`DefaultFunctionImpl`](excel-formula-calculation/src/main/java/io/github/kanfengjingren/excelformula/common/function/DefaultFunctionImpl.java)。

内置函数一览：

| 类别 | 函数 |
|------|------|
| 逻辑 | `IF` `AND` `OR` `NOT` `IFERROR` `IFNA` `ISERROR` `ISBLANK` `TRUE` `FALSE` |
| 统计 | `SUM` `PRODUCT` `AVERAGE` `COUNT` `COUNTA` `MAX` `MIN` |
| 数学 | `ABS` `ROUND` `ROUNDUP` `ROUNDDOWN` `INT` `MOD` `POWER` `SQRT` `SIGN` |
| 文本 | `LEN` `LEFT` `RIGHT` `MID` `TRIM` `UPPER` `LOWER` `CONCAT`/`CONCATENATE` `VALUE` |

## 使用方式

当前**未发布到 Maven Central**，其它工程不能只写依赖坐标就自动下载。可选：

**本机安装后再依赖（常用）**

```bash
mvn clean install
```

安装到本地仓库 `~/.m2` 后：

```xml
<!-- 纯 Java -->
<dependency>
    <groupId>io.github.kanfengjingren</groupId>
    <artifactId>excel-formula-calculation</artifactId>
    <version>1.0.0</version>
</dependency>

<!-- 或 Spring Boot -->
<dependency>
    <groupId>io.github.kanfengjingren</groupId>
    <artifactId>excel-formula-calculation-spring-boot-starter</artifactId>
    <version>1.0.0</version>
</dependency>
```

**仅打 jar**

```bash
mvn clean package
```

核心产物：`excel-formula-calculation/target/excel-formula-calculation-1.0.0.jar`。

核心库运行时请自行提供 SLF4J 实现（如 `slf4j-simple` / Logback）；Boot 工程一般已有日志实现。

## Spring Boot（可选）

引入 starter（会传递依赖核心库）后，可用 `excel.formula.*` 配置，并注入自动注册的 `ExcelFormulaTemplate`：

```yaml
excel:
  formula:
    error-handle-mode: THROW   # 或 RETURN_CODE
```

注入 `ExcelFormulaTemplate`，调用时传入实体类型即可：

```java
@Service
public class FormulaService {
    private final ExcelFormulaTemplate template;

    public FormulaService(ExcelFormulaTemplate template) {
        this.template = template;
    }

    public void calc(List<RowData> rows) {
        SheetCalculate<RowData> calc = template.sheet(RowData.class, rows);
        calc.calculate("C1=A1+B1*2");
        calc.integrationResult();  // 写回 rows
    }
}
```

也可用 `template.row(...)` / `template.book(...)`。需要自定义函数时先 `ExcelCalculateConfig.of(...).function(...).build()`，再交给 `template.sheet(config, rows)`（仍会套上 yml 全局默认）。

## 构建与测试

```bash
mvn test
mvn package
```

要求 **JDK 17+**（Starter 面向 Spring Boot 3.3+）。若要发布到 Central，需另配发布插件与签名。

## 示例与验算

集成测试按计算器拆分，样例实体与数据在 `support` 包；可用 [`excel公式计算.xlsx`](excel-formula-calculation/src/test/java/io/github/kanfengjingren/excelformula/excel公式计算.xlsx) 对照：

| 范围 | 测试类 |
|------|--------|
| 单行 | [`RowCalculateTest`](excel-formula-calculation/src/test/java/io/github/kanfengjingren/excelformula/integration/RowCalculateTest.java) |
| 单页签 | [`SheetCalculateTest`](excel-formula-calculation/src/test/java/io/github/kanfengjingren/excelformula/integration/SheetCalculateTest.java) |
| 多页签 | [`BookCalculateTest`](excel-formula-calculation/src/test/java/io/github/kanfengjingren/excelformula/integration/BookCalculateTest.java) |
| 引用/Codec | `common/reference/*Test` |
| 错误模式 | [`ExcelErrorHandleModeTest`](excel-formula-calculation/src/test/java/io/github/kanfengjingren/excelformula/common/exception/ExcelErrorHandleModeTest.java) |
| Boot 自动配置 | [`ExcelFormulaAutoConfigurationTest`](excel-formula-calculation-spring-boot-starter/src/test/java/io/github/kanfengjingren/excelformula/boot/ExcelFormulaAutoConfigurationTest.java) |

## License

[MIT](LICENSE)
