package io.github.kanfengjingren.excelformula.common.interfaces;

/**
 * 函数实现类标记接口。
 * <p>
 * 实现类中符合约定的公有方法会被扫描注册：方法名即 Excel 函数名，
 * 签名为 {@code public String 函数名(List<String> args)}。
 * 通过 {@code FunctionRegistry#register} 或配置 Builder 的 {@code functions} 叠加到内置函数之上。
 */
public interface IFunction {
}
