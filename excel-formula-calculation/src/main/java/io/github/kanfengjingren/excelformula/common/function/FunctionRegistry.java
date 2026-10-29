package io.github.kanfengjingren.excelformula.common.function;

import io.github.kanfengjingren.excelformula.common.config.enums.ExcelErrorHandleMode;
import io.github.kanfengjingren.excelformula.common.constant.ExcelConstant;
import io.github.kanfengjingren.excelformula.common.exception.ExcelErrorUtil;
import io.github.kanfengjingren.excelformula.common.exception.ExcelException;
import io.github.kanfengjingren.excelformula.common.interfaces.IFunction;
import io.github.kanfengjingren.excelformula.common.util.ExcelStrUtil;

import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Excel 函数注册表。
 * <p>
 * 先装入内置函数，再按注册顺序叠加类或 Lambda；同名后注册覆盖先注册。
 * 构建完成后应调用 {@link #freeze()}，之后不可再注册。
 */
public final class FunctionRegistry {

    /**
     * 允许参数中携带错误字面量的函数（否则注册表会提前按错误短路）。
     */
    private static final Set<String> ACCEPTS_ERROR_ARGS = Set.of(
            "IFERROR", "IFNA", "ISERROR"
    );

    private final Map<String, ExcelFunctionHandler> handlers = new LinkedHashMap<>();
    private boolean frozen;

    private FunctionRegistry() {
    }

    /**
     * 创建已包含 {@link DefaultFunctionImpl} 的注册表，可继续叠加自定义函数。
     */
    public static FunctionRegistry withDefaults() {
        FunctionRegistry registry = new FunctionRegistry();
        registry.register(new DefaultFunctionImpl());
        return registry;
    }

    /**
     * 注册 Lambda；同名覆盖已有实现。
     */
    public FunctionRegistry register(String functionName, ExcelFunctionHandler handler) {
        ensureMutable();
        if (ExcelStrUtil.isBlank(functionName)) {
            throw new ExcelException("函数名不能为空");
        }
        Objects.requireNonNull(handler, "函数处理器不能为空");
        handlers.put(functionName, handler);
        return this;
    }

    /**
     * 扫描实例中符合约定的公有方法并注册；方法名即函数名，同名覆盖。
     */
    public FunctionRegistry register(IFunction functionImpl) {
        ensureMutable();
        Objects.requireNonNull(functionImpl, "函数实现不能为空");
        scanAndRegister(functionImpl);
        return this;
    }

    /**
     * 无参构造实例后扫描注册；同名覆盖。
     */
    public FunctionRegistry register(Class<? extends IFunction> functionImplClass) {
        ensureMutable();
        Objects.requireNonNull(functionImplClass, "函数实现类不能为空");
        try {
            IFunction instance = functionImplClass.getDeclaredConstructor().newInstance();
            scanAndRegister(instance);
        } catch (ReflectiveOperationException e) {
            throw new ExcelException("无法实例化函数实现类: " + functionImplClass.getName(), e);
        }
        return this;
    }

    /**
     * 冻结注册表，之后不可再改。
     */
    public FunctionRegistry freeze() {
        this.frozen = true;
        return this;
    }

    /**
     * 按函数名调用（默认 THROW）。
     */
    public String invoke(String functionName, List<String> valueList) {
        return invoke(functionName, valueList, ExcelErrorHandleMode.THROW);
    }

    /**
     * 按函数名调用已注册处理器。
     */
    public String invoke(String functionName, List<String> valueList, ExcelErrorHandleMode mode) {
        ExcelFunctionHandler handler = handlers.get(functionName);
        if (handler == null) {
            return ExcelErrorUtil.fail(mode, ExcelConstant.EXCEL_ERROR_NAME, "未知函数: " + functionName);
        }
        if (!ACCEPTS_ERROR_ARGS.contains(functionName)) {
            for (String value : valueList) {
                if (ExcelErrorUtil.isError(value)) {
                    return ExcelErrorUtil.fail(mode, value, "函数参数错误: " + value);
                }
            }
        }
        try {
            return handler.apply(valueList);
        } catch (ExcelException e) {
            return ExcelErrorUtil.handleThrown(mode, e);
        } catch (NumberFormatException | ArithmeticException e) {
            return ExcelErrorUtil.fail(mode, ExcelConstant.EXCEL_ERROR_VALUE,
                    "函数调用失败，函数名称：" + functionName + "，原因：" + e.getMessage(), e);
        } catch (RuntimeException e) {
            throw new ExcelException("函数调用失败，函数名称：" + functionName + "，原因：" + e.getMessage(), e);
        }
    }

    /** 是否已注册指定函数名。 */
    public boolean contains(String functionName) {
        return handlers.containsKey(functionName);
    }

    /** 已注册函数名到处理器的只读视图。 */
    public Map<String, ExcelFunctionHandler> getHandlers() {
        return Collections.unmodifiableMap(handlers);
    }

    private void scanAndRegister(IFunction functionImpl) {
        for (Method method : functionImpl.getClass().getDeclaredMethods()) {
            if (Modifier.isPublic(method.getModifiers())
                    && method.getParameterCount() == 1
                    && List.class.isAssignableFrom(method.getParameterTypes()[0])
                    && String.class.equals(method.getReturnType())) {
                method.setAccessible(true);
                String name = method.getName();
                handlers.put(name, args -> invokeReflect(functionImpl, method, name, args));
            }
        }
    }

    private static String invokeReflect(IFunction target, Method method, String functionName, List<String> args) {
        try {
            return (String) method.invoke(target, args);
        } catch (ReflectiveOperationException e) {
            Throwable cause = e.getCause() != null ? e.getCause() : e;
            if (cause instanceof ExcelException excelException) {
                throw excelException;
            }
            if (cause instanceof NumberFormatException numberFormatException) {
                throw numberFormatException;
            }
            if (cause instanceof ArithmeticException arithmeticException) {
                throw arithmeticException;
            }
            if (cause instanceof RuntimeException runtimeException) {
                throw runtimeException;
            }
            throw new ExcelException("函数调用失败，函数名称：" + functionName + "，原因：" + cause.getMessage(), cause);
        }
    }

    private void ensureMutable() {
        if (frozen) {
            throw new ExcelException("函数注册表已冻结，不可再注册");
        }
    }

}
