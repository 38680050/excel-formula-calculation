package io.github.kanfengjingren.excelformula.common.reference;

/**
 * 行主键（嵌入公式单元格地址中的 ID 段）编解码策略。
 * <p>
 * 内置 NUMBER / SNOWFLAKE / UUID 等；自定义格式实现本接口并在注解或 Config 中挂载。
 */
public interface ExcelRowIdCodec {

    /**
     * 从 {@code text[from]} 起尝试读出行 ID。
     *
     * @return 成功时返回结束下标（不含）；失败返回 -1
     */
    int tryRead(String text, int from);

    /**
     * 整段字符串是否为合法行 ID（默认：{@link #tryRead} 恰好消费整串）。
     */
    default boolean matches(String id) {
        return id != null && tryRead(id, 0) == id.length();
    }

    /**
     * 注解占位：表示未指定自定义 Codec，回退到 {@code idType}。
     */
    final class None implements ExcelRowIdCodec {
        /** 占位实现：任意位置均无法读出行 ID。 */
        @Override
        public int tryRead(String text, int from) {
            return -1;
        }
    }
}
