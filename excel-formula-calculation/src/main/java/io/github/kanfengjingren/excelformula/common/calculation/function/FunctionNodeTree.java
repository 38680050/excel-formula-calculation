package io.github.kanfengjingren.excelformula.common.calculation.function;

import java.util.List;

/**
 * 函数节点树构建工具。
 * <p>
 * 按括号下标的包含关系，把内层函数挂到外层函数的子节点下；
 * 若与现有根节点互不包含，则作为新的根节点加入。
 */
final class FunctionNodeTree {

    private FunctionNodeTree() {
    }

    /**
     * 将节点插入函数树。
     *
     * @param roots 当前根节点列表（或某节点的子节点列表）
     * @param node  待插入的函数节点
     * @param <T>   实体类型
     */
    static <T> void insert(List<ExcelFunctionNode<T>> roots, ExcelFunctionNode<T> node) {
        if (roots.isEmpty()) {
            roots.add(node);
            return;
        }
        boolean nested = false;
        for (ExcelFunctionNode<T> existing : roots) {
            if (existing.contain(node)) {
                insert(existing.getExcelFunctionNodeList(), node);
                nested = true;
            }
        }
        if (!nested) {
            roots.add(node);
        }
    }
}
