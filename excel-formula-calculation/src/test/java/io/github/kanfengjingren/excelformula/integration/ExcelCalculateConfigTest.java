package io.github.kanfengjingren.excelformula.integration;

import io.github.kanfengjingren.excelformula.common.config.ExcelCalculateConfig;
import io.github.kanfengjingren.excelformula.module.row.RowCalculate;
import io.github.kanfengjingren.excelformula.support.ExcelNumber;
import io.github.kanfengjingren.excelformula.support.ExcelSampleData;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * {@link ExcelCalculateConfig} 在构造计算模块时拷贝快照、不污染共享配置实例。
 */
@DisplayName("集成：ExcelCalculateConfig 快照行为")
class ExcelCalculateConfigTest {

    @Test
    @DisplayName("forType 不修改共享配置实例")
    void forTypeDoesNotMutateSharedInstance() {
        ExcelCalculateConfig shared = ExcelSampleData.numberConfig();
        assertThat(shared.getCalculateType()).isNull();
        RowCalculate<ExcelNumber> row = new RowCalculate<>(ExcelSampleData.blankNumberRow(), shared);
        assertThat(shared.getCalculateType()).isNull();
        assertThat(row.getExcelCalculateConfig().getCalculateType().name()).isEqualTo("ROW");
    }
}
