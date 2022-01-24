package wang.joye.excelutil;

import lombok.Data;

import java.util.List;

/**
 * 实体类转换为excel的详细配置
 * @author joye
 * @since 2022/01/24
 */
@Data
public class BeanToExcelConfig {
    private String name;
    private List<ExcelColumn> fields;
}
