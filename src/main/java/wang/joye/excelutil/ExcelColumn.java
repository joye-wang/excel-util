package wang.joye.excelutil;

import lombok.Data;

import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * @author joye
 * @since 2022/01/24
 */
@Data
public class ExcelColumn {
    public static final String TYPE_STRING = "string";
    /**
     * 网络图片url
     */
    public static final String TYPE_IMAGE_URL = "imageUrl";
    /**
     * BASE64图片
     */
    public static final String TYPE_BASE64_IMAGE = "base64Image";
    /**
     * 本地图片
     */
    public static final String TYPE_LOCAL_IMAGE = "localImage";
    public static final String TYPE_DATE = "date";
    public static final String TYPE_DATETIME = "datetime";
    public static final String TYPE_MAP = "map";

    private String name;
    private String type;
    private String title;
    /**
     * 当连续的行数据重复时，是否合并
     */
    private Boolean mergeRepeatRow;
    private Map<String, String> valueMap;

    private List<Object> columnDataList;

    public void addColumnData(Object columnData) {
        if (columnDataList == null) {
            columnDataList = new LinkedList<>();
        }
        this.columnDataList.add(columnData);
    }
}
