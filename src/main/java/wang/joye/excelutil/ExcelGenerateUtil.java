package wang.joye.excelutil;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.RandomUtil;
import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.http.HttpUtil;
import lombok.Data;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * excel生成工具
 * @author joye
 * @since 2022/01/24
 */
@Data
public class ExcelGenerateUtil {

    /**
     * key:配置名, value: ExportConfig
     */
    private static Map<String, BeanToExcelConfig> configMap = new HashMap<>();
    /**
     * 默认列宽和默认行高
     */
    private static final short DEFAULT_COLUMN_WIDTH = 6000, DEFAULT_ROW_HEIGHT = 500;

    public static void initConfig(List<BeanToExcelConfig> configs) {
        for (BeanToExcelConfig config : configs) {
            configMap.put(config.getName(), config);
        }
    }

    /**
     * 将数据转换为excel列
     * @param list 数据列表
     * @param configName 配置项名
     */
    private static <T> List<ExcelColumn> list2Column(List<T> list, String configName) {
        return list2Column(list, configName, null);
    }

    /**
     * 将数据转换为excel列，可配置要转换的列
     * @param list 数据列表
     * @param configName 配置项名
     * @param fileNameList 要转换的字段，为Null代表导出全部字段
     */
    public static <T> List<ExcelColumn> list2Column(List<T> list, String configName, List<String> fileNameList) {
        if (list == null || list.size() == 0) {
            return null;
        }
        if (configMap == null || configMap.size() == 0) {
            throw new RuntimeException("请先初始化配置(initConfig)");
        }
        // 读取对应的配置
        BeanToExcelConfig config = configMap.get(configName);
        if (config == null) {
            throw new RuntimeException("未找到excel转换配置：" + configName);
        }

        List<ExcelColumn> realExportColumns = new LinkedList<>();
        // 如果要导出的属性为空，则默认全部导出
        if (fileNameList == null) {
            realExportColumns = config.getFields();
        } else {
            // 只保存要导出的属性
            for (ExcelColumn field : config.getFields()) {
                if (fileNameList.contains(field.getName())) {
                    realExportColumns.add(field);
                }
            }
        }

        for (T item : list) {
            // 对于每个item, 遍历要导出的属性
            for (ExcelColumn configField : realExportColumns) {
                // 有可能是复合属性，通过 . 分隔符来分割出数组
                String[] fields = configField.getName().split("\\.");
                // 比如属性名为 user.name
                // 先从原始item中查找user，再从user中查找name
                Object fieldValue = item;
                for (String field : fields) {
                    fieldValue = ReflectUtil.getFieldValue(fieldValue, field);
                }
                configField.addColumnData(fieldValue);
            }
        }
        return realExportColumns;
    }

    public static <T> HSSFWorkbook generateWorkbook(List<T> list, String configName) {
        return generateWorkbook(list, configName, null);
    }

    public static <T> HSSFWorkbook generateWorkbook(List<T> list, String configName, String excelTitle) {
        return generateWorkbookByColumn(list2Column(list, configName), excelTitle);
    }

    public static HSSFWorkbook generateWorkbookByColumn(List<ExcelColumn> data) {
        return generateWorkbookByColumn(data, null);
    }

    /**
     * 生成excel
     * 使用列的组织形式，是因为一行中每列的数据格式不一样。只有列上的数据是一样的。
     * 如果用行，则必须要用map保存一行中每个列的详细格式，所以使用列的组织形式更省内存
     * @param columnList excel的每一列格式及数据
     * @param excelTitle excel标题
     */
    public static HSSFWorkbook generateWorkbookByColumn(List<ExcelColumn> columnList, String excelTitle) {

        //防止NPE
        if (columnList == null || columnList.size() == 0) {
            throw new RuntimeException("记录为空");
            // return;
        }
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 表头行 单元格style 水平居中+垂直居中+加粗
        HSSFCellStyle tableHeaderStyle = workbook.createCellStyle();
        tableHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        tableHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        tableHeaderStyle.setFont(font);

        // 普通单元格style 水平居中+垂直居中+自动换行
        HSSFCellStyle commonStyle = workbook.createCellStyle();
        commonStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        commonStyle.setAlignment(HorizontalAlignment.CENTER);
        commonStyle.setWrapText(true);

        HSSFSheet sheet = workbook.createSheet();

        // 画图的顶级管理器，一个sheet只能获取一个
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();

        // 设置每列的宽度
        for (int i = 0; i < columnList.size(); i++) {
            sheet.setColumnWidth(i, DEFAULT_COLUMN_WIDTH);
        }

        // 当前行位置
        int rowOffset = 0;
        List<HSSFRow> rows = new LinkedList<>();

        // 因为所有列都是一样的，所以只需要检索第一列
        List<Object> firstColumn = columnList.get(0).getColumnDataList();
        for (int i = 0; i < firstColumn.size(); i++) {
            // 初始化所有行
            HSSFRow row = sheet.createRow(i);
            row.setHeight(DEFAULT_ROW_HEIGHT);
            rows.add(row);
        }
        // 额外添加一行，否则标题行会越界
        rows.add(sheet.createRow(rows.size()));
        // 额外添加一行，否则表头行会越界
        rows.add(sheet.createRow(rows.size()));

        // 检索出重复数据的所有位置
        List<Range> duplicateDataRanges = new LinkedList<>();
        for (int i = 0; i < firstColumn.size(); i++) {
            // 最后一条重复数据的索引位置
            int end = i;
            for (int j = i + 1; j < firstColumn.size(); j++) {
                if (firstColumn.get(i).equals(firstColumn.get(j))) {
                    end = j;
                } else {
                    break;
                }
            }
            if (i != end) {
                Range range = new Range();
                range.setStart(i);
                range.setEnd(end);
                duplicateDataRanges.add(range);
                i = end - 1;
            }
        }

        // 如果有标题，则设置标题为合并单元格样式
        if (excelTitle != null) {
            // 将标题写到第一个单元格
            // 获取标题行时，偏移量+1
            HSSFCell cell = rows.get(rowOffset++).createCell(0);
            cell.setCellStyle(commonStyle);
            cell.setCellValue(excelTitle);

            // 合并单元格
            // 四个参数：起始行，结束行，起始列，结束列
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, columnList.size() - 1);
            sheet.addMergedRegion(region);
        }

        // 设置表头样式和内容
        for (int i = 0; i < columnList.size(); i++) {
            ExcelColumn k = columnList.get(i);
            HSSFCell cell = rows.get(rowOffset).createCell(i);
            cell.setCellValue(k.getTitle());
            cell.setCellStyle(tableHeaderStyle);
        }
        // 处理完表头行时，偏移量+1
        rowOffset++;

        // j为列下标
        for (int j = 0; j < columnList.size(); j++) {
            ExcelColumn exportColumn = columnList.get(j);
            // i为行下标
            for (int i = 0; i < exportColumn.getColumnDataList().size(); i++) {
                // 获取行时，加上偏移量
                HSSFRow row = rows.get(rowOffset + i);
                HSSFCell cell = row.createCell(j);
                cell.setCellStyle(commonStyle);

                Object dataValue = exportColumn.getColumnDataList().get(i);
                if (dataValue == null) {
                    cell.setCellValue("");
                    continue;
                }

                switch (exportColumn.getType()) {
                    case ExcelColumn.TYPE_STRING:
                        cell.setCellValue(String.valueOf(dataValue));
                        break;
                    case ExcelColumn.TYPE_DATE:
                        LocalDate date = (LocalDate) dataValue;
                        cell.setCellValue(date.format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
                        break;
                    case ExcelColumn.TYPE_DATETIME:
                        LocalDateTime time = (LocalDateTime) dataValue;
                        cell.setCellValue(time.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
                        break;
                    case ExcelColumn.TYPE_MAP:
                        // 处理map类型，比如 0：进行中，1：已完成
                        Map<String, String> valueMap = exportColumn.getValueMap();
                        cell.setCellValue(valueMap.getOrDefault(dataValue.toString(), ""));
                        break;
                    case ExcelColumn.TYPE_IMAGE_URL:
                        // 网络图片
                        String url = (String) dataValue;
                        byte[] bytes = downloadImage(url);
                        if (bytes == null) {
                            break;
                        }
                        // dx2和dy2为1023和255，代表占满整个单元格
                        // col1,row1和col2,row2，代表绘图的左上角开始单元格与右下角结束单元格
                        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255, (short) j, (short) (i + rowOffset), (short) j, (short) (i + rowOffset));
                        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
                        patriarch.createPicture(anchor, workbook.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG));
                        break;
                    case ExcelColumn.TYPE_LOCAL_IMAGE:
                        // 本地图片
                        String path = (String) dataValue;
                        byte[] fileBytes = readLocalImage(path);
                        if (fileBytes == null) {
                            break;
                        }
                        HSSFClientAnchor anchor2 = new HSSFClientAnchor(0, 0, 1023, 255, (short) j, (short) (i + rowOffset), (short) j, (short) (i + rowOffset));
                        anchor2.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
                        patriarch.createPicture(anchor2, workbook.addPicture(fileBytes, HSSFWorkbook.PICTURE_TYPE_JPEG));
                        break;
                    case ExcelColumn.TYPE_BASE64_IMAGE:
                        byte[] base64Bytes = base64ToBytes(dataValue.toString());
                        HSSFClientAnchor anchor3 = new HSSFClientAnchor(0, 0, 1023, 255, (short) j, (short) (i - 1), (short) j, (short) (i - 1));
                        anchor3.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
                        patriarch.createPicture(anchor3, workbook.addPicture(base64Bytes, HSSFWorkbook.PICTURE_TYPE_JPEG));
                        break;
                    default:
                        throw new RuntimeException("未知的导出类型：" + exportColumn.getType());
                }
            }

            // 如果此列需要合并重复数据(合并单元格)
            if (exportColumn.getMergeRepeatRow() != null && exportColumn.getMergeRepeatRow()) {
                for (Range range : duplicateDataRanges) {
                    // 因为有表头行，所以合并单元格时需要往下偏移一行
                    int regionOffset = 1;
                    // 如果有标题行，需要再往下偏移一行
                    if (excelTitle != null) {
                        regionOffset++;
                    }
                    // 合并单元格
                    // 四个参数：起始行，结束行，起始列，结束列，合并时需要加上偏移行
                    CellRangeAddress region = new CellRangeAddress(range.getStart() + regionOffset, range.getEnd() + regionOffset, j, j);
                    sheet.addMergedRegion(region);
                }
            }
        }

        return workbook;
    }

    /**
     * 生成excel，并写入到file
     * @param list 数据列表
     * @param configName 使用的转换配置，list里的数据通过此配置转成map
     */
    public static <T> File generateFile(List<T> list, String configName) {
        return generateFile(list, configName, null, null);
    }

    /**
     * 生成excel，并写入到file
     * @param list 数据列表
     * @param configName 使用的转换配置，list里的数据通过此配置转成map
     */
    public static <T> File generateFile(List<T> list, String configName, String excelTitle) {
        return generateFile(list, configName, excelTitle, null);
    }

    /**
     * 生成excel，并写入到file
     * @param list 数据列表
     * @param configName 使用的转换配置，list里的数据通过此配置转成map
     * @param excelTitle excel标题行名称
     * @param fileName excel文件名称
     * @return file
     */
    public static <T> File generateFile(List<T> list, String configName, String excelTitle, String fileName) {
        List<ExcelColumn> maps = list2Column(list, configName);
        File file;
        // 如果未设置文件名，随机设置一个文件名
        if (StrUtil.isBlank(fileName)) {
            fileName = RandomUtil.randomString(10);
        }
        try {
            file = File.createTempFile(fileName, ".xlsx").getCanonicalFile();
        } catch (IOException e) {
            throw new RuntimeException("临时文件创建异常", e);
        }
        HSSFWorkbook workbook = generateWorkbookByColumn(maps, excelTitle);
        try {
            workbook.write(file);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException("HSSWorkbook导出到文件异常", e);
        }
        return file;
    }

    public static <T> void generateAndWrite(List<T> list, String configName, OutputStream out) {
        generateAndWrite(list, configName, null, null, out);
    }

    public static <T> void generateAndWrite(List<T> list, String configName, String fileName, OutputStream out) {
        generateAndWrite(list, configName, null, fileName, out);
    }

    public static <T> void generateAndWrite(List<T> list, String configName, String excelTitle, String fileName, OutputStream out) {
        HSSFWorkbook workbook = generateWorkbook(list, configName, excelTitle);
        try {
            workbook.write(out);
            workbook.close();
            out.flush();
            out.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 下载网络图片
     */
    private static byte[] downloadImage(String url) {
        byte[] bytes;
        try {
            bytes = HttpUtil.downloadBytes(url);
        } catch (Exception e) {
            throw new RuntimeException("下载图片失败: " + url, e);
        }
        return bytes;
    }

    /**
     * 读取本地文件
     */
    private static byte[] readLocalImage(String path) {
        byte[] bytes;
        try {
            bytes = FileUtil.readBytes(path);
        } catch (Exception e) {
            throw new RuntimeException("读取本地图片失败: " + path, e);
        }
        return bytes;
    }

    /**
     * base图片转换成byte数组
     */
    private static byte[] base64ToBytes(String base64) {
        // base64的特征前缀 如：data:image/jpeg;base64,/0Vysddf...
        String tag = "base64,";
        // 在前30个字符中查找 base64字样，图片的base64可能有前缀，有前缀要去掉
        int index = base64.substring(0, 30).indexOf(tag);
        // 如果未找到，代表
        if (index == -1) {
            index = 0;
        } else {
            index = index + tag.length();
        }
        return Base64.getDecoder().decode(base64.substring(index));
    }

    /**
     * 范围类，用于查找重复的行，然后合并单元格
     */
    @Data
    private static class Range {
        /**
         * start为重复数据的开始行位置
         */
        private Integer start;
        /**
         * end为重复数据的结束行位置
         */
        private Integer end;
    }

}