package org.fanlychie.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.fanlychie.excel.write.DataFormat;
import org.fanlychie.excel.write.RowStyle;
import org.fanlychie.excel.write.Sheet;
import org.fanlychie.excel.write.WritableExcel;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel 工具类
 * Created by fanlychie on 2017/3/5.
 */
public final class Excels {

    /**
     * 默认的工作表
     */
    private static final Sheet DEFAULT_SHEET = buildDefaultSheet();

    /**
     * 默认的布尔值字符串映射表
     */
    private static final Map<Boolean, String> DEFAULT_BOOLEAN_STRING_MAPPING = buildDefaultBooleanStringMapping();

    /**
     * 快速输出 Excel 文件, 若数据列表为空, 则抛出 {@link org.fanlychie.excel.exception.WriteExcelException} 异常
     *
     * @param data 数据列表
     * @return 返回可写的 Excel 对象
     */
    public static WritableExcel write(Collection<?> data) {
        return write(data, null);
    }

    /**
     * 快速输出 Excel 文件, 当数据列表为空时, 输出一个除了标题无实体内容的文件
     *
     * @param data     数据列表
     * @param dataType 数据类型
     * @return 返回可写的 Excel 对象
     */
    public static WritableExcel write(Collection<?> data, Class<?> dataType) {
        List<?> dataList = null;
        if (data instanceof List) {
            dataList = (List<?>) data;
        } else {
            dataList = new ArrayList<>(data);
        }
        return new WritableExcel(DEFAULT_SHEET).booleanStringMapping(DEFAULT_BOOLEAN_STRING_MAPPING).data(dataList, dataType);
    }

    /**
     * 获取默认的工作表
     *
     * @return 返回默认的工作表对象
     */
    public static Sheet getDefaultSheet() {
        return DEFAULT_SHEET;
    }

    /**
     * 构建默认的工作表对象
     */
    private static Sheet buildDefaultSheet() {
        Sheet sheet = new Sheet("Sheet1");
        sheet.setCellWidth(20);
        sheet.setBodyRowStyle(buildDefaultBodyRowStyle());
        sheet.setTitleRowStyle(buildDefaultTitleRowStyle());
        return sheet;
    }

    /**
     * 构建默认的标题行样式
     */
    private static RowStyle buildDefaultTitleRowStyle() {
        RowStyle titleRowStyle = new RowStyle();
        titleRowStyle.setIndex(0);
        titleRowStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleRowStyle.setBackgroundColor(IndexedColors.YELLOW.index);
        titleRowStyle.setBorder(CellStyle.BORDER_THIN, IndexedColors.GREY_25_PERCENT.index);
        titleRowStyle.setFont(12, IndexedColors.BLUE_GREY.index);
        titleRowStyle.setHeight(28);
        titleRowStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        titleRowStyle.setWrapText(false);
        titleRowStyle.setFormat(DataFormat.getDefault(String.class));
        return titleRowStyle;
    }

    /**
     * 构建默认的主体行样式
     */
    private static RowStyle buildDefaultBodyRowStyle() {
        RowStyle bodyRowStyle = new RowStyle();
        bodyRowStyle.setIndex(1);
        bodyRowStyle.setBackgroundColor(IndexedColors.LIGHT_TURQUOISE.index);
        bodyRowStyle.setBorder(CellStyle.BORDER_THIN, IndexedColors.GREY_25_PERCENT.index);
        bodyRowStyle.setFont(11, IndexedColors.GREY_50_PERCENT.index);
        bodyRowStyle.setHeight(24);
        bodyRowStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        bodyRowStyle.setWrapText(true);
        return bodyRowStyle;
    }

    /**
     * 构建默认的布尔值字符串映射表
     */
    private static Map<Boolean, String> buildDefaultBooleanStringMapping() {
        Map<Boolean, String> mapping = new HashMap<>();
        mapping.put(true, "是");
        mapping.put(false, "否");
        return mapping;
    }

    /**
     * 私有化构造
     */
    private Excels() {

    }

}