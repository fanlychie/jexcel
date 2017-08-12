package org.fanlychie.jexcel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.fanlychie.jexcel.read.ReadableExcel;
import org.fanlychie.jexcel.read.ReadableSheet;
import org.fanlychie.jexcel.spec.Format;
import org.fanlychie.jexcel.write.RowStyle;
import org.fanlychie.jexcel.write.WritableExcel;
import org.fanlychie.jexcel.write.WritableSheet;

import java.io.File;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Excel 执行器
 * Created by fanlychie on 2017/3/5.
 */
public final class ExcelExecutor {

    /**
     * 默认的只读工作表
     */
    private static final ReadableSheet DEFAULT_READ_ONLY_SHEET = buildDefaultReadOnlySheet();

    /**
     * 默认的布尔值字符串映射表
     */
    private static final Map<Boolean, String> DEFAULT_BOOLEAN_STRING_MAPPING = buildDefaultBooleanStringMapping();

    /**
     * 获取一个可写的 Excel 对象
     *
     * @param dataType 填充工作表的数据类型
     * @return 返回可写的 Excel 对象
     */
    public static WritableExcel getWritableExcel(Class<?> dataType) {
        return new WritableExcel(buildDefaultWritableSheet(dataType))
                .booleanStringMapping(DEFAULT_BOOLEAN_STRING_MAPPING);
    }

    /**
     * 快速读取 Excel 文件
     *
     * @param file Excel 文件
     * @return 返回可读的 Excel 对象
     */
    public static ReadableExcel read(File file) {
        return new ReadableExcel(DEFAULT_READ_ONLY_SHEET).load(file);
    }

    /**
     * 快速读取 Excel 文件
     *
     * @param pathname Excel 文件路径名称
     * @return 返回可读的 Excel 对象
     */
    public static ReadableExcel read(String pathname) {
        return new ReadableExcel(DEFAULT_READ_ONLY_SHEET).load(pathname);
    }

    /**
     * 快速读取 Excel 文件流
     *
     * @param inputStream Excel 输入流
     * @return 返回可读的 Excel 对象
     */
    public static ReadableExcel read(InputStream inputStream) {
        return new ReadableExcel(DEFAULT_READ_ONLY_SHEET).load(inputStream);
    }

    /**
     * 构建默认的工作表对象
     */
    private static WritableSheet buildDefaultWritableSheet(Class<?> dataType) {
        WritableSheet writableSheet = new WritableSheet(dataType);
        writableSheet.setCellWidth(20);
        writableSheet.setBodyRowStyle(buildDefaultBodyRowStyle());
        writableSheet.setTitleRowStyle(buildDefaultTitleRowStyle());
        return writableSheet;
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
        titleRowStyle.setFormat(Format.getDefault(String.class));
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
     * 构建只读的工作表
     */
    private static ReadableSheet buildDefaultReadOnlySheet() {
        ReadableSheet readableSheet = new ReadableSheet();
        readableSheet.setIndex(0);
        readableSheet.setFirstRowNum(2);
        return readableSheet;
    }

    /**
     * 私有化构造
     */
    private ExcelExecutor() {

    }

}