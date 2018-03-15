package org.fanlychie.jexcel;

import org.fanlychie.jexcel.read.ReadableExcel;
import org.fanlychie.jexcel.write.WritableExcel;
import org.fanlychie.jexcel.write.model.CustomExcelSheet;
import org.fanlychie.jexcel.write.model.DefaultExcelSheet;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * Excel 执行器
 * Created by fanlychie on 2017/3/5.
 */
public final class ExcelHelper {

    /**
     * 默认的布尔值字符串映射表
     */
    private static final Map<Boolean, String> DEFAULT_BOOLEAN_STRING_MAPPING = buildDefaultBooleanStringMapping();

    /**
     * 获取一个可写的默认的EXCEL对象
     *
     * @param dataClassType 填充工作表的数据类型
     * @return 返回可写的 EXCEL对象
     */
    public static WritableExcel getDefaultWritableExcel(Class<?> dataClassType) {
        return new WritableExcel(new DefaultExcelSheet(dataClassType))
                .booleanStringMapping(DEFAULT_BOOLEAN_STRING_MAPPING);
    }

    /**
     * 获取一个可写的默认的EXCEL对象
     *
     * @param dataClassType 填充工作表的数据类型
     * @return 返回可写的 EXCEL对象
     */
    public static WritableExcel getCustomWritableExcel(Class<?> dataClassType) {
        return new WritableExcel(new CustomExcelSheet(dataClassType))
                .booleanStringMapping(DEFAULT_BOOLEAN_STRING_MAPPING);
    }

    /**
     * 获取一个可读的EXCEL对象
     *
     * @param excelFile Excel 文件
     * @return 返回可读的 Excel 对象
     */
    public static ReadableExcel getReadableExcel(File excelFile) {
        return new ReadableExcel(excelFile).setStartRow(2);
    }

    /**
     * 获取一个可读的 Excel 对象
     *
     * @param excelFilePath Excel 文件路径
     * @return 返回可读的 Excel 对象
     */
    public static ReadableExcel getReadableExcel(String excelFilePath) {
        return new ReadableExcel(excelFilePath).setStartRow(2);
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
    private ExcelHelper() {

    }

}