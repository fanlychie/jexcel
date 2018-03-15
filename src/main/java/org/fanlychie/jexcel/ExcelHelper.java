package org.fanlychie.jexcel;

import org.fanlychie.jexcel.read.ReadableExcel;
import org.fanlychie.jexcel.write.WritableExcel;
import org.fanlychie.jexcel.write.model.ConfigurableExcelSheet;
import org.fanlychie.jexcel.write.model.CustomExcelSheet;
import org.fanlychie.jexcel.write.model.DefaultExcelSheet;

import java.io.File;
import java.io.InputStream;

/**
 * EXCEL帮助类
 *
 * Created by fanlychie on 2017/3/5.
 */
public final class ExcelHelper {

    /**
     * 获取一个可写的默认的EXCEL对象
     *
     * @return 返回可写的EXCEL对象
     */
    public static WritableExcel getDefaultWritableExcel() {
        return new WritableExcel(new DefaultExcelSheet());
    }

    /**
     * 获取一个可写的自定义的EXCEL对象
     *
     * @return 返回可写的EXCEL对象
     */
    public static WritableExcel getCustomWritableExcel() {
        return new WritableExcel(new CustomExcelSheet());
    }

    /**
     * 获取一个可写的自定义配置的EXCEL对象
     *
     * @param configFile 配置文件
     * @return 返回可写的EXCEL对象
     */
    public static WritableExcel getConfigurableWritableExcel(String configFile) {
        return new WritableExcel(new ConfigurableExcelSheet(configFile));
    }

    /**
     * 获取一个可写的自定义配置的EXCEL对象
     *
     * @param configFileInputStream 配置文件
     * @return 返回可写的EXCEL对象
     */
    public static WritableExcel getConfigurableWritableExcel(InputStream configFileInputStream) {
        return new WritableExcel(new ConfigurableExcelSheet(configFileInputStream));
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
     * 私有化构造
     */
    private ExcelHelper() {

    }

}