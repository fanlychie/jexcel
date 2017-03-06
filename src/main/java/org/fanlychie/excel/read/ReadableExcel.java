package org.fanlychie.excel.read;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.fanlychie.reflection.exception.ExcelCastException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

/**
 * 可读的 Excel, 用于读取 Excel 文件
 * Created by fanlychie on 2017/3/5.
 */
public class ReadableExcel {

    /**
     * 工作表
     */
    private Sheet sheet;

    /**
     * 工作表索引值
     */
    private int sheetIndex;

    /**
     * 工作表名称
     */
    private String sheetName;

    /**
     * Excel 文件输入流
     */
    private InputStream inputStream;

    /**
     * 构建实例
     *
     * @param pathname 文件路径名称
     */
    public ReadableExcel(String pathname) {
        this(new File(pathname));
    }

    /**
     * 构建实例
     *
     * @param file 文件对象
     */
    public ReadableExcel(File file) {
        try {
            this.inputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 构建对象
     *
     * @param inputStream 文件输入流
     */
    public ReadableExcel(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    /**
     * 设置读取的工作表索引
     * @param sheetIndex 读取的工作表索引值, 从0开始
     * @return 返回当前对象
     */
    public ReadableExcel sheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    /**
     * 设置读取的工作表名称
     * @param sheetName 读取的工作表名称
     * @return 返回当前对象
     */
    public ReadableExcel sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public <T> List<T> parse(Class<T> targetClass) {
        try {
            // 工作薄
            Workbook workbook = WorkbookFactory.create(inputStream);
            // 获取要解析的工作表
            if (sheetName != null && !sheetName.isEmpty()) {
                sheet = workbook.getSheet(sheetName);
            } else {
                sheet = workbook.getSheetAt(sheetIndex);
            }
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
        return null;
    }

}