package org.fanlychie.excel.read;

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

    public <T> List<T> parse(Class<T> targetClass) {
        try {
            // 工作薄
            Workbook workbook = WorkbookFactory.create(inputStream);
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
        return null;
    }

}