package org.fanlychie.jexcel.read;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.fanlychie.jexcel.annotation.AnnotationHandler;
import org.fanlychie.jexcel.annotation.CellField;
import org.fanlychie.jreflect.BeanDescriptor;
import org.fanlychie.jexcel.exception.ExcelCastException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 可读的 Excel, 用于读取 Excel 文件
 * Created by fanlychie on 2017/3/5.
 */
public class ReadableExcel {

    /**
     * 文件输入流
     */
    private InputStream inputStream;

    /**
     * 只读的工作表
     */
    private ReadOnlySheet readOnlySheet;

    /**
     * 构建实例
     *
     * @param readOnlySheet 只读的工作表
     */
    public ReadableExcel(ReadOnlySheet readOnlySheet) {
        this.readOnlySheet = readOnlySheet;
    }

    /**
     * 载入 Excel 文件
     *
     * @param file Excel 文件
     * @return 返回当前对象
     */
    public ReadableExcel load(File file) {
        try {
            this.inputStream = new FileInputStream(file);
            return this;
        } catch (FileNotFoundException e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 载入 Excel 文件
     *
     * @param pathname Excel 文件路径名称
     * @return 返回当前对象
     */
    public ReadableExcel load(String pathname) {
        return load(new File(pathname));
    }

    /**
     * 载入 Excel 文件流
     *
     * @param inputStream Excel 文件输入流
     * @return 返回当前对象
     */
    public ReadableExcel load(InputStream inputStream) {
        this.inputStream = inputStream;
        return this;
    }

    /**
     * 解析 Excel 文档内容
     *
     * @param targetClass 目标类, 每一行的内容解析为此类的一个实例
     * @param <T>         目标类型
     * @return 返回解析的数据列表
     */
    public <T> List<T> parse(Class<T> targetClass) {
        try {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = null;
            if (readOnlySheet.getName() != null) {
                sheet = workbook.getSheet(readOnlySheet.getName());
            } else {
                sheet = workbook.getSheetAt(readOnlySheet.getIndex());
            }
            int firstRowNum = readOnlySheet.getFirstRowNum();
            if (firstRowNum > 0) {
                firstRowNum--;
            }
            int lastRowNum;
            if (readOnlySheet.getLastRowNum() > firstRowNum) {
                lastRowNum = readOnlySheet.getLastRowNum() - 1;
            } else {
                lastRowNum = sheet.getLastRowNum();
            }
            List<T> list = new ArrayList<>();
            List<CellField> cellFields = AnnotationHandler.parseClass(targetClass);
            for (int i = firstRowNum; i < lastRowNum; i++) {
                list.add(convertRowToObject(sheet.getRow(i), targetClass, cellFields));
            }
            return list;
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 将行内容转换为对象表示
     *
     * @param row         行对象
     * @param targetClass 目标类
     * @param cellFields  单元格注解字段列表
     * @param <T>         目标类
     * @return 返回内容转换为的对象
     */
    private <T> T convertRowToObject(Row row, Class<T> targetClass, List<CellField> cellFields) {
        BeanDescriptor beanDescriptor = new BeanDescriptor(targetClass);
        T obj = beanDescriptor.newInstance();
        int size = cellFields.size();
        for (int i = 0; i < size; i++) {
            Cell cell = row.getCell(i);
            CellField cellField = cellFields.get(i);
            String field = cellField.getField();
            Class<?> type = cellField.getType();
            Object value = getCellValue(cell, type);
            beanDescriptor.setValueByName(field, value);
        }
        return obj;
    }

    /**
     * 获取单元格的值
     *
     * @param cell 单元格对象
     * @param type 数据类型
     * @return 返回单元格数据类型的值
     */
    private Object getCellValue(Cell cell, Class<?> type) {
        if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            return null;
        }
        if (type == Date.class) {
            return cell.getDateCellValue();
        }
        cell.setCellType(Cell.CELL_TYPE_STRING);
        String cellStringValue = cell.getStringCellValue();
        return ValueConverter.convertObjectValue(cellStringValue, type);
    }

}