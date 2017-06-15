package org.fanlychie.jexcel.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.fanlychie.jexcel.annotation.AnnotationHandler;
import org.fanlychie.jexcel.annotation.CellField;
import org.fanlychie.jreflect.BeanDescriptor;
import org.fanlychie.jexcel.exception.ExcelCastException;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * 可写的 Excel, 用于输出 Excel 文件
 * Created by fanlychie on 2017/3/4.
 */
public class WritableExcel {

    /**
     * 工作表
     */
    private Sheet sheet;

    /**
     * 数据列表
     */
    private List<?> data;

    /**
     * 脚部数据
     */
    private List<?> footerData;

    /**
     * XSSF 工作表
     */
    private XSSFSheet xSSFSheet;

    /**
     * XSSF 工作薄
     */
    private XSSFWorkbook xSSFWorkbook;

    /**
     * 单元格注解字段列表
     */
    private List<CellField> cellFields;

    /**
     * 脚部单元格注解字段列表
     */
    private List<CellField> footerCellFields;

    /**
     * 布尔值字符串映射表
     */
    private Map<Boolean, String> booleanStringMapping;

    /**
     * 构建实例
     *
     * @param sheet 工作表
     */
    public WritableExcel(Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * 填充数据, 若数据列表为空, 则抛出 {@link java.lang.IllegalArgumentException} 异常
     *
     * @param data 数据列表
     * @return 返回当前对象
     */
    public WritableExcel data(List<?> data) {
        if (data == null || data.size() == 0) {
            throw new IllegalArgumentException("data can not be empty");
        }
        return data(data, data.get(0).getClass());
    }

    /**
     * 填充数据, 当数据列表为空时, 输出一个除了标题无实体内容的文件
     *
     * @param data     数据列表
     * @param dataType 数据类型
     * @return 返回当前对象
     */
    public WritableExcel data(List<?> data, Class<?> dataType) {
        this.data = data;
        if (dataType == null) {
            if (data == null && data.isEmpty()) {
                throw new IllegalArgumentException("data and type can not be empty at the same time");
            } else {
                dataType = data.get(0).getClass();
            }
        }
        this.cellFields = AnnotationHandler.parseClass(dataType);
        return this;
    }

    /**
     * 脚部填充数据, 当数据列表为空时, 不创建脚部行
     *
     * @param footerData 数据列表
     * @return 返回当前对象
     */
    public WritableExcel footer(List<?> footerData) {
        this.footerData = footerData;
        if (footerData != null && footerData.size() > 0) {
            this.footerCellFields = AnnotationHandler.parseClass(footerData.get(0).getClass());
        }
        return this;
    }

    /**
     * 设置布尔值字符串映射表
     *
     * @param booleanStringMapping 布尔值字符串映射表
     * @return 返回当前对象
     */
    public WritableExcel booleanStringMapping(Map<Boolean, String> booleanStringMapping) {
        this.booleanStringMapping = booleanStringMapping;
        return this;
    }

    /**
     * 输出到文件
     *
     * @param pathname 文件路径名称
     */
    public void toFile(String pathname) {
        toFile(new File(pathname));
    }

    /**
     * 输出到文件
     *
     * @param file 文件对象
     */
    public void toFile(File file) {
        try {
            toStream(new FileOutputStream(file));
        } catch (FileNotFoundException e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 写出到输出流
     *
     * @param out 输出流
     */
    public void toStream(OutputStream out) {
        try {
            this.xSSFWorkbook = new XSSFWorkbook();
            this.xSSFSheet = xSSFWorkbook.createSheet(sheet.getName());
            buildExcelTitleRow();
            if (data != null && !data.isEmpty()) {
                RowStyle bodyRowStyle = sheet.getBodyRowStyle();
                int bodyIndex = bodyRowStyle.getIndex();
                for (Object item : data) {
                    buildExcelBodyRow(bodyRowStyle, bodyIndex++, item);
                }
                if (footerData != null && !footerData.isEmpty()) {
                    RowStyle footerRowStyle = sheet.getFooterRowStyle();
                    int footerIndex = data.size() + 1;
                    for (Object item : footerData) {
                        buildExcelFooterRow(footerRowStyle, footerIndex++, item);
                    }
                }
            }
            xSSFWorkbook.write(out);
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 构建 Excel 标题行内容
     *
     * @throws Throwable
     */
    private void buildExcelTitleRow() throws Throwable {
        RowStyle rowStyle = sheet.getTitleRowStyle();
        XSSFRow row = xSSFSheet.createRow(rowStyle.getIndex());
        row.setHeightInPoints(rowStyle.getHeight());
        CellStyle cellStyle = rowStyle.getCellStyle(xSSFWorkbook);
        for (CellField cellField : cellFields) {
            int index = cellField.getIndex();
            xSSFSheet.setColumnWidth(index, sheet.getCellWidth());
            XSSFCell cell = row.createCell(index);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(cellField.getName());
        }
    }

    /**
     * 构建 Excel 主体行内容
     *
     * @param rowStyle 行样式
     * @param index    行索引
     * @param obj      填充单元格的对象数据
     * @throws Throwable
     */
    private void buildExcelBodyRow(RowStyle rowStyle, int index, Object obj) throws Throwable {
        XSSFRow row = xSSFSheet.createRow(index);
        row.setHeightInPoints(rowStyle.getHeight());
        BeanDescriptor beanDescriptor = new BeanDescriptor(obj);
        for (CellField cellField : cellFields) {
            XSSFCell cell = row.createCell(cellField.getIndex());
            CellStyle cellStyle = rowStyle.getCellStyle(xSSFWorkbook);
            cellStyle.setAlignment(cellField.getAlign().getValue());
            Object value = beanDescriptor.getValueByName(cellField.getField());
            Class<?> type = cellField.getType();
            String format = cellField.getFormat();
            setCellValue(cell, cellStyle, value, type, format);
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * 构建 Excel 脚部行内容
     *
     * @param rowStyle 行样式
     * @param index    行索引
     * @param obj      填充单元格的对象数据
     * @throws Throwable
     */
    private void buildExcelFooterRow(RowStyle rowStyle, int index, Object obj) throws Throwable {
        XSSFRow row = xSSFSheet.createRow(index);
        row.setHeightInPoints(rowStyle.getHeight());
        BeanDescriptor beanDescriptor = new BeanDescriptor(obj);
        CellStyle cellStyle = rowStyle.getCellStyle(xSSFWorkbook);
        for (CellField cellField : footerCellFields) {
            int cellIndex = cellField.getIndex();
            XSSFCell cell = row.createCell(cellIndex);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(cellField.getName());
            cell = row.createCell(cellIndex + 1);
            Object value = beanDescriptor.getValueByName(cellField.getField());
            Class<?> type = cellField.getType();
            String format = cellField.getFormat();
            setCellValue(cell, cellStyle, value, type, format);
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * 设置单元格的值
     *
     * @param cell      单元格对象
     * @param cellStyle 单元格样式
     * @param value     值
     * @param type      值的类型
     * @param format    数据格式
     */
    private void setCellValue(XSSFCell cell, CellStyle cellStyle, Object value, Class<?> type, String format) {
        if (value == null) {
            setCellStringValue(cell, cellStyle, "");
        } else if (type == Boolean.TYPE || type == Boolean.class) {
            setCellBooleanValue(cell, cellStyle, value, format);
        } else if ((Number.class.isAssignableFrom(type) || type.isPrimitive()) && type != Byte.TYPE && type != Character.TYPE) {
            setCellDataFormat(cellStyle, format);
            cell.setCellValue(Double.parseDouble(value.toString()));
        } else if (type == Date.class) {
            cell.setCellValue((Date) value);
            setCellDataFormat(cellStyle, format);
        } else {
            setCellStringValue(cell, cellStyle, value.toString());
        }
    }

    /**
     * 设置单元格字符串值
     *
     * @param cell      单元格对象
     * @param cellStyle 单元格样式
     * @param value     值
     */
    private void setCellStringValue(XSSFCell cell, CellStyle cellStyle, String value) {
        cell.setCellValue(value);
        setCellDataFormat(cellStyle, DataFormat.STRING_FORMAT);
    }

    /**
     * 设置单元格布尔类型的值
     *
     * @param cell      单元格对象
     * @param cellStyle 单元格样式
     * @param value     值
     * @param format    数据格式
     */
    private void setCellBooleanValue(XSSFCell cell, CellStyle cellStyle, Object value, String format) {
        boolean boolValue = (boolean) value;
        if (booleanStringMapping != null) {
            String booleanStr = booleanStringMapping.get(boolValue);
            if (booleanStr != null) {
                setCellStringValue(cell, cellStyle, booleanStr);
            } else {
                cell.setCellValue(boolValue);
                setCellDataFormat(cellStyle, format);
            }
        } else {
            cell.setCellValue(boolValue);
            setCellDataFormat(cellStyle, format);
        }
    }

    /**
     * 设置单元格数据格式
     *
     * @param cellStyle 单元格样式
     * @param format    数据格式
     */
    private void setCellDataFormat(CellStyle cellStyle, String format) {
        XSSFDataFormat formatter = xSSFWorkbook.createDataFormat();
        short dataFormat = formatter.getFormat(format);
        cellStyle.setDataFormat(dataFormat);
    }

}