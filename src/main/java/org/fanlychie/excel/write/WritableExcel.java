package org.fanlychie.excel.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.fanlychie.excel.exception.WriteExcelException;
import org.fanlychie.reflection.BeanDescriptor;
import org.fanlychie.reflection.exception.ExcelCastException;

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
     * XSSF 工作表
     */
    private XSSFSheet xSSFSheet;

    /**
     * XSSF 工作薄
     */
    private XSSFWorkbook xSSFWorkbook;

    /**
     * Excel 字段域列表
     */
    private List<ExcelFieldDomain> excelFieldDomains;

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
     * 填充数据, 若数据列表为空, 则抛出 {@link org.fanlychie.excel.exception.WriteExcelException} 异常
     *
     * @param data 数据列表
     * @return 返回当前对象
     */
    public WritableExcel data(List<?> data) {
        // 数据为空时取不到目标 Class, 无法做解析
        if (data == null || data.size() == 0) {
            throw new WriteExcelException("输出 Excel 文件的数据不能为空");
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
                throw new IllegalArgumentException("输出 Excel 文件的数据和数据类型不能同时为空");
            } else {
                dataType = data.get(0).getClass();
            }
        }
        this.excelFieldDomains = AnnotationHandler.parseClass(dataType);
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
            // XSSF 工作薄
            this.xSSFWorkbook = new XSSFWorkbook();
            // XSSF 工作表
            this.xSSFSheet = xSSFWorkbook.createSheet(sheet.getName());
            // 构建标题行内容
            buildExcelTitleRow();
            // 只处理非空数据
            if (data != null && !data.isEmpty()) {
                // 主题行样式
                RowStyle bodyRowStyle = sheet.getBodyRowStyle();
                // 主体行索引
                int bodyIndex = bodyRowStyle.getIndex();
                // 迭代数据列表
                for (Object item : data) {
                    // 构建主体行内容
                    buildExcelBodyRow(bodyRowStyle, bodyIndex++, item);
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
        // 标题行样式
        RowStyle rowStyle = sheet.getTitleRowStyle();
        // 创建一行
        XSSFRow row = xSSFSheet.createRow(rowStyle.getIndex());
        // 设置行高
        row.setHeightInPoints(rowStyle.getHeight());
        // 单元格样式
        CellStyle cellStyle = rowStyle.getCellStyle(xSSFWorkbook);
        // 迭代数据域列表
        for (ExcelFieldDomain excelFieldDomain : excelFieldDomains) {
            // 单元格索引
            int index = excelFieldDomain.getIndex();
            // 设置宽度
            xSSFSheet.setColumnWidth(index, sheet.getCellWidth());
            // 创建单元格
            XSSFCell cell = row.createCell(index);
            // 设置单元格样式
            cell.setCellStyle(cellStyle);
            // 设置单元格的值
            cell.setCellValue(excelFieldDomain.getName());
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
        // 创建一行
        XSSFRow row = xSSFSheet.createRow(index);
        // 设置行高
        row.setHeightInPoints(rowStyle.getHeight());
        // Bean 描述符
        BeanDescriptor beanDescriptor = new BeanDescriptor(obj);
        // 迭代数据域列表
        for (ExcelFieldDomain excelFieldDomain : excelFieldDomains) {
            // 创建单元格
            XSSFCell cell = row.createCell(excelFieldDomain.getIndex());
            // 单元格样式
            CellStyle cellStyle = rowStyle.getCellStyle(xSSFWorkbook);
            // 对齐方式
            cellStyle.setAlignment(excelFieldDomain.getAlign().getValue());
            // 单元格值
            Object value = beanDescriptor.getValueByName(excelFieldDomain.getField());
            // 单元格数据类型
            Class<?> type = excelFieldDomain.getType();
            // 单元格格式
            String format = excelFieldDomain.getFormat();
            // 设置单元格的值
            setCellValue(cell, cellStyle, value, type, format);
            // 设置单元格样式
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
        // 空值以空白字符串填充
        if (value == null) {
            setCellStringValue(cell, cellStyle, "");
        }
        // 布尔类型, 若有做布尔->字符串内容映射, 则将布尔转成字符串表示
        else if (type == Boolean.TYPE || type == Boolean.class) {
            setCellBooleanValue(cell, cellStyle, value, format);
        }
        // 数值类型处理
        else if ((Number.class.isAssignableFrom(type) || type.isPrimitive()) && type != Byte.TYPE && type != Character.TYPE) {
            setCellDataFormat(cellStyle, format);
            cell.setCellValue(Double.parseDouble(value.toString()));
        }
        // 日期类型处理
        else if (type == Date.class) {
            cell.setCellValue((Date) value);
            setCellDataFormat(cellStyle, format);
        }
        // 其余类型视为字符串处理
        else {
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