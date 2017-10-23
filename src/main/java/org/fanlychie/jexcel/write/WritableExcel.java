package org.fanlychie.jexcel.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.fanlychie.jexcel.annotation.AnnotationHandler;
import org.fanlychie.jexcel.annotation.CellField;
import org.fanlychie.jexcel.exception.ExcelCastException;
import org.fanlychie.jexcel.spec.Sheet;
import org.fanlychie.jreflect.BeanDescriptor;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
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
    private WritableSheet writableSheet;

    /**
     * SXSSF 工作表
     */
    private SXSSFSheet sxssfSheet;

    /**
     * SXSSF 工作薄
     */
    private SXSSFWorkbook sxssfWorkbook;

    /**
     * 单元格注解字段列表
     */
    private List<CellField> cellFields;

    /**
     * 布尔值字符串映射表
     */
    private Map<Boolean, String> booleanStringMapping;

    /**
     * 工作表计数
     */
    private int sheetCount = 1;

    /**
     * 下一行索引值
     */
    private int nextRowIndex;

    /**
     * 脚部的单元格样式
     */
    private CellStyle footerCellStyle;

    /**
     * 构建实例
     *
     * @param writableSheet 工作表
     */
    public WritableExcel(WritableSheet writableSheet) {
        this.writableSheet = writableSheet;
        this.sxssfWorkbook = new SXSSFWorkbook();
        if (writableSheet.getDataType() == null) {
            throw new IllegalArgumentException("dataType can not be null");
        }
        this.cellFields = AnnotationHandler.parseClass(writableSheet.getDataType());
    }

    /**
     * 填充数据, 当数据列表为空时, 输出一个除了标题无实体内容的文件
     *
     * @param data 数据列表
     * @return 返回当前对象
     */
    public WritableExcel addSheet(List<?> data) {
        return addSheet(Sheet.getName(sheetCount++, writableSheet.getName()), data);
    }

    /**
     * 填充数据, 当数据列表为空时, 输出一个除了标题无实体内容的文件
     *
     * @param sheetName 工作表名称
     * @param data      数据列表
     * @return 返回当前对象
     */
    public WritableExcel addSheet(String sheetName, List<?> data) {
        try {
            nextRowIndex = data.size() + 1;
            sxssfSheet = sxssfWorkbook.createSheet(sheetName);
            buildExcelTitleRow();
            if (data != null && !data.isEmpty()) {
                RowStyle bodyRowStyle = writableSheet.getBodyRowStyle();
                buildExcelBodyColumnStyle(bodyRowStyle);
                int bodyIndex = bodyRowStyle.getIndex();
                for (Object item : data) {
                    buildExcelBodyRow(bodyRowStyle, bodyIndex++, item);
                }
            }
            return this;
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 添加一行
     *
     * @param startIndex 起始索引值, 从0开始, 表示第一列
     * @param values     文本值
     * @return 返回当前对象
     */
    public WritableExcel addRow(int startIndex, Object... values) {
        if (nextRowIndex == 0) {
            throw new IllegalStateException();
        }
        if (footerCellStyle == null) {
            RowStyle rowStyle = writableSheet.getFooterRowStyle();
            footerCellStyle = rowStyle.buildCellStyle(sxssfWorkbook);
        }
        int startColumnIndex = startIndex;
        SXSSFRow row = sxssfSheet.createRow(nextRowIndex++);
        RowStyle rowStyle = writableSheet.getFooterRowStyle();
        row.setHeightInPoints(rowStyle.getHeight());
        SXSSFCell cell;
        for (int i = 0; i < values.length; i++) {
            cell = row.createCell(startIndex++);
            cell.setCellStyle(footerCellStyle);
            setCellValue(cell, values[i], String.class);
        }
        // 参数startIndex前面的单元格样式补全
        if (startColumnIndex > 0) {
            for (int i = 0; i < startColumnIndex; i++) {
                cell = row.createCell(i++);
                cell.setCellStyle(footerCellStyle);
            }
        }
        // 参数startIndex后面的单元格样式补全
        int interval = cellFields.size() - (startColumnIndex + values.length);
        if (interval > 0) {
            for (int i = 0; i < interval; i++) {
                cell = row.createCell(startIndex++);
                cell.setCellStyle(footerCellStyle);
            }
        }
        return this;
    }

    /**
     * 输出到文件
     *
     * @param pathname 文件路径名称
     */
    public void toFile(String pathname) {
        try {
            sxssfWorkbook.write(new FileOutputStream(new File(pathname)));
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 输出到文件
     *
     * @param file 文件对象
     */
    public void toFile(File file) {
        OutputStream os = null;
        try {
            sxssfWorkbook.write(os);
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {}
            }
        }
    }

    /**
     * 输出到输出流
     *
     * @param os 输出流
     */
    public void toStream(OutputStream os) {
        try {
            sxssfWorkbook.write(os);
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 获取可写的工作表对象
     *
     * @return 返回可写的工作表对象
     */
    public WritableSheet getWritableSheet() {
        return writableSheet;
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
     * 构建 Excel 标题行内容
     *
     * @throws Throwable
     */
    private void buildExcelTitleRow() throws Throwable {
        RowStyle rowStyle = writableSheet.getTitleRowStyle();
        SXSSFRow row = sxssfSheet.createRow(rowStyle.getIndex());
        row.setHeightInPoints(rowStyle.getHeight());
        for (CellField cellField : cellFields) {
            int index = cellField.getIndex();
            sxssfSheet.setColumnWidth(index, writableSheet.getCellWidth());
            SXSSFCell cell = row.createCell(index);
            cell.setCellStyle(rowStyle.buildCellStyle(sxssfWorkbook));
            cell.setCellValue(cellField.getName());
        }
    }

    /**
     * 构建 Excel 主体的列的单元格统一样式
     *
     * @param rowStyle 行样式
     */
    private void buildExcelBodyColumnStyle(RowStyle rowStyle) {
        for (int i = 0; i < cellFields.size(); i++) {
            CellField cellField = cellFields.get(i);
            CellStyle cellStyle = rowStyle.buildCellStyle(sxssfWorkbook);
            cellStyle.setAlignment(cellField.getAlign().getValue());
            String format = cellField.getFormat();
            DataFormat formatter = sxssfWorkbook.createDataFormat();
            short dataFormat = formatter.getFormat(format);
            cellStyle.setDataFormat(dataFormat);
            sxssfSheet.setDefaultColumnStyle(i, cellStyle);
            sxssfSheet.setDefaultRowHeightInPoints(rowStyle.getHeight());
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
        SXSSFRow row = sxssfSheet.createRow(index);
        row.setHeightInPoints(rowStyle.getHeight());
        BeanDescriptor beanDescriptor = new BeanDescriptor(obj);
        for (CellField cellField : cellFields) {
            SXSSFCell cell = row.createCell(cellField.getIndex());
            Object value = beanDescriptor.getValueByName(cellField.getField());
            setCellValue(cell, value, cellField.getType());
            cell.setCellStyle(sxssfSheet.getColumnStyle(cellField.getIndex()));
        }
    }

    /**
     * 设置单元格的值
     *
     * @param cell  单元格对象
     * @param value 值
     * @param type  值的类型
     */
    private void setCellValue(SXSSFCell cell, Object value, Class<?> type) {
        if (value == null) {
            cell.setCellValue("");
        } else if (type == Boolean.TYPE || type == Boolean.class) {
            boolean boolValue = (boolean) value;
            if (booleanStringMapping != null) {
                String booleanStr = booleanStringMapping.get(boolValue);
                if (booleanStr != null) {
                    cell.setCellValue(booleanStr);
                } else {
                    cell.setCellValue(boolValue);
                }
            } else {
                cell.setCellValue(boolValue);
            }
        } else if ((Number.class.isAssignableFrom(type) || type.isPrimitive()) && type != Byte.TYPE && type != Character.TYPE) {
            cell.setCellValue(Double.parseDouble(value.toString()));
        } else if (type == Date.class) {
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

}