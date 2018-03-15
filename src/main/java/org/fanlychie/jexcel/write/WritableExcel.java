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
import org.fanlychie.jexcel.write.model.ExcelSheet;
import org.fanlychie.jexcel.write.model.SimpleCell;
import org.fanlychie.jexcel.write.model.SimpleRow;
import org.fanlychie.jreflect.BeanDescriptor;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
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
     * 工作表计数
     */
    private int sheetCount = 1;

    /**
     * 边界行索引值
     */
    private int boundRowIndex;

    /**
     * 脚部的单元格样式
     */
    private CellStyle footerCellStyle;

    /**
     * 关键字映射
     */
    private Map<Object, Object> keyMapping;

    /**
     * 构建实例
     *
     * @param excelSheet    工作表
     */
    public WritableExcel(ExcelSheet excelSheet) {
        this.writableSheet = excelSheet.getWritableSheet();
        this.sxssfWorkbook = new SXSSFWorkbook();
        this.keyMapping = writableSheet.getKeyMapping();
    }

    /**
     * 添加一个工作表
     *
     * @param data 数据列表
     * @return 返回当前对象
     */
    public WritableExcel addSheet(List<?> data) {
        return addSheet(writableSheet.getName() + (sheetCount++), data);
    }

    /**
     * 添加一个工作表
     *
     * @param sheetName 工作表名称
     * @param data      数据列表
     * @return 返回当前对象
     */
    public WritableExcel addSheet(String sheetName, Collection<?> data) {
        try {
            // 创建工作表
            sxssfSheet = sxssfWorkbook.createSheet(sheetName);
            // 处理工作表数据
            if (data != null) {
                Iterator<?> iterator = data.iterator();
                if (iterator.hasNext()) {
                    // 行数据对象
                    Object rowObj = iterator.next();
                    // 边界索引
                    boundRowIndex += data.size() + 1;
                    // 单元格字段解析
                    cellFields = AnnotationHandler.parseClass(rowObj.getClass());
                    // 行样式
                    RowStyle bodyRowStyle = writableSheet.getBodyRowStyle();
                    // 构建标题行
                    buildExcelTitleRow();
                    // 构建主体行样式
                    buildExcelBodyColumnStyle(bodyRowStyle);
                    // 主体索引
                    int bodyIndex = bodyRowStyle.getIndex();
                    do {
                        // 构建行数据
                        buildExcelBodyRow(bodyRowStyle, bodyIndex++, rowObj);
                    } while (iterator.hasNext() && (rowObj = iterator.next()) != null);
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
     * @param simpleRow 行
     * @return 返回当前对象
     */
    public WritableExcel addRow(SimpleRow simpleRow) {
        return addRow(simpleRow, null);
    }

    /**
     * 添加一行
     *
     * @param simpleRow 行
     * @return 返回当前对象
     */
    public WritableExcel addFooterRow(SimpleRow simpleRow) {
        return addRow(simpleRow, writableSheet.getFooterRowStyle());
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
     * 写出到客户端响应, 用于供客户端下载文件
     *
     * @param response HttpServletResponse
     * @param filename 下载时存储的文件名称
     */
    public void toHttpResponse(HttpServletResponse response, String filename) {
        try {
            filename = new String(filename.getBytes("UTF-8"), "ISO-8859-1");
        } catch (UnsupportedEncodingException e) {
            throw new ExcelCastException(e);
        }
        response.reset();
        response.setHeader("Content-Disposition", "attachment; filename=" + filename);
        response.setContentType("application/octet-stream; charset=ISO-8859-1");
        try {
            sxssfWorkbook.write(response.getOutputStream());
        } catch (IOException e) {
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

    private WritableExcel addRow(SimpleRow simpleRow, RowStyle rowStyle) {
        if (boundRowIndex == 0) {
            throw new IllegalStateException();
        }
        boundRowIndex += simpleRow.getIndex();
        SXSSFRow row = sxssfSheet.createRow(boundRowIndex++);
        if (writableSheet.getBodyRowStyle().getHeight() != null) {
            row.setHeightInPoints(writableSheet.getBodyRowStyle().getHeight());
        }
        List<SimpleCell> simpleCells = simpleRow.getSimpleCells();
        for (SimpleCell simpleCell : simpleCells) {
            SXSSFCell cell = row.createCell(simpleCell.getIndex());
            setCellValue(cell, simpleCell.getValue(), simpleCell.getValue().getClass());
            if (rowStyle != null) {
                cell.setCellStyle(rowStyle.buildCellStyle(sxssfWorkbook));
            }
        }
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
        if (rowStyle.getHeight() != null) {
            row.setHeightInPoints(rowStyle.getHeight());
        }
        for (CellField cellField : cellFields) {
            int index = cellField.getIndex();
            if (writableSheet.getCellWidth() != null) {
                sxssfSheet.setColumnWidth(index, writableSheet.getCellWidth());
            }
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
            if (rowStyle.getHeight() != null) {
                sxssfSheet.setDefaultRowHeightInPoints(rowStyle.getHeight());
            }
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
        if (rowStyle.getHeight() != null) {
            row.setHeightInPoints(rowStyle.getHeight());
        }
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
        }
        else if ((type == Boolean.TYPE || type == Boolean.class) && (keyMapping == null || !keyMapping.containsKey(value))) {
            cell.setCellValue((boolean) value);
        }
        else if ((Number.class.isAssignableFrom(value.getClass()))) {
            cell.setCellValue(Double.parseDouble(value.toString()));
        }
        else if (type == Date.class) {
            cell.setCellValue((Date) value);
        }
        else {
            String cellValue = value.toString();
            if (keyMapping != null && keyMapping.containsKey(value)) {
                cellValue = keyMapping.get(value).toString();
            }
            cell.setCellValue(cellValue);
        }
    }

}