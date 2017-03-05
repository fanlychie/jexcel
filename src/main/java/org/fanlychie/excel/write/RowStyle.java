package org.fanlychie.excel.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 行样式
 * Created by fanlychie on 2017/3/5.
 */
public class RowStyle {

    /**
     * 行的索引值
     */
    private int index;

    /**
     * 行的高度
     */
    private int height;

    /**
     * 水平方向对齐方式
     */
    private short alignment;

    /**
     * 垂直方向对齐方式
     */
    private short verticalAlignment;

    /**
     * 是否自动换行
     */
    private boolean wrapText;

    /**
     * 字体大小
     */
    private int fontSize;

    /**
     * 字体颜色
     */
    private short fontColor;

    /**
     * 边框
     */
    private short border;

    /**
     * 边框颜色
     */
    private short borderColor;

    /**
     * 背景颜色
     */
    private short backgroundColor;

    /**
     * 行单元格样式
     */
    private CellStyle cellStyle;

    /**
     * 设置行的起始索引
     *
     * @param index 起始索引, 索引值从0开始
     */
    public void setIndex(int index) {
        this.index = index;
    }

    /**
     * 设置行高
     *
     * @param height 行高
     */
    public void setHeight(int height) {
        this.height = height;
    }

    /**
     * 设置水平方向对齐方式
     *
     * @param alignment 水平方向对齐方式 {@link org.apache.poi.ss.usermodel.CellStyle}
     */
    public void setAlignment(short alignment) {
        this.alignment = alignment;
    }

    /**
     * 设置垂直方向对齐方式
     *
     * @param verticalAlignment 垂直方向对齐方式 {@link org.apache.poi.ss.usermodel.CellStyle}
     */
    public void setVerticalAlignment(short verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    /**
     * 设置是否自动换行
     *
     * @param wrapText 是否自动换行
     */
    public void setWrapText(boolean wrapText) {
        this.wrapText = wrapText;
    }

    /**
     * 设置行使用的字体
     *
     * @param fontSize  字体大小
     * @param fontColor 字体颜色 {@link org.apache.poi.ss.usermodel.IndexedColors}
     */
    public void setFont(int fontSize, short fontColor) {
        this.fontSize = fontSize;
        this.fontColor = fontColor;
    }

    /**
     * 设置边框样式
     *
     * @param border      边框 {@link org.apache.poi.ss.usermodel.CellStyle}
     * @param borderColor 边框样式 {@link org.apache.poi.ss.usermodel.IndexedColors}
     */
    public void setBorder(short border, short borderColor) {
        this.border = border;
        this.borderColor = borderColor;
    }

    /**
     * 设置背景颜色
     *
     * @param backgroundColor 背景颜色 {@link org.apache.poi.ss.usermodel.IndexedColors}
     */
    public void setBackgroundColor(short backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    /**
     * 获取行索引
     *
     * @return 返回行的索引值
     */
    int getIndex() {
        return index;
    }

    /**
     * 获取行高
     *
     * @return 返回行的高度
     */
    int getHeight() {
        return height;
    }

    /**
     * 获取行的单元格样式
     *
     * @param workbook 工作薄
     * @return 返回单元格样式
     */
    CellStyle getCellStyle(XSSFWorkbook workbook) {
        return getCellStyle(workbook, false);
    }

    /**
     * 获取行的单元格样式
     *
     * @param workbook 工作薄
     * @param recreate 重新创建
     * @return 返回单元格样式
     */
    CellStyle getCellStyle(XSSFWorkbook workbook, boolean recreate) {
        if (cellStyle == null || recreate) {
            cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(alignment);
            cellStyle.setVerticalAlignment(verticalAlignment);
            cellStyle.setFont(createFont(workbook));
            cellStyle.setWrapText(wrapText);
            cellStyle.setBorderBottom(border);
            cellStyle.setBorderLeft(border);
            cellStyle.setBorderRight(border);
            cellStyle.setBorderTop(border);
            cellStyle.setBottomBorderColor(borderColor);
            cellStyle.setLeftBorderColor(borderColor);
            cellStyle.setRightBorderColor(borderColor);
            cellStyle.setTopBorderColor(borderColor);
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(backgroundColor);
        }
        return cellStyle;
    }

    /**
     * 创建字体
     *
     * @return {@link org.apache.poi.ss.usermodel.Font}
     */
    private Font createFont(XSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setColor(fontColor);
        font.setFontHeightInPoints((short) fontSize);
        return font;
    }

}