package org.fanlychie.jexcel.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * 行样式
 * Created by fanlychie on 2017/3/5.
 */
public class RowStyle {

    /**
     * 行的索引值
     */
    private Integer index;

    /**
     * 行的高度
     */
    private Integer height;

    /**
     * 水平方向对齐方式
     */
    private Short alignment;

    /**
     * 垂直方向对齐方式
     */
    private Short verticalAlignment;

    /**
     * 是否自动换行
     */
    private Boolean wrapText;

    /**
     * 字体名称
     */
    private String fontName;

    /**
     * 字体大小
     */
    private Integer fontSize;

    /**
     * 字体颜色
     */
    private Short fontColor;

    /**
     * 边框
     */
    private Short border;

    /**
     * 边框颜色
     */
    private Short borderColor;

    /**
     * 背景颜色
     */
    private Short backgroundColor;

    /**
     * 行数据格式
     */
    private String format;

    /**
     * 设置行的起始索引
     *
     * @param index 起始索引, 索引值从0开始
     */
    public void setIndex(Integer index) {
        this.index = index;
    }

    /**
     * 设置行高
     *
     * @param height 行高
     */
    public void setHeight(Integer height) {
        this.height = height;
    }

    /**
     * 设置水平方向对齐方式
     *
     * @param alignment 水平方向对齐方式 {@link org.apache.poi.ss.usermodel.CellStyle}
     */
    public void setAlignment(Short alignment) {
        this.alignment = alignment;
    }

    /**
     * 设置垂直方向对齐方式
     *
     * @param verticalAlignment 垂直方向对齐方式 {@link org.apache.poi.ss.usermodel.CellStyle}
     */
    public void setVerticalAlignment(Short verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    /**
     * 设置是否自动换行
     *
     * @param wrapText 是否自动换行
     */
    public void setWrapText(Boolean wrapText) {
        this.wrapText = wrapText;
    }

    /**
     * 设置行使用的字体
     *
     * @param fontSize  字体大小
     * @param fontColor 字体颜色 {@link org.apache.poi.ss.usermodel.IndexedColors}
     */
    public void setFont(String fontName, Integer fontSize, Short fontColor) {
        this.fontName = fontName;
        this.fontSize = fontSize;
        this.fontColor = fontColor;
    }

    /**
     * 设置边框样式
     *
     * @param border      边框 {@link org.apache.poi.ss.usermodel.CellStyle}
     * @param borderColor 边框样式 {@link org.apache.poi.ss.usermodel.IndexedColors}
     */
    public void setBorder(Short border, Short borderColor) {
        this.border = border;
        this.borderColor = borderColor;
    }

    /**
     * 设置背景颜色
     *
     * @param backgroundColor 背景颜色 {@link org.apache.poi.ss.usermodel.IndexedColors}
     */
    public void setBackgroundColor(Short backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    /**
     * 设置行数据格式
     *
     * @param format 行数据格式
     */
    public void setFormat(String format) {
        this.format = format;
    }

    /**
     * 获取行索引
     *
     * @return 返回行的索引值
     */
    Integer getIndex() {
        return index;
    }

    /**
     * 获取行高
     *
     * @return 返回行的高度
     */
    Integer getHeight() {
        return height;
    }

    /**
     * 构建行的单元格样式
     *
     * @param workbook 工作薄
     * @return 返回单元格样式
     */
    CellStyle buildCellStyle(SXSSFWorkbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        if (alignment != null) {
            cellStyle.setAlignment(alignment);
        }
        if (verticalAlignment != null) {
            cellStyle.setVerticalAlignment(verticalAlignment);
        }
        if (fontSize != null || fontColor != null || fontName != null) {
            cellStyle.setFont(createFont(workbook));
        }
        if (wrapText != null) {
            cellStyle.setWrapText(wrapText);
        }
        if (border != null && borderColor != null) {
            cellStyle.setBorderTop(border);
            cellStyle.setTopBorderColor(borderColor);
            cellStyle.setBorderLeft(border);
            cellStyle.setLeftBorderColor(borderColor);
            cellStyle.setBorderRight(border);
            cellStyle.setRightBorderColor(borderColor);
            cellStyle.setBorderBottom(border);
            cellStyle.setBottomBorderColor(borderColor);
        }
        if (backgroundColor != null) {
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(backgroundColor);
        }
        if (format != null) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(format));
        }
        return cellStyle;
    }

    /**
     * 创建字体
     *
     * @return {@link org.apache.poi.ss.usermodel.Font}
     */
    private Font createFont(SXSSFWorkbook workbook) {
        Font font = workbook.createFont();
        if (fontColor != null) {
            font.setColor(fontColor);
        }
        if (fontSize != null) {
            font.setFontHeightInPoints(fontSize.shortValue());
        }
        if (fontName != null) {
            font.setFontName(fontName);
        }
        return font;
    }

}