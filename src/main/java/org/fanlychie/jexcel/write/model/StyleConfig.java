package org.fanlychie.jexcel.write.model;

import java.util.Map;

/**
 * 样式配置
 *
 * Created by fanlychie on 2018/3/14.
 */
public class StyleConfig {

    /**
     * 全局配置
     */
    private GlobalStyleConfig global;

    /**
     * 标题行配置
     */
    private RowStyleConfig titleRow;

    /**
     * 主体行配置
     */
    private RowStyleConfig bodyRow;

    /**
     * 脚部行配置
     */
    private RowStyleConfig footerRow;

    /**
     * 全局配置
     */
    public static class GlobalStyleConfig {

        /**
         * 单元格宽度
         */
        private Integer cellWidth;

        // ============== GETTERS AND SETTERS ====================

        public Integer getCellWidth() {
            return cellWidth;
        }

        public void setCellWidth(Integer cellWidth) {
            this.cellWidth = cellWidth;
        }

    }

    /**
     * 行配置
     */
    public static class RowStyleConfig {

        /**
         * 起始索引, 从0开始
         */
        private Integer startIndex;

        /**
         * 行高
         */
        private Integer height;

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
        private String fontColor;

        /**
         * 是否自动换行
         */
        private Boolean wrapText;

        /**
         * 背景颜色
         */
        private String backgroundColor;

        /**
         * 水平方向对齐方式
         */
        private String align;

        /**
         * 垂直方向对齐方式
         */
        private String verticalAlign;

        /**
         * 单元格格式
         */
        private String format;

        /**
         * 关键字映射
         */
        private Map<Object, Object> keyMapping;

        // ============== GETTERS AND SETTERS ====================

        public void setStartIndex(Integer startIndex) {
            this.startIndex = startIndex;
        }

        public void setHeight(Integer height) {
            this.height = height;
        }

        public void setFontSize(Integer fontSize) {
            this.fontSize = fontSize;
        }

        public void setWrapText(Boolean wrapText) {
            this.wrapText = wrapText;
        }

        public Integer getStartIndex() {
            return startIndex;
        }

        public Integer getHeight() {
            return height;
        }

        public Integer getFontSize() {
            return fontSize;
        }

        public Boolean getWrapText() {
            return wrapText;
        }

        public String getFontName() {
            return fontName;
        }

        public void setFontName(String fontName) {
            this.fontName = fontName;
        }

        public String getFontColor() {
            return fontColor;
        }

        public void setFontColor(String fontColor) {
            this.fontColor = fontColor;
        }

        public String getBackgroundColor() {
            return backgroundColor;
        }

        public void setBackgroundColor(String backgroundColor) {
            this.backgroundColor = backgroundColor;
        }

        public String getAlign() {
            return align;
        }

        public void setAlign(String align) {
            this.align = align;
        }

        public String getVerticalAlign() {
            return verticalAlign;
        }

        public void setVerticalAlign(String verticalAlign) {
            this.verticalAlign = verticalAlign;
        }

        public String getFormat() {
            return format;
        }

        public void setFormat(String format) {
            this.format = format;
        }

        public Map<Object, Object> getKeyMapping() {
            return keyMapping;
        }

        public void setKeyMapping(Map<Object, Object> keyMapping) {
            this.keyMapping = keyMapping;
        }
    }

    // ============== GETTERS AND SETTERS ====================

    public GlobalStyleConfig getGlobal() {
        return global;
    }

    public void setGlobal(GlobalStyleConfig global) {
        this.global = global;
    }

    public RowStyleConfig getTitleRow() {
        return titleRow;
    }

    public void setTitleRow(RowStyleConfig titleRow) {
        this.titleRow = titleRow;
    }

    public RowStyleConfig getBodyRow() {
        return bodyRow;
    }

    public void setBodyRow(RowStyleConfig bodyRow) {
        this.bodyRow = bodyRow;
    }

    public RowStyleConfig getFooterRow() {
        return footerRow;
    }

    public void setFooterRow(RowStyleConfig footerRow) {
        this.footerRow = footerRow;
    }
}