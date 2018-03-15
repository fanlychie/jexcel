package org.fanlychie.jexcel.write.model;

/**
 * 样式配置
 * Created by fanlychie on 2018/3/14.
 */
public class StyleConfig {

    private GlobalStyleConfig global;

    private RowStyleConfig titleRow;

    private RowStyleConfig bodyRow;

    private RowStyleConfig footRow;

    /**
     * 全局配置
     */
    public static class GlobalStyleConfig {

        private Integer cellWidth;

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

        private Integer startIndex;

        private Integer height;

        private String fontName;

        private Integer fontSize;

        private String fontColor;

        private Boolean wrapText;

        private String backgroundColor;

        private String align;

        private String verticalAlign;

        private String format;

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
    }

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

    public RowStyleConfig getFootRow() {
        return footRow;
    }

    public void setFootRow(RowStyleConfig footRow) {
        this.footRow = footRow;
    }

}