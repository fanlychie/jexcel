package org.fanlychie.jexcel.write;

/**
 * 工作表
 * Created by fanlychie on 2017/3/4.
 */
public class Sheet {

    /**
     * 工作表名称
     */
    private String name;

    /**
     * 单元格宽度
     */
    private int cellWidth;

    /**
     * 标题行样式
     */
    private RowStyle titleRowStyle;

    /**
     * 主体行样式
     */
    private RowStyle bodyRowStyle;

    /**
     * 构建实例
     *
     * @param name 工作表名称
     */
    public Sheet(String name) {
        this.name = name;
    }

    /**
     * 设置单元格宽度
     *
     * @param cellWidth 单元格宽度
     */
    public void setCellWidth(int cellWidth) {
        this.cellWidth = cellWidth;
    }

    /**
     * 设置标题行样式
     *
     * @param titleRowStyle 标题行样式
     */
    public void setTitleRowStyle(RowStyle titleRowStyle) {
        this.titleRowStyle = titleRowStyle;
    }

    /**
     * 设置主体行样式
     *
     * @param bodyRowStyle 主体行样式
     */
    public void setBodyRowStyle(RowStyle bodyRowStyle) {
        this.bodyRowStyle = bodyRowStyle;
    }

    /**
     * 获取工作表名称
     *
     * @return 返回工作表名称
     */
    String getName() {
        return name;
    }

    /**
     * 获取单元格宽度
     *
     * @return 返回单元格宽度
     */
    int getCellWidth() {
        return cellWidth * 256 + 184;
    }

    /**
     * 获取标题行样式
     *
     * @return 返回标题行样式
     */
    RowStyle getTitleRowStyle() {
        return titleRowStyle;
    }

    /**
     * 获取主体行样式
     *
     * @return 返回主体行样式
     */
    RowStyle getBodyRowStyle() {
        return bodyRowStyle;
    }

}