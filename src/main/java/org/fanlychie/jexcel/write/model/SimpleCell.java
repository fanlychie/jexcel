package org.fanlychie.jexcel.write.model;

import org.fanlychie.jexcel.exception.WriteExcelException;

/**
 * 单元格
 * Created by fanlychie on 2018/3/15.
 */
public class SimpleCell {

    /**
     * 列索引, 从0开始
     */
    private int index;

    /**
     * 用来指定单元格横向跨越的列数
     */
    private Integer colspan;

    /**
     * 用来指定单元格纵向跨越的行数
     */
    private Integer rowspan;

    /**
     * 单元格的值
     */
    private Object value;

    /**
     * 构造器
     *
     * @param index 索引, 从0开始
     * @param value 单元格的值
     */
    public SimpleCell(int index, Object value) {
        this.index = index;
        this.value = value;
    }

    /**
     * 构造器
     *
     * @param index   索引, 从0开始
     * @param colspan 用来指定单元格横向跨越的列数
     * @param rowspan 用来指定单元格纵向跨越的行数
     * @param value   单元格的值
     */
    public SimpleCell(int index, int colspan, int rowspan, Object value) {
        if (colspan < 1) {
            throw new WriteExcelException("colspan cant not be less than 1");
        }
        if (rowspan < 1) {
            throw new WriteExcelException("rowspan cant not be less than 1");
        }
        this.index = index;
        this.value = value;
        this.colspan = colspan;
        this.rowspan = rowspan;
    }

    // ============== GETTERS AND SETTERS ====================

    public int getIndex() {
        return index;
    }

    public Object getValue() {
        return value;
    }

    public Integer getColspan() {
        return colspan;
    }

    public Integer getRowspan() {
        return rowspan;
    }
}