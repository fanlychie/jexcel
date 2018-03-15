package org.fanlychie.jexcel.write.model;

/**
 * 单元格
 * <p>
 * Created by fanlychie on 2018/3/15.
 */
public class SimpleCell {

    /**
     * 索引, 从0开始
     */
    private int index;

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

    // ============== GETTERS AND SETTERS ====================

    public int getIndex() {
        return index;
    }

    public Object getValue() {
        return value;
    }
}