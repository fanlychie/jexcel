package org.fanlychie.jexcel.write.model;

import java.util.Arrays;
import java.util.List;

/**
 * 行
 * <p>
 * Created by fanlychie on 2018/3/15.
 */
public class SimpleRow {

    /**
     * 索引, 0表示从工作表的现有行的下一行开始
     */
    private int index;

    /**
     * 单元格
     */
    private List<SimpleCell> simpleCells;

    /**
     * 构造器
     *
     * @param index       索引, 0表示从工作表的现有行的下一行开始
     * @param simpleCells 单元格
     */
    public SimpleRow(int index, SimpleCell... simpleCells) {
        this(index, Arrays.asList(simpleCells));
    }

    /**
     * 构造器
     *
     * @param index       索引, 0表示从工作表的现有行的下一行开始
     * @param simpleCells 单元格
     */
    public SimpleRow(int index, List<SimpleCell> simpleCells) {
        this.index = index;
        this.simpleCells = simpleCells;
    }

    // ============== GETTERS AND SETTERS ====================

    public int getIndex() {
        return index;
    }

    public List<SimpleCell> getSimpleCells() {
        return simpleCells;
    }
}