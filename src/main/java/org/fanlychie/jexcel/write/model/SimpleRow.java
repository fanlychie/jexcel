package org.fanlychie.jexcel.write.model;

import java.util.Arrays;
import java.util.List;

/**
 * 行
 * Created by fanlychie on 2018/3/15.
 */
public class SimpleRow {

    /**
     * 单元格
     */
    private List<SimpleCell> simpleCells;

    /**
     * 构造器, 没有指定特定的行索引, 则紧接着当前行
     *
     * @param simpleCells 单元格
     */
    public SimpleRow(SimpleCell... simpleCells) {
        this(Arrays.asList(simpleCells));
    }

    /**
     * 构造器, 没有指定特定的行索引, 则紧接着当前行
     *
     * @param simpleCells 单元格
     */
    public SimpleRow(List<SimpleCell> simpleCells) {
        this.simpleCells = simpleCells;
    }

    // ============== GETTERS AND SETTERS ====================

    public List<SimpleCell> getSimpleCells() {
        return simpleCells;
    }
}