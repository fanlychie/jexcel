package org.fanlychie.jexcel.spec;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 对齐方式枚举
 * Created by fanlychie on 2017/3/5.
 */
public enum Align {

    /**
     * 水平方向 - 左对齐
     */
    LEFT(CellStyle.ALIGN_LEFT),

    /**
     * 水平方向 - 右对齐
     */
    RIGHT(CellStyle.ALIGN_RIGHT),

    /**
     * 水平方向 - 居中对齐
     */
    CENTER(CellStyle.ALIGN_CENTER),

    /**
     * 垂直方向 - 居中对齐
     */
    VERTICAL_TOP(CellStyle.VERTICAL_TOP),

    /**
     * 垂直方向 - 居中对齐
     */
    VERTICAL_CENTER(CellStyle.VERTICAL_CENTER),

    /**
     * 垂直方向 - 居中对齐
     */
    VERTICAL_BOTTOM(CellStyle.VERTICAL_BOTTOM),

    ;

    private final short value;

    private Align(short value) {
        this.value = value;
    }

    public short getValue() {
        return value;
    }

}