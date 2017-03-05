package org.fanlychie.excel.write;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 对齐方式枚举
 * Created by fanlychie on 2017/3/5.
 */
public enum Align {

    /**
     * 左对齐
     */
    LEFT(CellStyle.ALIGN_LEFT),

    /**
     * 右对齐
     */
    RIGHT(CellStyle.ALIGN_RIGHT),

    /**
     * 居中对齐
     */
    CENTER(CellStyle.ALIGN_CENTER),

    ;

    private final short value;

    private Align(short value) {
        this.value = value;
    }

    public short getValue() {
        return value;
    }

}