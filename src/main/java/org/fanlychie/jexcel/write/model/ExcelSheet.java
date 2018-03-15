package org.fanlychie.jexcel.write.model;

import org.fanlychie.jexcel.write.WritableSheet;

/**
 * Excel工作表
 *
 * Created by fanlychie on 2018/3/14.
 */
public interface ExcelSheet {

    /**
     * 获取工作表
     *
     * @return WritableSheet
     */
    WritableSheet getWritableSheet();

}