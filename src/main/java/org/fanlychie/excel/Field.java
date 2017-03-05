package org.fanlychie.excel;

import org.fanlychie.excel.write.Align;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 字段, 用于标记 Excel 单元格数据
 * Created by fanlychie on 2017/3/4.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Field {

    /**
     * 单元格索引, 从左至右数, 数值从0开始
     *
     * @return
     */
    int index();

    /**
     * 单元格标题名称
     *
     * @return
     */
    String name();

    /**
     * 数据格式
     *
     * @return
     */
    String format() default "";

    /**
     * 对齐方式
     *
     * @return
     */
    Align align() default Align.LEFT;

}