package org.fanlychie.jexcel.annotation;

import org.fanlychie.jexcel.spec.Align;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 注解, 用于标注单元格数据字段
 * Created by fanlychie on 2017/3/4.
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Cell {

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