package org.fanlychie.excel.write;

import org.fanlychie.excel.Field;
import org.fanlychie.excel.exception.WriteExcelException;
import org.fanlychie.reflection.BeanDescriptor;
import org.fanlychie.reflection.FieldDescriptor;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 注解解析器
 * Created by fanlychie on 2017/3/5.
 */
public final class AnnotationHandler {

    /**
     * 注解缓存
     */
    private static final Map<Class<?>, List<FieldDomain>> ANNOTATION_CACHE = new HashMap<>();

    /**
     * 解析类声明的 @Field 注解
     *
     * @param targetClass 目标类
     * @return 返回解析出来的 FieldDomain 数据列表
     */
    public static List<FieldDomain> parseClass(Class<?> targetClass) {
        List<FieldDomain> fields = ANNOTATION_CACHE.get(targetClass);
        if (fields == null) {
            fields = new ArrayList<>();
            BeanDescriptor beanDescriptor = new BeanDescriptor(targetClass);
            FieldDescriptor fieldDescriptor = beanDescriptor.getFieldDescriptor();
            Map<java.lang.reflect.Field, Field> map = fieldDescriptor.getAnnotationsMap(Field.class);
            if (map.isEmpty()) {
                throw new WriteExcelException("你必须在 " + targetClass.getName() + " 类中使用 @Field 注解标注数据属性");
            }
            for (java.lang.reflect.Field key : map.keySet()) {
                Field field = map.get(key);
                FieldDomain domain = new FieldDomain();
                domain.setName(field.name());
                domain.setIndex(field.index());
                domain.setAlign(field.align());
                domain.setField(key.getName());
                String format = field.format();
                if (format != null && !format.isEmpty()) {
                    domain.setFormat(format);
                } else {
                    domain.setFormat(DataFormat.getDefault(key.getType()));
                }
                fields.add(domain);
            }
            Collections.sort(fields, new Comparator<FieldDomain>() {
                @Override
                public int compare(FieldDomain f1, FieldDomain f2) {
                    int index1 = f1.getIndex();
                    int index2 = f2.getIndex();
                    return (index1 < index2) ? -1 : ((index1 == index2) ? 0 : 1);
                }
            });
            ANNOTATION_CACHE.put(targetClass, fields);
        }
        return fields;
    }

    /**
     * 私有化
     */
    private AnnotationHandler() {

    }

}