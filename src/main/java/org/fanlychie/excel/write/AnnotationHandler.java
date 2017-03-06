package org.fanlychie.excel.write;

import org.fanlychie.excel.ExcelField;
import org.fanlychie.excel.exception.WriteExcelException;
import org.fanlychie.reflection.BeanDescriptor;
import org.fanlychie.reflection.FieldDescriptor;

import java.lang.reflect.Field;
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
    private static final Map<Class<?>, List<ExcelFieldDomain>> ANNOTATION_CACHE = new HashMap<>();

    /**
     * 解析类声明的 @Field 注解
     *
     * @param targetClass 目标类
     * @return 返回解析出来的 ExcelFieldDomain 数据列表
     */
    public static List<ExcelFieldDomain> parseClass(Class<?> targetClass) {
        List<ExcelFieldDomain> excelFieldDomains = ANNOTATION_CACHE.get(targetClass);
        if (excelFieldDomains == null) {
            Map<Field, ExcelField> fieldAnnotationMap = getFieldAnnotationMap(targetClass);
            excelFieldDomains = convertAnnotationToDomains(fieldAnnotationMap);
            sortExcelFieldDomains(excelFieldDomains);
            ANNOTATION_CACHE.put(targetClass, excelFieldDomains);
        }
        return excelFieldDomains;
    }

    /**
     * 转换注解为域的形式表示
     *
     * @param fieldAnnotationMap 字段属性注解表
     * @return 返回域列表数据
     */
    private static List<ExcelFieldDomain> convertAnnotationToDomains(Map<Field, ExcelField> fieldAnnotationMap) {
        List<ExcelFieldDomain> excelFieldDomains = new ArrayList<>();
        for (Field field : fieldAnnotationMap.keySet()) {
            ExcelField excelField = fieldAnnotationMap.get(field);
            ExcelFieldDomain excelFieldDomain = new ExcelFieldDomain();
            excelFieldDomain.setType(field.getType());
            excelFieldDomain.setField(field.getName());
            excelFieldDomain.setName(excelField.name());
            excelFieldDomain.setIndex(excelField.index());
            excelFieldDomain.setAlign(excelField.align());
            String format = excelField.format();
            if (format != null && !format.isEmpty()) {
                excelFieldDomain.setFormat(format);
            } else {
                excelFieldDomain.setFormat(DataFormat.getDefault(field.getType()));
            }
            excelFieldDomains.add(excelFieldDomain);
        }
        return excelFieldDomains;
    }

    /**
     * 获取字段属性注解的映射表
     *
     * @param targetClass 目标类
     * @return 返回 Map<字段对象, 注解对象>
     */
    private static Map<Field, ExcelField> getFieldAnnotationMap(Class<?> targetClass) {
        BeanDescriptor beanDescriptor = new BeanDescriptor(targetClass);
        FieldDescriptor fieldDescriptor = beanDescriptor.getFieldDescriptor();
        Map<Field, ExcelField> fieldAnnotationMap = fieldDescriptor.getAnnotationsMap(ExcelField.class);
        if (fieldAnnotationMap.isEmpty()) {
            throw new WriteExcelException("你必须在 " + targetClass.getName() + " 类中使用 @ExcelField 注解标注数据属性");
        }
        return fieldAnnotationMap;
    }

    /**
     * 排序字段域列表
     *
     * @param excelFieldDomains 字段域列表
     */
    private static void sortExcelFieldDomains(List<ExcelFieldDomain> excelFieldDomains) {
        Collections.sort(excelFieldDomains, new Comparator<ExcelFieldDomain>() {
            @Override
            public int compare(ExcelFieldDomain e1, ExcelFieldDomain e2) {
                int i1 = e1.getIndex();
                int i2 = e2.getIndex();
                return (i1 < i2) ? -1 : ((i1 == i2) ? 0 : 1);
            }
        });
    }

    /**
     * 私有化构造器
     */
    private AnnotationHandler() {

    }

}