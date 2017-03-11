package org.fanlychie.jexcel.annotation;

import org.fanlychie.jexcel.write.DataFormat;
import org.fanlychie.jreflect.BeanDescriptor;
import org.fanlychie.jreflect.FieldDescriptor;

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
     * 缓存
     */
    private static final Map<Class<?>, List<CellField>> CELL_FIELD_CACHE = new HashMap<>();

    /**
     * 解析类声明的 @Cell 注解
     *
     * @param targetClass 目标类
     * @return 返回解析出来的 CellField 数据列表
     */
    public static List<CellField> parseClass(Class<?> targetClass) {
        List<CellField> cellFields = CELL_FIELD_CACHE.get(targetClass);
        if (cellFields == null) {
            Map<Field, Cell> cellAnnotationMap = getCellAnnotationMap(targetClass);
            cellFields = convertCellAnnotationToCellFields(cellAnnotationMap);
            sortCellFieldIndex(cellFields);
            CELL_FIELD_CACHE.put(targetClass, cellFields);
        }
        return cellFields;
    }

    /**
     * 转换 @Cell 注解为 {@link CellField} 集合表示
     *
     * @param annotationMap @Cell 注解映射表
     * @return 返回 {@link CellField} 列表数据
     */
    private static List<CellField> convertCellAnnotationToCellFields(Map<Field, Cell> annotationMap) {
        List<CellField> cellFields = new ArrayList<>();
        for (Field field : annotationMap.keySet()) {
            Cell cell = annotationMap.get(field);
            CellField cellField = new CellField();
            cellField.setType(field.getType());
            cellField.setField(field.getName());
            cellField.setName(cell.name());
            cellField.setIndex(cell.index());
            cellField.setAlign(cell.align());
            String format = cell.format();
            if (format != null && !format.isEmpty()) {
                cellField.setFormat(format);
            } else {
                cellField.setFormat(DataFormat.getDefault(field.getType()));
            }
            cellFields.add(cellField);
        }
        return cellFields;
    }

    /**
     * 获取 @Cell 注解标注的 Map<字段对象, 注解对象>
     *
     * @param targetClass 目标类
     * @return 返回 Map<字段对象, 注解对象>
     */
    private static Map<Field, Cell> getCellAnnotationMap(Class<?> targetClass) {
        BeanDescriptor beanDescriptor = new BeanDescriptor(targetClass);
        FieldDescriptor fieldDescriptor = beanDescriptor.getFieldDescriptor();
        Map<Field, Cell> annotationMap = fieldDescriptor.getAnnotationsMap(Cell.class);
        if (annotationMap.isEmpty()) {
            throw new UnsupportedOperationException("you must mark the data field with the @Cell annotation in " + targetClass);
        }
        return annotationMap;
    }

    /**
     * 排序 {@link CellField} 的 index 顺序
     *
     * @param cellFields 单元格字段列表
     */
    private static void sortCellFieldIndex(List<CellField> cellFields) {
        Collections.sort(cellFields, new Comparator<CellField>() {
            @Override
            public int compare(CellField cf1, CellField cf2) {
                int i1 = cf1.getIndex();
                int i2 = cf2.getIndex();
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