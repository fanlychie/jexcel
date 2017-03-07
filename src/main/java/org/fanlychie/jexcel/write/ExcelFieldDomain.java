package org.fanlychie.jexcel.write;

/**
 * 字段域, 用于转换注解
 * Created by fanlychie on 2017/3/5.
 */
public class ExcelFieldDomain {

    /**
     * 单元格索引
     *
     * @return
     */
    private int index;

    /**
     * 单元格标题名称
     *
     * @return
     */
    private String name;

    /**
     * 数据格式
     *
     * @return
     */
    private String format;

    /**
     * 对齐方式
     *
     * @return
     */
    private Align align;

    /**
     * 字段名称
     */
    private String field;

    /**
     * 字段类型
     */
    private Class<?> type;

    /**
     * 外部不可访问
     */
    ExcelFieldDomain() {

    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public Align getAlign() {
        return align;
    }

    public void setAlign(Align align) {
        this.align = align;
    }

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public Class<?> getType() {
        return type;
    }

    public void setType(Class<?> type) {
        this.type = type;
    }

}