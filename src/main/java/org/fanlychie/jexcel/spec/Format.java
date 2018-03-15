package org.fanlychie.jexcel.spec;

import java.util.Date;

/**
 * 数据格式枚举
 * Created by fanlychie on 2017/3/4.
 */
public enum Format {

    /**
     * 缺省的文本格式
     */
    STRING("@"),

    /**
     * 缺省的整型格式
     */
    INTEGER("0"),

    /**
     * 缺省的小数格式
     */
    DECIMAL("0.00"),

    /**
     * 缺省的日期格式
     */
    DATE("yyyy-MM-dd"),

    /**
     * 缺省的日期时间格式
     */
    DATETIME("yyyy-MM-dd HH:mm:ss")

    ;

    private final String format;

    // 私有化构造
    private Format(String format) {
        this.format = format;
    }

    /**
     * 获取默认的数据格式
     *
     * @param type 数据类型
     * @return 返回参数类型对应的数据格式
     */
    public static String getDefault(Class<?> type) {
        if (type == String.class) {
            return STRING.format;
        }
        if (type == Short.TYPE || type == Short.class
                || type == Integer.TYPE || type == Integer.class
                || type == Long.TYPE || type == Long.class) {
            return INTEGER.format;
        }
        if (type == Float.TYPE || type == Float.class
                || type == Double.TYPE || type == Double.class) {
            return DECIMAL.format;
        }
        if (type == Date.class) {
            return DATETIME.format;
        }
        return STRING.format;
    }

    public String getFormat() {
        return format;
    }

}