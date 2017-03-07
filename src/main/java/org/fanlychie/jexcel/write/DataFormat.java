package org.fanlychie.jexcel.write;

import java.util.Date;

/**
 * 数据格式
 * Created by fanlychie on 2017/3/4.
 */
public final class DataFormat {

    /**
     * 缺省的文本格式
     */
    static final String STRING_FORMAT = "@";

    /**
     * 缺省的整型格式
     */
    private static final String INTEGER_FORMAT = "0";

    /**
     * 缺省的小数格式
     */
    private static final String DECIMAL_FORMAT = "0.00";

    /**
     * 缺省的日期格式
     */
    private static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    /**
     * 缺省的类型数据格式
     */
    private static final String DEFAULT_TYPE_DATA_FORMAT = STRING_FORMAT;

    // 私有化构造
    private DataFormat() {

    }

    /**
     * 获取默认的数据格式
     *
     * @param type 数据类型
     * @return 返回参数类型对应的数据格式
     */
    public static String getDefault(Class<?> type) {
        if (type == Short.TYPE || type == Short.class || type == Integer.TYPE || type == Integer.class || type == Long.TYPE || type == Long.class) {
            return INTEGER_FORMAT;
        }
        if (type == Float.TYPE || type == Float.class || type == Double.TYPE || type == Double.class) {
            return DECIMAL_FORMAT;
        }
        if (type == String.class) {
            return STRING_FORMAT;
        }
        if (type == Date.class) {
            return DATE_FORMAT;
        }
        return DEFAULT_TYPE_DATA_FORMAT;
    }

}