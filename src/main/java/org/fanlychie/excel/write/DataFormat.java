package org.fanlychie.excel.write;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * 数据格式
 * Created by fanlychie on 2017/3/4.
 */
public final class DataFormat {

    /**
     * 缺省的类型数据格式
     */
    static final String DEFAULT_TYPE_DATA_FORMAT = "GENERAL";

    /**
     * 默认的数据格式表
     */
    private static final Map<Class<?>, String> DEFAULT_DATA_FORMAT = new HashMap<>();

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
        String format = DEFAULT_DATA_FORMAT.get(type);
        return format == null ? DEFAULT_TYPE_DATA_FORMAT : format;
    }

    static {
        DEFAULT_DATA_FORMAT.put(Short.TYPE, "0");
        DEFAULT_DATA_FORMAT.put(Short.class, "0");
        DEFAULT_DATA_FORMAT.put(Integer.TYPE, "0");
        DEFAULT_DATA_FORMAT.put(Integer.class, "0");
        DEFAULT_DATA_FORMAT.put(Long.TYPE, "0");
        DEFAULT_DATA_FORMAT.put(Long.class, "0");
        DEFAULT_DATA_FORMAT.put(Float.TYPE, "0.00");
        DEFAULT_DATA_FORMAT.put(Float.class, "0.00");
        DEFAULT_DATA_FORMAT.put(Double.TYPE, "0.00");
        DEFAULT_DATA_FORMAT.put(Double.class, "0.00");
        DEFAULT_DATA_FORMAT.put(String.class, "GENERAL");
        DEFAULT_DATA_FORMAT.put(Date.class, "yyyy-MM-dd HH:mm:ss");
    }

}