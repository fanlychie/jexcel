package org.fanlychie.jexcel.read;

/**
 * 值转换器
 * Created by fanlychie on 2017/3/7.
 */
public final class ValueConverter {

    public static Object convertObjectValue(String value, Class<?> type) {
        if (type == String.class) {
            return value;
        }
        if (type == Boolean.TYPE || type == Boolean.class) {
            return convertBooleanValue(value);
        }
        Double doubleValue = Double.valueOf(value);
        if (type == Byte.TYPE || type == Byte.class) {
            return doubleValue.byteValue();
        }
        if (type == Short.TYPE || type == Short.class) {
            return doubleValue.shortValue();
        }
        if (type == Integer.TYPE || type == Integer.class) {
            return doubleValue.intValue();
        }
        if (type == Long.TYPE || type == Long.class) {
            return doubleValue.longValue();
        }
        if (type == Float.TYPE || type == Float.class) {
            return doubleValue.floatValue();
        }
        if (type == Double.TYPE || type == Double.class) {
            return doubleValue;
        }
        throw new ClassCastException("Cannot cast java.lang.String to " + type.getName());
    }

    private static boolean convertBooleanValue(String value) {
        if (value == "1" || value.equals("是")
                || value.equalsIgnoreCase("Y") || value.equalsIgnoreCase("YES")
                || value.equalsIgnoreCase("T") || value.equalsIgnoreCase("TRUE")) {
            return true;
        }
        if (value == "0" || value.equals("否")
                || value.equalsIgnoreCase("N") || value.equalsIgnoreCase("NO")
                || value.equalsIgnoreCase("F") || value.equalsIgnoreCase("FALSE")) {
            return false;
        }
        throw new ClassCastException("cannot cast java.lang.String to boolean");
    }

}