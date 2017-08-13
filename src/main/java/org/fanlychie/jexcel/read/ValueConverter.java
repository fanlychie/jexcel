package org.fanlychie.jexcel.read;

import org.fanlychie.jexcel.exception.ExcelCastException;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 值转换器
 * Created by fanlychie on 2017/3/7.
 */
public final class ValueConverter {

    private static final Pattern DATE_STRING_REGEX = Pattern.compile("(\\d{4})(\\S)(\\d{1,2})(\\S)(\\d{1,2})([^ \\f\\n\\r\\t\\v\\d]?)");

    private static final Pattern TIME_STRING_REGEX = Pattern.compile("(\\d{1,2})(\\S)(\\d{1,2})(\\S)(\\d{1,2})([^ \\f\\n\\r\\t\\v\\d]?)");

    private static final Pattern DATETIME_STRING_REGEX = Pattern.compile(DATE_STRING_REGEX + "(\\s)" + TIME_STRING_REGEX);

    private static final Pattern TIMESTAMP_STRING_REGEX = Pattern.compile("[1-9]\\d{12,}");

    private static final ConcurrentHashMap<String, DateFormat> PATTERN_FORMAT = new ConcurrentHashMap<>();

    public static Object convertObjectValue(String value, Class<?> type) {
        if (value == null) {
            if (!type.isPrimitive()) {
                return null;
            } else {
                throw new ClassCastException("Cannot cast java.lang.String to " + type.getName());
            }
        }
        if (type == String.class) {
            return value;
        }
        if (type == Boolean.TYPE || type == Boolean.class) {
            return convertBooleanValue(value);
        }
        if (type == Date.class) {
            Matcher matcher = DATETIME_STRING_REGEX.matcher(value);
            if (matcher.matches()) {
                String pattern = matcher.replaceAll("yyyy$2MM$4dd$6$7HH$9mm$11ss$13");
                return parseStringToDate(value, pattern);
            }
            matcher = DATE_STRING_REGEX.matcher(value);
            if (matcher.matches()) {
                String pattern = matcher.replaceAll("yyyy$2MM$4dd$6");
                return parseStringToDate(value, pattern);
            }
            matcher = TIMESTAMP_STRING_REGEX.matcher(value);
            if (matcher.matches()) {
                return new Date(Long.parseLong(value));
            }
            matcher = TIME_STRING_REGEX.matcher(value);
            if (matcher.matches()) {
                String pattern = matcher.replaceAll("HH$2mm$4ss$6");
                return parseStringToDate(value, pattern);
            }
            throw new ClassCastException("can not parse \"" + value + "\" to java.util.Date");
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

    private static Date parseStringToDate(String value, String pattern) {
        if (!PATTERN_FORMAT.containsKey(pattern)) {
            PATTERN_FORMAT.put(pattern, new SimpleDateFormat(pattern));
        }
        try {
            return PATTERN_FORMAT.get(pattern).parse(value);
        } catch (ParseException e) {
            throw new ExcelCastException(e);
        }
    }

}