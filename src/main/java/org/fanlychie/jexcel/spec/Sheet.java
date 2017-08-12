package org.fanlychie.jexcel.spec;

/**
 * 工作表
 * Created by fanzyun on 2017/8/12.
 */
public class Sheet {

    // 第一个工作表的索引
    private static final int FIRST_SHEET_INDEX = 1;

    // 工作表名称的模式串
    private static final String SHEET_NAME_PATTERN = "第%s页";

    // 工作表名称数据字典
    private static final String[] DATA_DICTIONARY = {"一", "二", "三", "四", "五", "六", "七", "八", "九", "十"};

    /**
     * 根据索引获取工作表的名称
     *
     * @param index 工作表的索引, 从 1 开始
     * @return 返回工作表的名称
     */
    public static String getName(int index, String initName) {
        if (index < 0) {
            index = FIRST_SHEET_INDEX;
        }
        if (initName != null) {
            return index == FIRST_SHEET_INDEX ? initName : initName + index;
        }
        if (index <= DATA_DICTIONARY.length) {
            return String.format(SHEET_NAME_PATTERN, DATA_DICTIONARY[index - 1]);
        }
        return String.format(SHEET_NAME_PATTERN, String.valueOf(index));
    }

}