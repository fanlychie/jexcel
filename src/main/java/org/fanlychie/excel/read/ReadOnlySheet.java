package org.fanlychie.excel.read;

/**
 * 只读的工作表
 * Created by fanlychie on 2017/3/6.
 */
public class ReadOnlySheet {

    /**
     * 工作表索引
     */
    private int index;

    /**
     * 工作表名称
     */
    private String name;

    /**
     * 起始行号
     */
    private int firstRowNum;

    /**
     * 结束行号
     */
    private int lastRowNum;

    /**
     * 设置读取的工作表索引, 索引参数和名称参数二选一, 名称参数优先处理, 若没有, 则使用此项参数处理. {@link #setName(String)}
     *
     * @param index 读取的工作表索引值, 数值从0开始
     */
    public void setIndex(int index) {
        this.index = index;
    }

    /**
     * 设置读取的工作表名称, 索引参数和名称参数二选一, 名称参数优先处理, 若没有, 则使用索引参数处理. {@link #setIndex(int)}
     *
     * @param name 读取的工作表名称
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * 设置读取的起始行号
     *
     * @param firstRowNum 读取的起始行号, 数值从1开始
     */
    public void setFirstRowNum(int firstRowNum) {
        this.firstRowNum = firstRowNum;
    }

    /**
     * 设置读取结束的行号, 若不设置, 则读取到 Excel 表的最后一行内容
     *
     * @param lastRowNum 读取结束的行号, 数值从1开始
     */
    public void setLastRowNum(int lastRowNum) {
        this.lastRowNum = lastRowNum;
    }

    /**
     * 获取读取的工作表名称
     *
     * @return 返回读取的工作表名称
     */
    public String getName() {
        return name;
    }

    /**
     * 获取读取的工作表索引
     *
     * @return 返回读取的工作表索引值
     */
    public int getIndex() {
        return index;
    }

    /**
     * 获取读取的起始行号
     *
     * @return 返回读取的起始行号
     */
    public int getFirstRowNum() {
        return firstRowNum;
    }

    /**
     * 获取读取结束的行号
     *
     * @return 返回读取结束的行号
     */
    public int getLastRowNum() {
        return lastRowNum;
    }

}