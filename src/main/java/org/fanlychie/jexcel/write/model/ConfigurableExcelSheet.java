package org.fanlychie.jexcel.write.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.fanlychie.jexcel.spec.Align;
import org.fanlychie.jexcel.spec.Format;
import org.fanlychie.jexcel.write.RowStyle;
import org.fanlychie.jexcel.write.WritableSheet;
import org.fanlychie.jexcel.write.model.StyleConfig.RowStyleConfig;
import org.yaml.snakeyaml.Yaml;

import java.io.InputStream;

/**
 * 可配置的EXCEL工作表, 用于加载YAML配置文件初始化配置
 *
 * Created by fanlychie on 2018/3/14.
 */
public class ConfigurableExcelSheet implements ExcelSheet {

    /**
     * 配置文件流
     */
    private InputStream configFileInputStream;

    /**
     * 构造器
     *
     * @param configFile 类路径下的配置文件
     */
    public ConfigurableExcelSheet(String configFile) {
        this.configFileInputStream = ConfigurableExcelSheet.class.getResourceAsStream("/" + configFile);
    }

    /**
     * 构造器
     *
     * @param configFileInputStream 配置文件流
     */
    public ConfigurableExcelSheet(InputStream configFileInputStream) {
        this.configFileInputStream = configFileInputStream;
    }

    @Override
    public WritableSheet getWritableSheet() {
        // 样式配置
        StyleConfig config = new Yaml().loadAs(configFileInputStream, StyleConfig.class);
        // 工作表
        WritableSheet sheet = new WritableSheet();
        // 单元格宽度
        sheet.setCellWidth(config.getGlobal().getCellWidth());
        // 标题行样式
        sheet.setTitleRowStyle(createRowStyle(config.getTitleRow()));
        // 主体行样式
        sheet.setBodyRowStyle(createRowStyle(config.getBodyRow()));
        // 脚部行样式
        sheet.setFooterRowStyle(createRowStyle(config.getFooterRow()));
        // 关键字映射表
        sheet.setKeyMapping(config.getBodyRow().getKeyMapping());
        return sheet;
    }

    /**
     * 创建行样式
     *
     * @param config 配置对象
     * @return RowStyle
     */
    private RowStyle createRowStyle(RowStyleConfig config) {
        RowStyle rowStyle = new RowStyle();
        // 边框样式
        rowStyle.setBorder(CellStyle.BORDER_THIN, IndexedColors.GREY_25_PERCENT.index);
        if (config != null) {
            // 起始索引
            rowStyle.setIndex(config.getStartIndex());
            // 背景颜色
            rowStyle.setBackgroundColor(config.getBackgroundColor() == null ? null : IndexedColors.valueOf(
                    config.getBackgroundColor().toUpperCase()).index);
            // 字体样式
            rowStyle.setFont(config.getFontName(), config.getFontSize(), config.getFontColor() == null ? null :
                    IndexedColors.valueOf(config.getFontColor().toUpperCase()).index);
            // 高度
            rowStyle.setHeight(config.getHeight());
            // 水平方向对齐方式
            rowStyle.setAlignment(config.getAlign() == null ? null : Align.valueOf(config.getAlign().toUpperCase()).getValue());
            // 垂直方向对齐方式
            rowStyle.setVerticalAlignment(config.getVerticalAlign() == null ? null :
                    Align.valueOf("VERTICAL_" + config.getVerticalAlign().toUpperCase()).getValue());
            // 是否自动换行
            rowStyle.setWrapText(config.getWrapText());
            // 数据格式
            rowStyle.setFormat(config.getFormat() == null ? null : Format.valueOf(config.getFormat().toUpperCase()).getFormat());
        }
        return rowStyle;
    }

}