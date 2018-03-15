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
 * Created by fanlychie on 2018/3/14.
 */
public class ConfigurableExcelSheet implements ExcelSheet {

    private Class<?> dataClassType;

    private InputStream configFileInputStream;

    public ConfigurableExcelSheet(Class<?> dataClassType, String configFile) {
        this.dataClassType = dataClassType;
        this.configFileInputStream = ConfigurableExcelSheet.class.getResourceAsStream("/" + configFile);
    }

    public ConfigurableExcelSheet(Class<?> dataClassType, InputStream configFileInputStream) {
        this.dataClassType = dataClassType;
        this.configFileInputStream = configFileInputStream;
    }

    @Override
    public WritableSheet getWritableSheet() {
        StyleConfig config = new Yaml().loadAs(configFileInputStream, StyleConfig.class);
        WritableSheet sheet = new WritableSheet(dataClassType);
        sheet.setCellWidth(config.getGlobal().getCellWidth());
        sheet.setTitleRowStyle(createRowStyle(config.getTitleRow()));
        sheet.setBodyRowStyle(createRowStyle(config.getBodyRow()));
        sheet.setFooterRowStyle(createRowStyle(config.getFootRow()));
        return sheet;
    }

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