package org.fanlychie.jexcel.write.model;

/**
 * Created by fanlychie on 2018/3/14.
 */
public class DefaultExcelSheet extends ConfigurableExcelSheet {

    public DefaultExcelSheet(Class<?> dataClassType) {
        super(dataClassType, "jexcel-default.yml");
    }

}