package org.fanlychie.jexcel.write.model;

/**
 * Created by fanlychie on 2018/3/14.
 */
public class CustomExcelSheet extends ConfigurableExcelSheet {

    public CustomExcelSheet(Class<?> dataClassType) {
        super(dataClassType, "jexcel-custom.yml");
    }

}