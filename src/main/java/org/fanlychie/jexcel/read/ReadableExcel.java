package org.fanlychie.jexcel.read;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.StylesTable;
import org.fanlychie.jexcel.annotation.AnnotationHandler;
import org.fanlychie.jexcel.annotation.CellField;
import org.fanlychie.jexcel.exception.ExcelCastException;
import org.fanlychie.jexcel.exception.ReadExcelException;
import org.fanlychie.jreflect.BeanDescriptor;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * 可读的 Excel, 用于读取 Excel 文件
 * Created by fanlychie on 2017/3/5.
 */
public class ReadableExcel {

    private StylesTable stylesTable;

    private SheetIterator sheetIterator;

    private ReadOnlySharedStringsTable sharedStringsTable;

    private int startRow;

    private Map<Integer, CellField> cellFieldMap;

    private BeanDescriptor beanDescriptor;

    /**
     * 构建一个可读的 Excel 对象
     *
     * @param excelFile Excel 文件
     */
    public ReadableExcel(File excelFile) {
        try {
            OPCPackage opcPackage = OPCPackage.open(excelFile);
            XSSFReader reader = new XSSFReader(opcPackage);
            this.sharedStringsTable = new ReadOnlySharedStringsTable(opcPackage);
            this.stylesTable = reader.getStylesTable();
            this.sheetIterator = (SheetIterator) reader.getSheetsData();
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 构建一个可读的 Excel 对象
     *
     * @param excelFilePath Excel 文件路径
     */
    public ReadableExcel(String excelFilePath) {
        this(new File(excelFilePath));
    }

    /**
     * 解析工作表
     *
     * @param index       工作表索引, 从1开始
     * @param targetClass 目标类型
     * @param <T>
     * @return 返回解析的结果列表
     */
    public <T> List<T> parseSheetAt(int index, Class<T> targetClass) {
        init(targetClass);
        int sheetCount = 1;
        while (true) {
            if (sheetIterator.hasNext()) {
                if (index == sheetCount) {
                    return processSheet(targetClass);
                } else {
                    ++sheetCount;
                    sheetIterator.next();
                }
            } else {
                break;
            }
        }
        throw new ReadExcelException("can not found sheet index : " + index);
    }

    /**
     * 解析所有的工作表
     *
     * @param targetClass 目标类型
     * @param <T>
     * @return 返回解析的结果列表
     */
    public <T> List<T> parseAllSheet(Class<T> targetClass) {
        init(targetClass);
        List<T> list = new LinkedList<>();
        list.addAll(processSheet(targetClass));
        return list;
    }

    /**
     * 设置解析的起始行, 从1开始
     *
     * @param startRow 起始行, 从1开始
     * @return 返回当前对象
     */
    public ReadableExcel setStartRow(int startRow) {
        this.startRow = startRow;
        return this;
    }

    // 初始化工作
    private void init(Class<?> targetClass) {
        this.cellFieldMap = new HashMap<>();
        this.beanDescriptor = new BeanDescriptor(targetClass);
        List<CellField> cellFields = AnnotationHandler.parseClass(targetClass);
        for (CellField cellField : cellFields) {
            cellFieldMap.put(cellField.getIndex(), cellField);
        }
    }

    // 解析工作表
    private void parseSheet(InputStream sheetInputStream, final List list) throws Throwable {
        XMLReader sheetParser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
        sheetParser.setContentHandler(new XSSFSheetHandler(stylesTable, sharedStringsTable) {
            Object item = beanDescriptor.newInstance();
            @Override
            public void postCellHandle(int index, String name, String value, int row, boolean newRow) {
                if (row >= startRow) {
                    if (newRow) {
                        if (!list.isEmpty() || startRow != row) {
                            list.add(item);
                        }
                        item = beanDescriptor.newInstance();
                    }
                    CellField cellField = cellFieldMap.get(index);
                    try {
                        Object cellValue = ValueConverter.convertObjectValue(value, cellField.getType());
                        beanDescriptor.setValueByName(cellField.getField(), cellValue);
                    } catch (Exception e) {
                        throw new ReadExcelException("Parse " + name + " error : " + e);
                    }
                }
            }

            @Override
            public void endDocument() throws SAXException {
                list.add(item);
            }
        });
        sheetParser.parse(new InputSource(sheetInputStream));
    }

    // 处理工作表
    private <T> List<T> processSheet(Class<T> targetClass) {
        InputStream stream = null;
        try {
            stream = sheetIterator.next();
            List<T> list = new ArrayList<>();
            parseSheet(stream, list);
            return list;
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        } finally {
            if (stream != null) {
                try {
                    stream.close();
                } catch (IOException e) {}
            }
        }
    }

}