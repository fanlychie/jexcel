package org.fanlychie.jexcel.read;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.fanlychie.jexcel.exception.ExcelCastException;
import org.fanlychie.jexcel.exception.ReadExcelException;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * XSSF Sheet 处理器
 * Created by fanzyun on 2017/8/13.
 */
public abstract class XSSFSheetHandler extends DefaultHandler {

    private StylesTable stylesTable;

    private ReadOnlySharedStringsTable sharedStringsTable;

    private XSSFDataType nextDataType;

    private short formatIndex;

    private String formatString;

    private boolean nextIsRow;

    private int currCellIndex;

    private String currCellName;

    private String currCellValue;

    private DataFormatter formatter = new DataFormatter();

    private StringBuilder cellValueBuilder = new StringBuilder();

    enum XSSFDataType {BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER}

    public XSSFSheetHandler(StylesTable stylesTable, ReadOnlySharedStringsTable sharedStringsTable) {
        this.stylesTable = stylesTable;
        this.sharedStringsTable = sharedStringsTable;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        // v, is => inline str
        if ("v".equals(qName) || "is".equals(qName)) {
            cellValueBuilder.setLength(0);
        }
        // c => cell
        else if ("c".equals(qName)) {
            // r => name
            currCellName = attributes.getValue("r");
            currCellIndex = parseCellIndex(currCellName);
            // t => type
            String cellType = attributes.getValue("t");
            // s => style
            String cellStyleStr = attributes.getValue("s");
            formatIndex = -1;
            formatString = null;
            nextDataType = XSSFDataType.NUMBER;
            if ("b".equals(cellType)) {
                nextDataType = XSSFDataType.BOOL;
            } else if ("e".equals(cellType)) {
                nextDataType = XSSFDataType.ERROR;
            } else if ("inlineStr".equals(cellType)) {
                nextDataType = XSSFDataType.INLINESTR;
            } else if ("s".equals(cellType)) {
                nextDataType = XSSFDataType.SSTINDEX;
            } else if ("str".equals(cellType)) {
                nextDataType = XSSFDataType.FORMULA;
            } else if (cellStyleStr != null) {
                XSSFCellStyle style = stylesTable.getStyleAt(Integer.parseInt(cellStyleStr));
                formatIndex = style.getDataFormat();
                formatString = style.getDataFormatString();
                if (formatString == null) {
                    formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
                }
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        // v, is => contents of a cell
        if ("v".equals(qName) || "is".equals(qName)) {
            switch (nextDataType) {
                case BOOL:
                case FORMULA:
                    currCellValue = cellValueBuilder.toString();
                    break;
                case ERROR:
                    currCellValue = "\"ERROR:" + cellValueBuilder.toString() + '"';
                    break;
                case INLINESTR:
                    XSSFRichTextString rtsi = new XSSFRichTextString(cellValueBuilder.toString());
                    currCellValue = rtsi.toString();
                    break;
                case SSTINDEX:
                    try {
                        int n = Integer.parseInt(cellValueBuilder.toString());
                        XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(n));
                        currCellValue = rtss.toString();
                    } catch (NumberFormatException e) {
                        throw new ExcelCastException(e);
                    }
                    break;
                case NUMBER:
                    String n = cellValueBuilder.toString();
                    if (formatString != null) {
                        currCellValue = formatter.formatRawCellContents(Double.parseDouble(n), formatIndex, formatString);
                    } else {
                        currCellValue = n;
                    }
                    break;
                default:
                    throw new ReadExcelException("Undefined type: " + nextDataType);
            }
            postCellHandle(currCellIndex, currCellName, currCellValue, Integer.parseInt(currCellName.substring(1)), nextIsRow);
        }
        // row => new row
        else if ("row".equals(qName)) {
            nextIsRow = true;
        } else {
            nextIsRow = false;
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        cellValueBuilder.append(ch, start, length);
    }

    /**
     * 单元格处理
     *
     * @param index  单元格的索引
     * @param name   单元格的名称
     * @param value  单元格的字符串值
     * @param row    单元格的行号
     * @param newRow 是否是新的一行
     */
    public abstract void postCellHandle(int index, String name, String value, int row, boolean newRow);

    /**
     * 解析单元格的索引值
     *
     * @param cellName 单元格名称
     * @return 返回单元格的索引值
     */
    private int parseCellIndex(String cellName) {
        int index = -1;
        String cellNameLetter = null;
        int length = cellName.length();
        for (int i = 0; i < length; ++i) {
            if (Character.isDigit(cellName.charAt(i))) {
                cellNameLetter = cellName.substring(0, i);
                break;
            }
        }
        length = cellNameLetter.length();
        for (int i = 0; i < length; ++i) {
            int c = cellNameLetter.charAt(i);
            index = (index + 1) * 26 + c - 'A';
        }
        return index;
    }

}