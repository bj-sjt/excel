package com.itao.excel.parse;

import com.itao.excel.bean.FontAttr;
import com.itao.excel.bean.StyleAttr;
import com.itao.excel.constant.*;
import com.itao.excel.util.StringUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

public class SaxHandler extends DefaultHandler {

    private Workbook workbook;
    private Sheet sheet;
    private Cell cell;
    private Map<String, StyleAttr> styleMap;
    // 行索引
    private int rowIdx;
    // 列索引
    private int colIdx;
    private int sheetIdx;

    private boolean isTd;

    private String content;
    @Override
    public void startDocument()  {
        workbook = new XSSFWorkbook();
        styleMap = new HashMap<>();
    }

    @Override
    public void endDocument()  {

    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {
        if (Tags.style.equals(qName)) {
            parseStyleAttr(attributes);
        } else if (Tags.table.equals(qName)) {
            String name = attributes.getValue(Attrs.name);
            rowIdx = 0;
            if (StringUtil.isBlank(name)) {
                sheet = workbook.createSheet("sheet" + ++sheetIdx);
            } else{
                sheet = workbook.createSheet(name);
                sheetIdx++;
            }
        } else if (Tags.tr.equals(qName)) {
            sheet.createRow(rowIdx);
        } else if (Tags.td.equals(qName)) {
            isTd = true;
            parseTdAttr(attributes);
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        if (isTd) {
            content = new String(ch, start, length);
            System.out.println(content);
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) {
        if (Tags.table.equals(qName)) {
            this.sheet = null;
            rowIdx = 0;
        } else if (Tags.tr.equals(qName)) {
            rowIdx++;
            colIdx = 0;
        } else if (Tags.td.equals(qName)) {
            isTd = false;
            cell.setCellValue(content);
            content = null;
            colIdx++;
        }
    }

    private void parseStyleAttr(Attributes attributes) {
        if (attributes != null) {
            StyleAttr styleAttr = new StyleAttr();
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.id))) {
                styleAttr.setId(attributes.getValue(Attrs.id));
            } else {
                return;
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.border))) {
                styleAttr.setBorder(Boolean.parseBoolean(attributes.getValue(Attrs.border)));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.wrapText))) {
                styleAttr.setWrapText(Boolean.parseBoolean(attributes.getValue(Attrs.wrapText)));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.hAlignment))) {
                styleAttr.setHAlignment(attributes.getValue(Attrs.hAlignment));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.vAlignment))) {
                styleAttr.setVAlignment(attributes.getValue(Attrs.vAlignment));
            }

            if (StringUtil.isNotBlank(attributes.getValue(Attrs.foregroundColor))) {
                styleAttr.setForegroundColor(attributes.getValue(Attrs.foregroundColor));
            }

            FontAttr fontAttr = new FontAttr();
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.fontName))) {
                fontAttr.setFontName(attributes.getValue(Attrs.fontName));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.color))) {
                fontAttr.setColor(attributes.getValue(Attrs.color));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.size))) {
                fontAttr.setSize(Short.parseShort(attributes.getValue(Attrs.size)));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.italic))) {
                fontAttr.setItalic(Boolean.parseBoolean(attributes.getValue(Attrs.italic)));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.bold))) {
                fontAttr.setBold(Boolean.parseBoolean(attributes.getValue(Attrs.bold)));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.strikeout))) {
                fontAttr.setStrikeout(Boolean.parseBoolean(attributes.getValue(Attrs.strikeout)));
            }
            styleAttr.setFont(fontAttr);
            styleMap.put(styleAttr.getId(), styleAttr);
        }
    }

    private void parseTdAttr(Attributes attributes) {
        if (attributes != null) {
            CellStyle cellStyle = workbook.createCellStyle();
            StyleAttr styleAttr = null;
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.styleId))) {
                styleAttr = styleMap.get(attributes.getValue(Attrs.styleId));
                if (styleAttr != null) {
                    if (StringUtil.isNotBlank(styleAttr.getForegroundColor())) {
                        short color = ColorEnum.getColor(styleAttr.getForegroundColor());
                        cellStyle.setFillForegroundColor(color);
                        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }
                    if (StringUtil.isNotBlank(styleAttr.getHAlignment())) {
                        int code = HorizontalAlignmentEnum.getCode(styleAttr.getHAlignment());
                        cellStyle.setAlignment(HorizontalAlignment.forInt(code));
                    }
                    if (StringUtil.isNotBlank(styleAttr.getVAlignment())) {
                        int code = VerticalAlignmentEnum.getCode(styleAttr.getVAlignment());
                        cellStyle.setVerticalAlignment(VerticalAlignment.forInt(code));
                    }
                    cellStyle.setWrapText(styleAttr.isWrapText());
                    FontAttr fontAttr = styleAttr.getFont();
                    if (fontAttr != null) {
                        cellStyle.setFont(createFont(fontAttr));
                    }
                }
            }

            boolean isMerge = false;
            int colspan = 0;
            int rowspan = 0;
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.colspan))) {
                colspan = Integer.parseInt(attributes.getValue(Attrs.colspan));
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.rowspan))) {
                rowspan = Integer.parseInt(attributes.getValue(Attrs.rowspan));
            }

            if (colspan > 1 || rowspan > 1) {
                // 合并单元格
                sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx + rowspan, colIdx, colIdx + colspan));
                if (styleAttr != null && styleAttr.isBorder()) {
                    cellStyle.setBorderBottom(BorderStyle.THIN);
                    cellStyle.setBorderLeft(BorderStyle.THIN);
                    cellStyle.setBorderRight(BorderStyle.THIN);
                    cellStyle.setBorderTop(BorderStyle.THIN);
                    Row tempRow;
                    Cell tempCell;
                    for (int i = rowIdx; i <= rowIdx + rowspan; i++ ) {
                        tempRow = sheet.createRow(i);
                        for (int j = colIdx; j <= colIdx + colspan; j++) {
                            tempCell = tempRow.createCell(j);
                            tempCell.setCellStyle(cellStyle);
                        }
                    }
                    isMerge = true;
                }
            }
            if (StringUtil.isNotBlank(attributes.getValue(Attrs.width))) {
                float width = Float.parseFloat(attributes.getValue(Attrs.width));
                sheet.setColumnWidth(colIdx, (int) (256 * width + 184));
            }
            Row row = sheet.getRow(rowIdx);
            if (isMerge) {
                cell = row.getCell(colIdx);
            } else {
                cell = row.createCell(colIdx);
            }
            rowIdx += rowspan;
            colIdx += colspan;
            cell.setCellStyle(cellStyle);
            cell.setCellValue(content);
        }
    }

    private Font createFont(FontAttr fontAttr) {
        Font font = workbook.createFont();
        if (StringUtil.isNotBlank(fontAttr.getFontName())) {
            font.setFontName(fontAttr.getFontName());
        }
        if (StringUtil.isNotBlank(fontAttr.getColor())) {
            short color = ColorEnum.getColor(fontAttr.getColor());
            font.setColor(color);
        }
        font.setItalic(fontAttr.isItalic());
        font.setBold(fontAttr.isBold());
        font.setStrikeout(fontAttr.isStrikeout());
        if (fontAttr.getSize() > 0) {
            font.setFontHeightInPoints(fontAttr.getSize());
        }
        return font;
    }
    public void write(OutputStream outputStream) {
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public void write(File file) {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
