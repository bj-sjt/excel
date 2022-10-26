package com.itao.excel.parse;


import com.itao.excel.bean.FontAttr;
import com.itao.excel.bean.StyleAttr;
import com.itao.excel.constant.*;
import com.itao.excel.util.StringUtil;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * dom4j解析
 */
public class Dom4jParser implements Parser {

    private Workbook workbook;
    private Sheet sheet;
    private Document document;
    private Element root;

    private Map<String, StyleAttr> styleMap;
    // 行索引
    private int rowIdx;
    // 列索引
    private int colIdx;

    @Override
    public void parse(String path) {
        parse(new File(path));
    }

    @Override
    public void parse(File xmlFile) {
        try {
            SAXReader saxReader = new SAXReader();
            document = saxReader.read(xmlFile);
            rootElement();
        } catch (DocumentException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void parse(Reader reader) {
        try {
            SAXReader saxReader = new SAXReader();
            document = saxReader.read(reader);
            rootElement();
        } catch (DocumentException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
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

    @Override
    public void write(File file) {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取根元素
     */
    private void rootElement() {
        root = document.getRootElement();
        String name = root.getName();
        if (!Tags.tables.equals(name)) {
            throw new RuntimeException("根标签不是" + Tags.tables);
        }
        parseStyle();
        parseTable();
    }

    /**
     * 解析Style标签
     */
    private void parseStyle() {
        Element styles = root.element(Tags.styles);
        List<Element> elements = styles.elements(Tags.style);
        if (CollectionUtils.isNotEmpty(elements)) {
            if (styleMap == null) {
                styleMap = new HashMap<>();
            }
            for (Element element : elements) {
                List<Attribute> attributes = element.attributes();
                parseAttr(attributes);
            }
        }
    }

    private void parseAttr(List<Attribute> attributes) {
        if (CollectionUtils.isNotEmpty(attributes)) {
            StyleAttr styleAttr = new StyleAttr();
            FontAttr fontAttr = new FontAttr();
            styleAttr.setFont(fontAttr);
            for (Attribute attribute : attributes) {
                if (attribute.getName().equals(Attrs.id)) {
                    styleAttr.setId(attribute.getValue());
                } else if (attribute.getName().equals(Attrs.border)) {
                    styleAttr.setBorder(Boolean.parseBoolean(attribute.getValue()));
                } else if (attribute.getName().equals(Attrs.wrapText)) {
                    styleAttr.setWrapText(Boolean.parseBoolean(attribute.getValue()));
                } else if (attribute.getName().equals(Attrs.hAlignment)) {
                    styleAttr.setHAlignment(attribute.getValue());
                } else if (attribute.getName().equals(Attrs.vAlignment)) {
                    styleAttr.setVAlignment(attribute.getValue());
                } else if (attribute.getName().equals(Attrs.foregroundColor)) {
                    styleAttr.setForegroundColor(attribute.getValue());
                } else if (attribute.getName().equals(Attrs.fontName)) {
                    fontAttr.setFontName(attribute.getValue());
                } else if (attribute.getName().equals(Attrs.color)) {
                    fontAttr.setColor(attribute.getValue());
                } else if (attribute.getName().equals(Attrs.size)) {
                    fontAttr.setSize(Short.parseShort(attribute.getValue()));
                } else if (attribute.getName().equals(Attrs.italic)) {
                    fontAttr.setItalic(Boolean.parseBoolean(attribute.getValue()));
                } else if (attribute.getName().equals(Attrs.bold)) {
                    fontAttr.setBold(Boolean.parseBoolean(attribute.getValue()));
                } else if (attribute.getName().equals(Attrs.strikeout)) {
                    fontAttr.setStrikeout(Boolean.parseBoolean(attribute.getValue()));
                }
            }
            if (StringUtil.isNotBlank(styleAttr.getId()))
                styleMap.put(styleAttr.getId(), styleAttr);
        }
    }

    /**
     * 解析table标签
     */
    private void parseTable() {
        List<Element> tables = root.elements(Tags.table);
        if (CollectionUtils.isEmpty(tables)) {
            return;
        }
        workbook = new XSSFWorkbook();
        int index = 1;
        for (Element table : tables) {
            rowIdx = 0;
            Attribute attribute = table.attribute(Attrs.name);
            if (attribute == null || StringUtil.isBlank(attribute.getValue())) {
                sheet = workbook.createSheet("sheet" + index++);
            } else{
                sheet = workbook.createSheet(attribute.getValue());
                index++;
            }
            parseTr(table);
        }
    }

    /**
     * 解析tr标签
     */
    private void parseTr(Element element) {
        List<Element> trs = element.elements(Tags.tr);
        if (CollectionUtils.isEmpty(trs)) {
            return;
        }
        for (Element tr : trs) {
            colIdx = 0;
            sheet.createRow(rowIdx);
            parseTd(tr);
            rowIdx++;
        }
    }
    /**
     * 解析td标签
     */
    private void parseTd(Element element) {
        List<Element> tds = element.elements(Tags.td);
        if (CollectionUtils.isEmpty(tds)) {
            return;
        }
        for (Element td : tds) {
            createCell(td);
            colIdx++;
        }
    }

    private void createCell(Element td) {
        boolean isMerge = false;
        CellStyle cellStyle = createCellStyle(td);
        Attribute styleIdAttr = td.attribute(Attrs.styleId);
        StyleAttr styleAttr = null;
        if (styleIdAttr != null && StringUtil.isNotBlank(styleIdAttr.getValue())) {
            styleAttr = styleMap.get(styleIdAttr.getValue());
        }
        Attribute colspanAttr = td.attribute(Attrs.colspan);
        Attribute rowspanAttr = td.attribute(Attrs.rowspan);
        Attribute widthAttr = td.attribute(Attrs.width);
        if (widthAttr != null && StringUtil.isNotBlank(widthAttr.getValue())) {
            float width = Float.parseFloat(widthAttr.getValue());
            sheet.setColumnWidth(colIdx, (int) (256 * width + 184));
        }
        int colspan = 0;
        int rowspan = 0;
        if (rowspanAttr != null && StringUtil.isNotBlank(rowspanAttr.getValue())) {
            rowspan = Integer.parseInt(rowspanAttr.getValue());
        }
        if (colspanAttr != null && StringUtil.isNotBlank(colspanAttr.getValue())) {
            colspan = Integer.parseInt(colspanAttr.getValue());
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
        Cell cell;
        Row row = sheet.getRow(rowIdx);
        if (isMerge) {
            cell = row.getCell(colIdx);
        } else {
            cell = row.createCell(colIdx);
        }
        rowIdx += rowspan;
        colIdx += colspan;
        cell.setCellStyle(cellStyle);
        String value = td.getStringValue();
        cell.setCellValue(value);
    }

    /**
     * 创建 CellStyle
     */
    private CellStyle createCellStyle(Element td) {
        CellStyle cellStyle = workbook.createCellStyle();
        Attribute styleIdAttr = td.attribute(Attrs.styleId);
        if (styleIdAttr != null && StringUtil.isNotBlank(styleIdAttr.getValue())) {
            StyleAttr styleAttr = styleMap.get(styleIdAttr.getValue());
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
        return cellStyle;
    }
    /**
     * 创建 Font
     */
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
}
