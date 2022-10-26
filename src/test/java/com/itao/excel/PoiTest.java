package com.itao.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.nio.file.Files;
import java.nio.file.Paths;

public class PoiTest {

    public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();

            CellStyle cellStyle = workbook.createCellStyle();

            // 合并单元格
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            sheet.setColumnWidth(0, (int) (256 * 10 + 184));
            row.setHeightInPoints(20.5f);
            cell.setCellValue("测试excel");
            cell.setCellStyle(cellStyle);
            workbook.write(Files.newOutputStream(Paths.get("D://1.xlsx")));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void m1() {
        try (Workbook workbook = new XSSFWorkbook()) {
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setFillForegroundColor(IndexedColors.GOLD.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setWrapText(true);
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontName("黑体");
            font.setColor(IndexedColors.GREEN.getIndex());
            font.setItalic(true);
            cellStyle.setFont(font);
            Sheet sheet = workbook.createSheet();
            // 列宽
            sheet.setColumnWidth(0, 30 * 256);
            // 合并单元格
            //sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 3));
            Row row = sheet.createRow(0);
            // 行高
            row.setHeightInPoints(100);
            Cell cell = row.createCell(0);
            cell.setCellStyle(cellStyle);
            cell.setCellValue("测试excel");
            workbook.write(Files.newOutputStream(Paths.get("D://1.xlsx")));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
