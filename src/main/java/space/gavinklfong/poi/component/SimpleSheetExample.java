package space.gavinklfong.poi.component;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class SimpleSheetExample {

    public void generateSimpleWorkbook(String filename) {
        Workbook workbook = createSimpleWorkbook();
        saveWorkbook(workbook, filename);
    }

    public void generateWorkbookWithStyle(String filename) {
        Workbook workbook = createWorkbookWithStyle();
        saveWorkbook(workbook, filename);
    }

    private Workbook createSimpleWorkbook() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("main");
        Sheet sheet2 = workbook.createSheet("secondary");

        Row row = sheet1.createRow(2);
        Cell cell = row.createCell(3);
        cell.setCellValue("ABC");

        sheet1.createRow(5)
                .createCell(2)
                .setCellValue(12345);

        return workbook;
    }

    private Workbook createWorkbookWithStyle() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("main");

        Font numericFont = workbook.createFont();
        numericFont.setBold(true);
        numericFont.setColor(IndexedColors.BLUE.getIndex());

        CellStyle numericStyle = workbook.createCellStyle();
        numericStyle.setDataFormat(workbook.createDataFormat().
                getFormat("$##,##0.00"));
        numericStyle.setAlignment(HorizontalAlignment.CENTER);
        numericStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        numericStyle.setBorderBottom(BorderStyle.DOUBLE);
        numericStyle.setBottomBorderColor(IndexedColors.GREEN.getIndex());
        numericStyle.setFont(numericFont);

        Row numericRow = sheet1.createRow(5);
        numericRow.setHeightInPoints(50);

        sheet1.setColumnWidth(2, 20 * 256);
        Cell numericCell = numericRow.createCell(2);
        numericCell.setCellStyle(numericStyle);
        numericCell.setCellValue(12345);

        return workbook;
    }

    private void saveWorkbook(Workbook workbook, String filename) {
        try (OutputStream fileOut = new FileOutputStream(filename)) {
            workbook.write(fileOut);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
