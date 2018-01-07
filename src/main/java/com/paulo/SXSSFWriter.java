package com.paulo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class SXSSFWriter {

    public static void createExcelFile(CsvFile csvFile, String path) {
        System.out.println("Creating Excel file withstreaming!");
        SXSSFWorkbook workbook = createWorkbook(csvFile);
        createFile(path, workbook);
        System.out.println("Done creating Excel file with streaming!");
        System.out.println();
    }

    private static SXSSFWorkbook createWorkbook(CsvFile csvFile) {
        SXSSFWorkbook workbook = new SXSSFWorkbook(2); // flush window, must be chosen wisely
        SXSSFSheet sheet = workbook.createSheet();
        for (int line = 0; line < csvFile.getLineNumber(); line++) {
            SXSSFRow headerRow = sheet.createRow(line);
            for (int column = 0; column < csvFile.getColumnNumber(); column++) {
                Cell cell = headerRow.createCell(column);
                cell.setCellValue(csvFile.getArray()[line][column]);

                // When streaming, all data is flushed according to the window defined in workbook constructor
                // Uncomment this to see it happen!
                /*if (line == 8 && column == 4) {
                    System.out.println("Cannot access the row 1, while writing cell " + csvFile.getArray()[line][column] + "! Exception will be thrown:");
                    sheet.getRow(1).getLastCellNum();
                }*/
            }
        }
        return workbook;
    }

    static void createFile(String path, SXSSFWorkbook workbook) {
        File file = new File(path);
        try (OutputStream outputStream = new FileOutputStream(file)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.dispose(); // temporary files created by poi cleaned
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
