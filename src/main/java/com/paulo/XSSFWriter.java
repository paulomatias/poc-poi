package com.paulo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class XSSFWriter {
    public static void createExcelFile(CsvFile csvFile, String path) {
        System.out.println("Creating Excel file without streaming!");
        XSSFWorkbook workbook = createWorkBook(csvFile);
        createFile(path, workbook);
        System.out.println("Done creating Excel file without streaming!");
        System.out.println();
    }

    private static XSSFWorkbook createWorkBook(CsvFile csvFile) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        for (int line = 0; line < csvFile.getLineNumber(); line++) {
            XSSFRow headerRow = sheet.createRow(line);
            for (int column = 0; column < csvFile.getColumnNumber(); column++) {
                Cell cell = headerRow.createCell(column);
                cell.setCellValue(csvFile.getArray()[line][column]);

                // Without streaming, all data is in the memory. You can always access any row at anytime while writing.
//                if (line == 8 && column == 4) {
//                    System.out.println("Can access the row 1, while writing cell " + csvFile.getArray()[line][column] + "! Exception will be thrown:");
//                    sheet.getRow(1).getLastCellNum();
//                }
            }
        }
        return workbook;
    }

    private static void createFile(String path, XSSFWorkbook workbook) {
        File file = new File(path);
        try (OutputStream outputStream = new FileOutputStream(file)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
