package com.paulo;

import java.io.File;
import java.io.IOException;

public class Main {
    // test this with -Xms16m -Xmx64m memory limitation and a lineNUmber and columnNuner of 500 to see OOM in XSSF and not in SXSSF.
    public static void main(String[] args) throws IOException {
        // cleaning previous files
        String xssfPath = "src/main/resources/xssf.xslx";
        String sxssfPath = "src/main/resources/sxssf.xslx";
        deleteFiles(xssfPath, sxssfPath);

        // creates a fake csv
        int lineNumber = 500;
        int columnNumber = 500;
        CsvFile csvFile = new CsvFile(lineNumber, columnNumber);

        // create the excel file without stream
        XSSFWriter.createExcelFile(csvFile, xssfPath);

        // create the excel file with stream
        SXSSFWriter.createExcelFile(csvFile, sxssfPath);
    }

    private static void deleteFiles(String xssfPath, String sxssfPath) {
        File xssfFile = new File(xssfPath);
        File sxssfFile = new File(sxssfPath);
        xssfFile.delete();
        sxssfFile.delete();
    }
}
