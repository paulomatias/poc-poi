package com.paulo;

import java.io.File;
import java.io.IOException;

public class Main {
    // test this with -Xms16m -Xmx64m memory limitation and a lineNUmber and columnNuner of 500 to see OOM in XSSF and not in SXSSF.
    public static void main(String[] args) throws IOException {
        // cleaning previous files
        String xssfPath = "src/main/resources/xssf.xslx";
        String sxssfPath = "src/main/resources/sxssf.xslx";
        String templatePath = "src/main/resources/template.xslx";
        deleteFiles(xssfPath, sxssfPath, templatePath);

        // creates a fake csv
        int lineNumber = 500;
        int columnNumber = 500;
        CsvFile csvFile = new CsvFile(lineNumber, columnNumber);

        // create the excel file without stream
        XSSFWriter.createExcelFile(csvFile, xssfPath);

        // create the excel file with stream
        SXSSFWriter.createExcelFile(csvFile, sxssfPath);

        // create an excel templating file
        TemplateWriter.createTemplate(templatePath);
    }

    private static void deleteFiles(String... path) {
        for (String s : path) {
            new File(s).delete();
        }
    }
}
